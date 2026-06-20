import "dotenv/config";
import dotenv from "dotenv";
import express from "express";
import bcrypt from "bcryptjs";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createSessionToken, requireAuth } from "./auth.js";
import { databaseStatus, describeDatabaseError, ensureSchema, getPool, query } from "./db.js";

dotenv.config({ path: ".env.local", override: false });

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const clientDistPath = path.resolve(__dirname, "..", "dist");
const app = express();
const port = Number(process.env.API_PORT || 8001);
const isProduction = process.env.NODE_ENV === "production";

app.use(express.json({ limit: "50mb" }));

app.get("/api/health", async (_req, res) => {
  const db = await ensureSchema();
  res.status(db.ok ? 200 : 503).json({
    ok: db.ok,
    database: databaseStatus(),
  });
});

app.post("/api/auth/signup", async (req, res) => {
  const name = String(req.body?.name || "").trim();
  const email = normalizeEmail(req.body?.email);
  const password = String(req.body?.password || "");

  if (name.length < 2) return res.status(400).json({ error: "Enter your full name." });
  if (!isEmail(email)) return res.status(400).json({ error: "Enter a valid email address." });
  if (password.length < 8 || !/\d/.test(password)) return res.status(400).json({ error: "Use at least 8 characters and one number." });

  try {
    const id = randomUUID();
    const passwordHash = await bcrypt.hash(password, 12);
    const result = await query(
      `INSERT INTO users (id, full_name, email, password_hash, last_login_at)
       VALUES ($1, $2, $3, $4, now())
       RETURNING id, full_name, email`,
      [id, name, email, passwordHash],
    );
    const user = result.rows[0];
    res.status(201).json({ user, token: createSessionToken(user) });
  } catch (error) {
    if (error.code === "23505") return res.status(409).json({ error: "An account already exists for this email." });
    sendDatabaseError(res, error);
  }
});

app.post("/api/auth/login", async (req, res) => {
  const email = normalizeEmail(req.body?.email);
  const password = String(req.body?.password || "");

  if (!isEmail(email) || password.length < 6) return res.status(400).json({ error: "Enter a valid email and password." });

  try {
    const result = await query("SELECT id, full_name, email, password_hash FROM users WHERE email = $1", [email]);
    const user = result.rows[0];
    if (!user || !(await bcrypt.compare(password, user.password_hash))) {
      return res.status(401).json({ error: "Email or password is incorrect." });
    }
    await query("UPDATE users SET last_login_at = now() WHERE id = $1", [user.id]);
    res.json({
      user: { id: user.id, full_name: user.full_name, email: user.email },
      token: createSessionToken(user),
    });
  } catch (error) {
    if (isOfflineLoginAllowed(email, password, error)) {
      const user = {
        id: process.env.OFFLINE_LOGIN_USER_ID || "00000000-0000-4000-8000-000000000001",
        full_name: process.env.OFFLINE_LOGIN_NAME || "Offline ParkSight User",
        email,
      };
      return res.json({
        user,
        token: createSessionToken(user),
        offline: true,
        warning: "Logged in with offline credentials because PostgreSQL is not reachable.",
      });
    }
    sendDatabaseError(res, error);
  }
});

app.get("/api/auth/me", requireAuth, async (req, res) => {
  try {
    const result = await query("SELECT id, full_name, email FROM users WHERE id = $1", [req.user.id]);
    const user = result.rows[0];
    if (!user) return res.status(404).json({ error: "User session was not found." });
    res.json({ user });
  } catch (error) {
    sendDatabaseError(res, error);
  }
});

app.post("/api/uploads/start", requireAuth, async (req, res) => {
  const fileName = String(req.body?.fileName || "").trim();
  if (!fileName) return res.status(400).json({ error: "Dataset file name is required." });

  try {
    const id = randomUUID();
    const result = await query(
      `INSERT INTO dataset_uploads (id, user_id, file_name, file_size, mime_type, status)
       VALUES ($1, $2, $3, $4, $5, 'processing')
       RETURNING id, status, created_at`,
      [id, req.user.id, fileName, numberOrNull(req.body?.fileSize), String(req.body?.mimeType || "") || null],
    );
    res.status(201).json({ upload: result.rows[0] });
  } catch (error) {
    sendDatabaseError(res, error);
  }
});

app.patch("/api/uploads/:id/results", requireAuth, async (req, res) => {
  const data = req.body?.data;
  if (!data?.summary || !Array.isArray(data.hotspots)) {
    return res.status(400).json({ error: "Upload results must include summary and hotspots." });
  }

  let client;
  try {
    const ready = await ensureSchema();
    if (!ready.ok) throw ready.error;
    client = await getPool().connect();
    await client.query("BEGIN");
    const ownUpload = await client.query("SELECT id FROM dataset_uploads WHERE id = $1 AND user_id = $2 FOR UPDATE", [req.params.id, req.user.id]);
    if (!ownUpload.rowCount) {
      await client.query("ROLLBACK");
      return res.status(404).json({ error: "Upload session was not found." });
    }

    await client.query("DELETE FROM upload_hotspots WHERE upload_id = $1", [req.params.id]);
    await client.query("DELETE FROM upload_stations WHERE upload_id = $1", [req.params.id]);
    await client.query("DELETE FROM upload_plates WHERE upload_id = $1", [req.params.id]);

    for (const hotspot of data.hotspots) {
      await client.query(
        `INSERT INTO upload_hotspots
         (upload_id, hotspot_id, rank, station, area, lat, lng, violations, impact_score, priority, payload)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)`,
        [
          req.params.id,
          String(hotspot.id),
          numberOrNull(hotspot.rank),
          hotspot.station || null,
          hotspot.area || null,
          numberOrNull(hotspot.lat),
          numberOrNull(hotspot.lng),
          numberOrNull(hotspot.violations),
          numberOrNull(hotspot.impactScore),
          hotspot.priority || null,
          JSON.stringify(hotspot),
        ],
      );
    }

    const hotspotCountByStation = new Map();
    for (const hotspot of data.hotspots) {
      const station = hotspot.station || "Unknown station";
      hotspotCountByStation.set(station, (hotspotCountByStation.get(station) || 0) + 1);
    }

    for (const station of data.stations || []) {
      const stationName = station.name || station.station || "Unknown station";
      await client.query(
        `INSERT INTO upload_stations
         (upload_id, station, violations, impact_score, hotspot_count, payload)
         VALUES ($1, $2, $3, $4, $5, $6)
         ON CONFLICT (upload_id, station) DO UPDATE SET
           violations = EXCLUDED.violations,
           impact_score = EXCLUDED.impact_score,
           hotspot_count = EXCLUDED.hotspot_count,
           payload = EXCLUDED.payload`,
        [
          req.params.id,
          stationName,
          firstNumber(station.violations, station.count, station.cases),
          numberOrNull(station.impactScore),
          firstNumber(station.hotspots, station.hotspotCount, hotspotCountByStation.get(stationName)),
          JSON.stringify(station),
        ],
      );
    }

    for (const plate of data.plates || []) {
      await client.query(
        `INSERT INTO upload_plates
         (upload_id, plate, count, vehicle, station, payload)
         VALUES ($1, $2, $3, $4, $5, $6)
         ON CONFLICT (upload_id, plate) DO UPDATE SET
           count = EXCLUDED.count,
           vehicle = EXCLUDED.vehicle,
           station = EXCLUDED.station,
           payload = EXCLUDED.payload`,
        [
          req.params.id,
          plate.plate || "UNKNOWN",
          numberOrNull(plate.count),
          plate.vehicle || null,
          plate.station || null,
          JSON.stringify(plate),
        ],
      );
    }

    await client.query(
      `UPDATE dataset_uploads
       SET status = 'completed',
           error_message = NULL,
           total_violations = $2,
           hotspot_count = $3,
           station_count = $4,
           plate_count = $5,
           summary = $6,
           result = $7,
           completed_at = now()
       WHERE id = $1`,
      [
        req.params.id,
        numberOrNull(data.summary.totalViolations),
        data.hotspots.length,
        data.stations?.length || 0,
        data.plates?.length || 0,
        JSON.stringify(data.summary),
        JSON.stringify(data),
      ],
    );
    await client.query("COMMIT");
    res.json({ ok: true, uploadId: req.params.id, hotspotsStored: data.hotspots.length });
  } catch (error) {
    if (client) await client.query("ROLLBACK").catch(() => {});
    sendDatabaseError(res, error);
  } finally {
    if (client) client.release();
  }
});

app.patch("/api/uploads/:id/fail", requireAuth, async (req, res) => {
  try {
    await query(
      `UPDATE dataset_uploads
       SET status = 'failed', error_message = $3, completed_at = now()
       WHERE id = $1 AND user_id = $2`,
      [req.params.id, req.user.id, String(req.body?.message || "Dataset analysis failed.")],
    );
    res.json({ ok: true });
  } catch (error) {
    sendDatabaseError(res, error);
  }
});

if (isProduction) {
  app.use(express.static(clientDistPath));
}

app.use("/api", (req, res) => {
  res.status(404).json({ error: `No API route for ${req.method} ${req.path}` });
});

if (isProduction) {
  app.get(/.*/, (_req, res) => {
    res.sendFile(path.join(clientDistPath, "index.html"));
  });
}

ensureSchema().then((status) => {
  if (!status.ok) {
    console.warn(`PostgreSQL is not ready: ${describeDatabaseError(status.error)}`);
    return;
  }
  ensureJudgeUser().catch((error) => {
    console.warn(`Judge login seed failed: ${describeDatabaseError(error)}`);
  });
});

app.listen(port, () => {
  console.log(`ParkSight API listening on http://localhost:${port}`);
});

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function numberOrNull(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function firstNumber(...values) {
  for (const value of values) {
    const number = numberOrNull(value);
    if (number !== null) return number;
  }
  return null;
}

function sendDatabaseError(res, error) {
  console.error(error);
  res.status(503).json({
    error: "Database is not reachable. Check PostgreSQL/pgAdmin connection settings and try again.",
    detail: process.env.NODE_ENV === "production" ? undefined : describeDatabaseError(error),
  });
}

function isOfflineLoginAllowed(email, password, error) {
  if (!isDatabaseConnectionError(error)) return false;
  const offlineEmail = normalizeEmail(process.env.OFFLINE_LOGIN_EMAIL || "officer@parksight.ai");
  const offlinePassword = process.env.OFFLINE_LOGIN_PASSWORD || "Password1";
  return email === offlineEmail && password === offlinePassword;
}

function isDatabaseConnectionError(error) {
  if (!error) return false;
  if (["ECONNREFUSED", "ENOTFOUND", "ETIMEDOUT", "EHOSTUNREACH"].includes(error.code)) return true;
  return Boolean(error.errors?.some((item) => ["ECONNREFUSED", "ENOTFOUND", "ETIMEDOUT", "EHOSTUNREACH"].includes(item.code)));
}

async function ensureJudgeUser() {
  const email = normalizeEmail(process.env.JUDGE_LOGIN_EMAIL || "judge@parksight.ai");
  const password = process.env.JUDGE_LOGIN_PASSWORD || "Judge@123";
  const name = String(process.env.JUDGE_LOGIN_NAME || "ParkSight Judge").trim();
  const passwordHash = await bcrypt.hash(password, 12);
  await query(
    `INSERT INTO users (id, full_name, email, password_hash)
     VALUES ($1, $2, $3, $4)
     ON CONFLICT (email) DO UPDATE SET
       full_name = EXCLUDED.full_name,
       password_hash = EXCLUDED.password_hash`,
    [process.env.JUDGE_LOGIN_USER_ID || "00000000-0000-4000-8000-000000000100", name, email, passwordHash],
  );
  console.log(`Judge login ready for ${email}`);
}
