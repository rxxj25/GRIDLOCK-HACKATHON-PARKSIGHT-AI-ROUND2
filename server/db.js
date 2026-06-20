import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import pg from "pg";

const { Pool } = pg;
const __dirname = path.dirname(fileURLToPath(import.meta.url));

let pool;
let schemaReady = false;
let schemaError = null;

export function getPool() {
  if (!pool) {
    pool = new Pool({
      connectionString: process.env.DATABASE_URL,
      host: process.env.PGHOST,
      port: process.env.PGPORT ? Number(process.env.PGPORT) : undefined,
      database: process.env.PGDATABASE,
      user: process.env.PGUSER,
      password: process.env.PGPASSWORD,
      ssl: process.env.PGSSL === "true" ? { rejectUnauthorized: false } : undefined,
    });
  }
  return pool;
}

export async function ensureSchema() {
  if (schemaReady) return { ok: true };
  try {
    const sql = await fs.readFile(path.join(__dirname, "schema.sql"), "utf8");
    await getPool().query(sql);
    schemaReady = true;
    schemaError = null;
    return { ok: true };
  } catch (error) {
    schemaError = error;
    return { ok: false, error };
  }
}

export async function query(text, params) {
  const ready = await ensureSchema();
  if (!ready.ok) throw ready.error;
  return getPool().query(text, params);
}

export function databaseStatus() {
  return {
    ready: schemaReady,
    error: schemaError ? describeDatabaseError(schemaError) : null,
  };
}

export function describeDatabaseError(error) {
  if (!error) return null;
  if (error.errors?.length) {
    return error.errors.map((item) => item.message).filter(Boolean).join("; ") || error.message || String(error);
  }
  return error.message || String(error);
}
