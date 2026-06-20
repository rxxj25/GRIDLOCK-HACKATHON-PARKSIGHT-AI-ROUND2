import crypto from "node:crypto";

const TOKEN_TTL_MS = 1000 * 60 * 60 * 12;

function base64url(input) {
  return Buffer.from(input).toString("base64url");
}

function sign(value) {
  const secret = process.env.APP_SESSION_SECRET || "parksight-local-dev-secret";
  return crypto.createHmac("sha256", secret).update(value).digest("base64url");
}

export function createSessionToken(user) {
  const payload = {
    sub: user.id,
    email: user.email,
    exp: Date.now() + TOKEN_TTL_MS,
  };
  const body = base64url(JSON.stringify(payload));
  return `${body}.${sign(body)}`;
}

export function verifySessionToken(token) {
  if (!token || !token.includes(".")) return null;
  const [body, signature] = token.split(".");
  if (sign(body) !== signature) return null;
  try {
    const payload = JSON.parse(Buffer.from(body, "base64url").toString("utf8"));
    if (!payload.sub || !payload.exp || Date.now() > payload.exp) return null;
    return payload;
  } catch {
    return null;
  }
}

export function requireAuth(req, res, next) {
  const token = req.headers.authorization?.replace(/^Bearer\s+/i, "");
  const session = verifySessionToken(token);
  if (!session) {
    res.status(401).json({ error: "Please log in again before uploading a dataset." });
    return;
  }
  req.user = { id: session.sub, email: session.email };
  next();
}
