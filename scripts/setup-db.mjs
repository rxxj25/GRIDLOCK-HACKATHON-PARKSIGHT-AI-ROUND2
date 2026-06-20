import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";

dotenv.config({ path: ".env.local", override: false });
dotenv.config({ override: false });

const psqlCandidates = [
  process.env.PSQL_PATH,
  "D:\\Postgresql\\18\\bin\\psql.exe",
  "D:\\pgAdmin 4\\runtime\\psql.exe",
  "C:\\Program Files\\PostgreSQL\\18\\bin\\psql.exe",
  "psql",
].filter(Boolean);

const psql = psqlCandidates.find((candidate) => candidate === "psql" || fs.existsSync(candidate));
if (!psql) {
  console.error("Could not find psql.exe. Install PostgreSQL tools or set PSQL_PATH.");
  process.exit(1);
}

const databaseUrl = process.env.DATABASE_URL || "";
const parsed = databaseUrl ? new URL(databaseUrl) : null;
const host = process.env.PGHOST || parsed?.hostname || "localhost";
const port = process.env.PGPORT || parsed?.port || "5432";
const user = process.env.PGUSER || parsed?.username || "postgres";
const database = process.env.PGDATABASE || parsed?.pathname?.slice(1) || "parksight_ai";
const maintenanceDb = process.env.PGMAINTENANCE_DB || "postgres";
const password = process.env.PGPASSWORD || (parsed?.password ? decodeURIComponent(parsed.password) : "");
const psqlEnv = password ? { ...process.env, PGPASSWORD: password } : { ...process.env };

function run(args, options = {}) {
  const result = spawnSync(psql, args, {
    stdio: "inherit",
    shell: false,
    env: psqlEnv,
    ...options,
  });
  return result.status || 0;
}

const common = ["--host", host, "--port", port, "--username", user];
console.log(`Checking database "${database}" on ${host}:${port} as ${user}`);

const existsStatus = spawnSync(psql, [
  ...common,
  "--dbname",
  maintenanceDb,
  "--tuples-only",
  "--command",
  `SELECT 1 FROM pg_database WHERE datname = '${database.replaceAll("'", "''")}'`,
], { encoding: "utf8", env: psqlEnv });

if (existsStatus.status !== 0) process.exit(existsStatus.status || 1);
if (!existsStatus.stdout.includes("1")) {
  const createStatus = run([...common, "--dbname", maintenanceDb, "--command", `CREATE DATABASE ${quoteIdent(database)}`]);
  if (createStatus !== 0) process.exit(createStatus);
}

const schemaPath = path.resolve("server", "schema.sql");
process.exit(run([...common, "--dbname", database, "--file", schemaPath]));

function quoteIdent(value) {
  return `"${String(value).replaceAll('"', '""')}"`;
}
