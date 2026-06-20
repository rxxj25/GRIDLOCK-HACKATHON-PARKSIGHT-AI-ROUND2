CREATE TABLE IF NOT EXISTS users (
  id uuid PRIMARY KEY,
  full_name text NOT NULL,
  email text NOT NULL UNIQUE,
  password_hash text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  last_login_at timestamptz
);

CREATE TABLE IF NOT EXISTS dataset_uploads (
  id uuid PRIMARY KEY,
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  file_name text NOT NULL,
  file_size bigint,
  mime_type text,
  status text NOT NULL DEFAULT 'processing',
  error_message text,
  total_violations integer,
  hotspot_count integer,
  station_count integer,
  plate_count integer,
  summary jsonb,
  result jsonb,
  created_at timestamptz NOT NULL DEFAULT now(),
  completed_at timestamptz
);

CREATE INDEX IF NOT EXISTS dataset_uploads_user_created_idx
  ON dataset_uploads(user_id, created_at DESC);

CREATE TABLE IF NOT EXISTS upload_hotspots (
  upload_id uuid NOT NULL REFERENCES dataset_uploads(id) ON DELETE CASCADE,
  hotspot_id text NOT NULL,
  rank integer,
  station text,
  area text,
  lat double precision,
  lng double precision,
  violations integer,
  impact_score numeric,
  priority text,
  payload jsonb NOT NULL,
  PRIMARY KEY (upload_id, hotspot_id)
);

CREATE INDEX IF NOT EXISTS upload_hotspots_upload_rank_idx
  ON upload_hotspots(upload_id, rank);

CREATE TABLE IF NOT EXISTS upload_stations (
  upload_id uuid NOT NULL REFERENCES dataset_uploads(id) ON DELETE CASCADE,
  station text NOT NULL,
  violations integer,
  impact_score numeric,
  hotspot_count integer,
  payload jsonb NOT NULL,
  PRIMARY KEY (upload_id, station)
);

CREATE TABLE IF NOT EXISTS upload_plates (
  upload_id uuid NOT NULL REFERENCES dataset_uploads(id) ON DELETE CASCADE,
  plate text NOT NULL,
  count integer,
  vehicle text,
  station text,
  payload jsonb NOT NULL,
  PRIMARY KEY (upload_id, plate)
);
