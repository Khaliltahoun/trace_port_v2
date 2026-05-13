CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

CREATE TYPE stop_status AS ENUM ('draft', 'pending', 'validated', 'rejected', 'closed');
CREATE TYPE report_type AS ENUM ('daily', 'monthly', 'annual');
CREATE TYPE log_action AS ENUM ('login', 'create', 'update', 'validate', 'reject', 'delete', 'export', 'status_change');

CREATE TABLE roles (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  code VARCHAR(40) UNIQUE NOT NULL,
  label VARCHAR(120) NOT NULL,
  permissions JSONB NOT NULL DEFAULT '[]',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE users (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  first_name VARCHAR(80) NOT NULL,
  last_name VARCHAR(80) NOT NULL,
  email VARCHAR(160) UNIQUE NOT NULL,
  password_hash TEXT NOT NULL,
  service VARCHAR(120),
  is_active BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE user_roles (
  user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  role_id UUID NOT NULL REFERENCES roles(id) ON DELETE CASCADE,
  PRIMARY KEY (user_id, role_id)
);

CREATE TABLE circuits (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  code VARCHAR(40) UNIQUE NOT NULL,
  label VARCHAR(160) NOT NULL,
  type VARCHAR(40) NOT NULL,
  is_active BOOLEAN NOT NULL DEFAULT true
);

CREATE TABLE equipments (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  code VARCHAR(60) UNIQUE NOT NULL,
  label VARCHAR(160) NOT NULL,
  circuit_id UUID REFERENCES circuits(id),
  parent_id UUID REFERENCES equipments(id),
  is_active BOOLEAN NOT NULL DEFAULT true
);

CREATE TABLE stop_types (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  code VARCHAR(60) UNIQUE NOT NULL,
  label VARCHAR(180) NOT NULL,
  category VARCHAR(80) NOT NULL,
  examples TEXT,
  is_active BOOLEAN NOT NULL DEFAULT true
);

CREATE TABLE stops (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  reference VARCHAR(80) UNIQUE NOT NULL,
  equipment_id UUID REFERENCES equipments(id),
  circuit_id UUID REFERENCES circuits(id),
  stop_type_id UUID REFERENCES stop_types(id),
  quality VARCHAR(80),
  assignment VARCHAR(160),
  destination VARCHAR(160),
  started_at TIMESTAMPTZ NOT NULL,
  ended_at TIMESTAMPTZ,
  duration_minutes INTEGER GENERATED ALWAYS AS (
    CASE WHEN ended_at IS NULL THEN NULL ELSE GREATEST(0, EXTRACT(EPOCH FROM (ended_at - started_at)) / 60)::INTEGER END
  ) STORED,
  comment TEXT,
  status stop_status NOT NULL DEFAULT 'pending',
  declared_by UUID REFERENCES users(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE validations (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  stop_id UUID NOT NULL REFERENCES stops(id) ON DELETE CASCADE,
  validator_id UUID NOT NULL REFERENCES users(id),
  decision stop_status NOT NULL,
  comment TEXT,
  decided_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE kpis (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  period_start DATE NOT NULL,
  period_end DATE NOT NULL,
  circuit_id UUID REFERENCES circuits(id),
  trs_global NUMERIC(8, 4),
  trs_exploitation NUMERIC(8, 4),
  trs_maintenance NUMERIC(8, 4),
  trg_global NUMERIC(8, 4),
  downtime_hours NUMERIC(12, 2),
  throughput_tph NUMERIC(12, 2),
  mttr_hours NUMERIC(12, 2),
  mtbf_hours NUMERIC(12, 2),
  computed_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE reports (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  type report_type NOT NULL,
  period_start DATE NOT NULL,
  period_end DATE NOT NULL,
  file_url TEXT,
  payload JSONB NOT NULL DEFAULT '{}',
  generated_by UUID REFERENCES users(id),
  generated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE logs (
  id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
  user_id UUID REFERENCES users(id),
  action log_action NOT NULL,
  object_type VARCHAR(80) NOT NULL,
  object_id UUID,
  detail TEXT,
  ip_address INET,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_stops_status ON stops(status);
CREATE INDEX idx_stops_started_at ON stops(started_at);
CREATE INDEX idx_stops_equipment ON stops(equipment_id);
CREATE INDEX idx_validations_stop ON validations(stop_id);
CREATE INDEX idx_logs_created_at ON logs(created_at);
