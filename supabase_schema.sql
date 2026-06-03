create extension if not exists pgcrypto;

create table if not exists formularios_periodos (
    id uuid primary key default gen_random_uuid(),
    form_key text not null,
    equipment_code text not null,
    month integer not null check (month between 1 and 12),
    year integer not null check (year between 2000 and 2100),
    payload jsonb not null,
    created_by text,
    updated_by text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    unique (form_key, equipment_code, year, month)
);

create index if not exists idx_formularios_periodos_lookup
    on formularios_periodos (form_key, equipment_code, year, month);

create index if not exists idx_formularios_periodos_period
    on formularios_periodos (year desc, month desc);

create table if not exists formularios_auditoria (
    id_evento bigserial primary key,
    email text not null,
    accion text not null,
    detalle text,
    form_key text,
    equipment_code text,
    month integer,
    year integer,
    created_at timestamptz not null default now()
);

create index if not exists idx_formularios_auditoria_created_at
    on formularios_auditoria (created_at desc);

create index if not exists idx_formularios_auditoria_email
    on formularios_auditoria (email);

create table if not exists formularios_trazabilidad (
    id uuid primary key default gen_random_uuid(),
    form_key text not null,
    equipment_code text not null,
    entry_type text not null,
    status text not null default 'programado',
    scheduled_for date,
    completed_on date,
    provider text,
    notes text,
    created_by text,
    updated_by text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create index if not exists idx_formularios_trazabilidad_lookup
    on formularios_trazabilidad (form_key, equipment_code, entry_type);

create index if not exists idx_formularios_trazabilidad_schedule
    on formularios_trazabilidad (scheduled_for, status);
