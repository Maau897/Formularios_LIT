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
