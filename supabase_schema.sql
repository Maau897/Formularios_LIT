create table if not exists monthly_periods (
    id uuid primary key default gen_random_uuid(),
    equipment_code text not null,
    month integer not null check (month between 1 and 12),
    year integer not null,
    correction_range_1 numeric(6,2) not null,
    correction_range_2 numeric(6,2) not null,
    correction_range_3 numeric(6,2) not null,
    observations text,
    reviewed_by text,
    reviewed_on date,
    created_at timestamptz not null default now(),
    unique (equipment_code, month, year)
);

create table if not exists non_working_days (
    id uuid primary key default gen_random_uuid(),
    period_id uuid not null references monthly_periods(id) on delete cascade,
    day integer not null check (day between 1 and 31),
    reason text,
    created_at timestamptz not null default now(),
    unique (period_id, day)
);

create table if not exists temperature_readings (
    id uuid primary key default gen_random_uuid(),
    period_id uuid not null references monthly_periods(id) on delete cascade,
    day integer not null check (day between 1 and 31),
    slot_index integer not null check (slot_index between 1 and 3),
    slot_label text not null,
    temperature_text text,
    performed_by text not null,
    verified_by text not null,
    recorded_on date not null,
    notes text,
    created_at timestamptz not null default now(),
    unique (period_id, day, slot_index)
);
