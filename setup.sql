-- Run this in Supabase → SQL Editor

CREATE TABLE clients (
    id             UUID    DEFAULT gen_random_uuid() PRIMARY KEY,
    name           TEXT    NOT NULL UNIQUE,
    account_number INTEGER NOT NULL,
    created_at     TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE account_columns (
    id             UUID    DEFAULT gen_random_uuid() PRIMARY KEY,
    column_name    TEXT    NOT NULL UNIQUE,
    account_number INTEGER NOT NULL,
    is_vat_exempt  BOOLEAN NOT NULL DEFAULT FALSE,
    created_at     TIMESTAMPTZ DEFAULT NOW()
);
