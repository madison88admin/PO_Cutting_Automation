-- Learning Layer Tables for PO Cutting Automation
-- Run this in Supabase Studio SQL Editor.
-- The system works without these tables (in-memory cache),
-- but creating them enables persistent cross-session learning.

-- Enable required extensions
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- 1. NextGen Match Cache
-- Stores successful style+color → NextGen product matches
CREATE TABLE IF NOT EXISTS nextgen_match_cache (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    style_key TEXT NOT NULL,
    color_key TEXT NOT NULL,
    brand TEXT NOT NULL DEFAULT '',
    nextgen_product TEXT NOT NULL,
    nextgen_product_id BIGINT NOT NULL DEFAULT 0,
    nextgen_style_name TEXT NOT NULL DEFAULT '',
    nextgen_color_name TEXT NOT NULL DEFAULT '',
    nextgen_color_code TEXT NOT NULL DEFAULT '',
    unit_cost DECIMAL(12,4),
    currency TEXT,
    factory TEXT,
    costing_reference TEXT,
    match_score INT NOT NULL DEFAULT 0,
    match_reason TEXT NOT NULL DEFAULT '',
    hit_count INT NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(style_key, color_key, brand)
);
CREATE INDEX IF NOT EXISTS idx_nextgen_match_cache_lookup
    ON nextgen_match_cache(style_key, color_key, brand);

-- 2. User Correction Memory
-- Stores user overrides when the system matched the wrong product
CREATE TABLE IF NOT EXISTS nextgen_user_corrections (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    style_key TEXT NOT NULL,
    color_key TEXT NOT NULL,
    brand TEXT NOT NULL DEFAULT '',
    original_nextgen_product TEXT NOT NULL DEFAULT '',
    corrected_nextgen_product TEXT NOT NULL,
    corrected_nextgen_product_id BIGINT NOT NULL,
    corrected_color_name TEXT NOT NULL DEFAULT '',
    corrected_color_code TEXT NOT NULL DEFAULT '',
    corrected_by TEXT NOT NULL DEFAULT '',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_nextgen_user_corrections_lookup
    ON nextgen_user_corrections(style_key, color_key, brand);

-- 3. Header Mapping Cache
-- Stores successful header mappings per brand + file layout
CREATE TABLE IF NOT EXISTS header_mapping_cache (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    brand TEXT NOT NULL DEFAULT '',
    file_signature TEXT NOT NULL,
    raw_headers JSONB NOT NULL DEFAULT '[]',
    mapped_headers JSONB NOT NULL DEFAULT '{}',
    confidence INT NOT NULL DEFAULT 0,
    hit_count INT NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(brand, file_signature)
);
CREATE INDEX IF NOT EXISTS idx_header_mapping_cache_lookup
    ON header_mapping_cache(brand, file_signature);

-- 4. Color Mapping Cache
-- Stores successful color mappings per brand
CREATE TABLE IF NOT EXISTS color_mapping_cache (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    brand TEXT NOT NULL DEFAULT '',
    raw_color TEXT NOT NULL,
    canonical_color TEXT NOT NULL DEFAULT '',
    nextgen_color_name TEXT NOT NULL DEFAULT '',
    nextgen_color_code TEXT NOT NULL DEFAULT '',
    hit_count INT NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE(brand, raw_color)
);
CREATE INDEX IF NOT EXISTS idx_color_mapping_cache_lookup
    ON color_mapping_cache(brand, raw_color);

-- Enable Row Level Security (RLS)
ALTER TABLE nextgen_match_cache ENABLE ROW LEVEL SECURITY;
ALTER TABLE nextgen_user_corrections ENABLE ROW LEVEL SECURITY;
ALTER TABLE header_mapping_cache ENABLE ROW LEVEL SECURITY;
ALTER TABLE color_mapping_cache ENABLE ROW LEVEL SECURITY;

-- Allow service role full access (server-side only)
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'nextgen_match_cache' AND policyname = 'Service role full access') THEN
        CREATE POLICY "Service role full access" ON nextgen_match_cache
            FOR ALL USING (auth.role() = 'service_role');
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'nextgen_user_corrections' AND policyname = 'Service role full access') THEN
        CREATE POLICY "Service role full access" ON nextgen_user_corrections
            FOR ALL USING (auth.role() = 'service_role');
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'header_mapping_cache' AND policyname = 'Service role full access') THEN
        CREATE POLICY "Service role full access" ON header_mapping_cache
            FOR ALL USING (auth.role() = 'service_role');
    END IF;
    IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'color_mapping_cache' AND policyname = 'Service role full access') THEN
        CREATE POLICY "Service role full access" ON color_mapping_cache
            FOR ALL USING (auth.role() = 'service_role');
    END IF;
END $$;
