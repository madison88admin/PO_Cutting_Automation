// One-shot idempotent DB setup for the self-hosted Supabase instance.
// Creates any missing app tables (schema.sql + buy_file_templates +
// learning-cache tables). Safe to re-run: CREATE TABLE IF NOT EXISTS only.
import { readFileSync } from 'fs';

const env = readFileSync('.env.local', 'utf-8');
const get = (k: string) => (env.match(new RegExp(`^${k}=(.*)$`, 'm')) || [])[1]?.trim();
const url = get('NEXT_PUBLIC_SUPABASE_URL');
const key = get('SUPABASE_SERVICE_ROLE_KEY');
if (!url || !key) { console.error('Missing SUPABASE URL/key in .env.local'); process.exit(1); }

async function execSql(query: string): Promise<any> {
    const resp = await fetch(`${url}/pg/query`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', apikey: key!, Authorization: `Bearer ${key}` },
        body: JSON.stringify({ query }),
    });
    const text = await resp.text();
    if (!resp.ok) throw new Error(`SQL ${resp.status}: ${text.slice(0, 300)}`);
    return text ? JSON.parse(text) : null;
}

const STATEMENTS: { label: string; sql: string }[] = [
    { label: 'ext: pgcrypto/uuid-ossp', sql: `CREATE EXTENSION IF NOT EXISTS "pgcrypto"` },
    { label: 'table: users', sql: `CREATE TABLE IF NOT EXISTS users (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        name TEXT NOT NULL,
        email TEXT NOT NULL UNIQUE,
        role TEXT NOT NULL CHECK (role IN ('Admin','PBD Planner','Reviewer','IT Manager','Read-Only')),
        is_active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW())` },
    { label: 'table: factory_mapping', sql: `CREATE TABLE IF NOT EXISTS factory_mapping (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        brand TEXT NOT NULL,
        category TEXT NOT NULL,
        product_supplier TEXT NOT NULL,
        updated_by UUID REFERENCES users(id),
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(brand, category))` },
    { label: 'table: run_history', sql: `CREATE TABLE IF NOT EXISTS run_history (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        user_id UUID REFERENCES users(id),
        filename TEXT NOT NULL,
        status TEXT NOT NULL CHECK (status IN ('Processing','Validation Failed','Pending Review','Approved','Rejected')),
        error_count INTEGER DEFAULT 0,
        warning_count INTEGER DEFAULT 0,
        orders_rows INTEGER DEFAULT 0,
        lines_rows INTEGER DEFAULT 0,
        order_sizes_rows INTEGER DEFAULT 0,
        reviewed_by UUID REFERENCES users(id),
        review_notes TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        completed_at TIMESTAMPTZ)` },
    { label: 'table: audit_logs', sql: `CREATE TABLE IF NOT EXISTS audit_logs (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        event TEXT NOT NULL,
        user_id UUID REFERENCES users(id),
        run_id UUID REFERENCES run_history(id),
        metadata JSONB,
        ip_address TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW())` },
    { label: 'table: column_mapping', sql: `CREATE TABLE IF NOT EXISTS column_mapping (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        customer TEXT NOT NULL,
        buy_file_column TEXT NOT NULL,
        internal_field TEXT NOT NULL,
        notes TEXT,
        updated_by UUID REFERENCES users(id),
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(customer, buy_file_column))` },
    { label: 'table: mlo_mapping', sql: `CREATE TABLE IF NOT EXISTS mlo_mapping (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        brand TEXT NOT NULL UNIQUE,
        keyuser1 TEXT, keyuser2 TEXT, keyuser4 TEXT, keyuser5 TEXT,
        orders_template TEXT, lines_template TEXT,
        valid_statuses TEXT[],
        updated_by UUID REFERENCES users(id),
        updated_at TIMESTAMPTZ DEFAULT NOW())` },
    { label: 'table: buy_file_templates', sql: `CREATE TABLE IF NOT EXISTS buy_file_templates (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        customer TEXT,
        headers TEXT[] NOT NULL,
        normalized_headers TEXT[] NOT NULL,
        mapping JSONB NOT NULL,
        created_at TIMESTAMPTZ DEFAULT now(),
        updated_at TIMESTAMPTZ DEFAULT now())` },
    { label: 'idx: buy_file_templates', sql: `CREATE INDEX IF NOT EXISTS idx_bft_customer ON buy_file_templates(customer)` },
    { label: 'gin: buy_file_templates', sql: `CREATE INDEX IF NOT EXISTS idx_bft_norm_headers ON buy_file_templates USING GIN(normalized_headers)` },
    { label: 'table: header_mapping_cache', sql: `CREATE TABLE IF NOT EXISTS header_mapping_cache (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        brand TEXT NOT NULL DEFAULT '',
        file_signature TEXT NOT NULL,
        raw_headers JSONB NOT NULL DEFAULT '[]',
        mapped_headers JSONB NOT NULL DEFAULT '{}',
        confidence INT NOT NULL DEFAULT 0,
        hit_count INT NOT NULL DEFAULT 0,
        created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        UNIQUE(brand, file_signature))` },
    { label: 'table: color_mapping_cache', sql: `CREATE TABLE IF NOT EXISTS color_mapping_cache (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        brand TEXT NOT NULL DEFAULT '',
        raw_color TEXT NOT NULL,
        canonical_color TEXT NOT NULL DEFAULT '',
        nextgen_color_name TEXT NOT NULL DEFAULT '',
        nextgen_color_code TEXT NOT NULL DEFAULT '',
        hit_count INT NOT NULL DEFAULT 0,
        created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
        UNIQUE(brand, raw_color))` },
    { label: 'table: nextgen_match_cache', sql: `CREATE TABLE IF NOT EXISTS nextgen_match_cache (
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
        UNIQUE(style_key, color_key, brand))` },
    { label: 'table: nextgen_user_corrections', sql: `CREATE TABLE IF NOT EXISTS nextgen_user_corrections (
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
        created_at TIMESTAMPTZ NOT NULL DEFAULT NOW())` },
    { label: 'seed admin user', sql: `INSERT INTO users (name, email, role, is_active)
        VALUES ('Admin','admin@madison88.com','Admin',true)
        ON CONFLICT (email) DO NOTHING` },
];

async function main() {
    // connectivity probe + inventory
    const before = await execSql(`SELECT tablename FROM pg_tables WHERE schemaname='public' ORDER BY tablename`);
    const beforeNames: string[] = before.map((r: any) => r.tablename || r[0]);
    console.log('Existing public tables:', beforeNames.length ? beforeNames.join(', ') : '(none)');

    let ok = 0, skipped = 0, failed = 0;
    for (const s of STATEMENTS) {
        try {
            await execSql(s.sql);
            ok++;
            console.log(`  OK   ${s.label}`);
        } catch (e) {
            const msg = String(e);
            if (/already exists/i.test(msg)) { skipped++; console.log(`  SKIP ${s.label} (already exists)`); }
            else { failed++; console.log(`  FAIL ${s.label}: ${msg.slice(0, 160)}`); }
        }
    }

    const after = await execSql(`SELECT tablename FROM pg_tables WHERE schemaname='public' ORDER BY tablename`);
    const afterNames: string[] = after.map((r: any) => r.tablename || r[0]);
    console.log('\nFinal public tables:', afterNames.join(', ') || '(none)');
    console.log(`\nDone. ok=${ok} skipped=${skipped} failed=${failed}`);
    process.exit(failed ? 1 : 0);
}

main().catch((e) => { console.error(e); process.exit(1); });
