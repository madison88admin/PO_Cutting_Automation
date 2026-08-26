// One-shot idempotent DB setup for the PO Cutting system.
// Creates ALL tables in the po_cutting schema.
// Safe to re-run: CREATE TABLE IF NOT EXISTS only.
import { readFileSync } from 'fs';

const env = readFileSync('.env.local', 'utf-8');
const get = (k: string) => (env.match(new RegExp(`^${k}=(.*)$`, 'm')) || [])[1]?.trim();
const url = get('NEXT_PUBLIC_SUPABASE_URL');
const key = get('SUPABASE_SERVICE_ROLE_KEY');
const schema = get('SUPABASE_DB_SCHEMA') || 'po_cutting';
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

const S = schema; // short alias

const STATEMENTS: { label: string; sql: string }[] = [
    // === Schema setup ===
    { label: `schema: ${S}`, sql: `CREATE SCHEMA IF NOT EXISTS "${S}"` },
    { label: 'ext: pgcrypto', sql: `CREATE EXTENSION IF NOT EXISTS "pgcrypto"` },

    // === Core tables ===
    { label: 'table: users', sql: `CREATE TABLE IF NOT EXISTS "${S}".users (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        name TEXT NOT NULL,
        email TEXT NOT NULL UNIQUE,
        role TEXT NOT NULL CHECK (role IN ('Admin','PBD Planner','Reviewer','IT Manager','Read-Only')),
        is_active BOOLEAN DEFAULT true,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW())` },

    { label: 'table: run_history', sql: `CREATE TABLE IF NOT EXISTS "${S}".run_history (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        user_id UUID REFERENCES "${S}".users(id),
        filename TEXT NOT NULL,
        status TEXT NOT NULL CHECK (status IN ('Processing','Validation Failed','Pending Review','Approved','Rejected')),
        error_count INTEGER DEFAULT 0,
        warning_count INTEGER DEFAULT 0,
        orders_rows INTEGER DEFAULT 0,
        lines_rows INTEGER DEFAULT 0,
        order_sizes_rows INTEGER DEFAULT 0,
        reviewed_by UUID REFERENCES "${S}".users(id),
        review_notes TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        completed_at TIMESTAMPTZ)` },

    { label: 'table: audit_logs', sql: `CREATE TABLE IF NOT EXISTS "${S}".audit_logs (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        event TEXT NOT NULL,
        user_id UUID REFERENCES "${S}".users(id),
        run_id UUID REFERENCES "${S}".run_history(id),
        metadata JSONB,
        ip_address TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW())` },

    // === Mapping tables ===
    { label: 'table: factory_mapping', sql: `CREATE TABLE IF NOT EXISTS "${S}".factory_mapping (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        brand TEXT NOT NULL,
        category TEXT NOT NULL,
        product_supplier TEXT NOT NULL,
        updated_by UUID REFERENCES "${S}".users(id),
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(brand, category))` },

    { label: 'table: column_mapping', sql: `CREATE TABLE IF NOT EXISTS "${S}".column_mapping (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        customer TEXT NOT NULL,
        buy_file_column TEXT NOT NULL,
        internal_field TEXT NOT NULL,
        notes TEXT,
        updated_by UUID REFERENCES "${S}".users(id),
        updated_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(customer, buy_file_column))` },

    { label: 'table: mlo_mapping', sql: `CREATE TABLE IF NOT EXISTS "${S}".mlo_mapping (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        brand TEXT NOT NULL UNIQUE,
        keyuser1 TEXT, keyuser2 TEXT, keyuser4 TEXT, keyuser5 TEXT,
        orders_template TEXT, lines_template TEXT,
        valid_statuses TEXT[],
        updated_by UUID REFERENCES "${S}".users(id),
        updated_at TIMESTAMPTZ DEFAULT NOW())` },

    // === Buy file templates ===
    { label: 'table: buy_file_templates', sql: `CREATE TABLE IF NOT EXISTS "${S}".buy_file_templates (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        customer TEXT,
        headers TEXT[] NOT NULL,
        normalized_headers TEXT[] NOT NULL,
        mapping JSONB NOT NULL,
        created_at TIMESTAMPTZ DEFAULT now(),
        updated_at TIMESTAMPTZ DEFAULT now())` },

    // === Learning layer ===
    { label: 'table: nextgen_match_cache', sql: `CREATE TABLE IF NOT EXISTS "${S}".nextgen_match_cache (
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

    { label: 'table: nextgen_user_corrections', sql: `CREATE TABLE IF NOT EXISTS "${S}".nextgen_user_corrections (
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

    { label: 'table: header_mapping_cache', sql: `CREATE TABLE IF NOT EXISTS "${S}".header_mapping_cache (
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

    { label: 'table: color_mapping_cache', sql: `CREATE TABLE IF NOT EXISTS "${S}".color_mapping_cache (
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

    // === Indexes ===
    { label: 'idx: users_email', sql: `CREATE INDEX IF NOT EXISTS idx_users_email ON "${S}".users(email)` },
    { label: 'idx: run_history_created_at', sql: `CREATE INDEX IF NOT EXISTS idx_run_history_created_at ON "${S}".run_history(created_at)` },
    { label: 'idx: audit_logs_created_at', sql: `CREATE INDEX IF NOT EXISTS idx_audit_logs_created_at ON "${S}".audit_logs(created_at)` },
    { label: 'idx: bft_customer', sql: `CREATE INDEX IF NOT EXISTS idx_bft_customer ON "${S}".buy_file_templates(customer)` },
    { label: 'idx: bft_norm_headers', sql: `CREATE INDEX IF NOT EXISTS idx_bft_norm_headers ON "${S}".buy_file_templates USING GIN(normalized_headers)` },
    { label: 'idx: nmc_lookup', sql: `CREATE INDEX IF NOT EXISTS idx_nmc_lookup ON "${S}".nextgen_match_cache(style_key, color_key, brand)` },
    { label: 'idx: nuc_lookup', sql: `CREATE INDEX IF NOT EXISTS idx_nuc_lookup ON "${S}".nextgen_user_corrections(style_key, color_key, brand)` },
    { label: 'idx: hmc_lookup', sql: `CREATE INDEX IF NOT EXISTS idx_hmc_lookup ON "${S}".header_mapping_cache(brand, file_signature)` },
    { label: 'idx: cmc_lookup', sql: `CREATE INDEX IF NOT EXISTS idx_cmc_lookup ON "${S}".color_mapping_cache(brand, raw_color)` },

    // === RLS ===
    { label: 'rls: nextgen_match_cache', sql: `ALTER TABLE "${S}".nextgen_match_cache ENABLE ROW LEVEL SECURITY` },
    { label: 'rls: nextgen_user_corrections', sql: `ALTER TABLE "${S}".nextgen_user_corrections ENABLE ROW LEVEL SECURITY` },
    { label: 'rls: header_mapping_cache', sql: `ALTER TABLE "${S}".header_mapping_cache ENABLE ROW LEVEL SECURITY` },
    { label: 'rls: color_mapping_cache', sql: `ALTER TABLE "${S}".color_mapping_cache ENABLE ROW LEVEL SECURITY` },
    { label: 'policy: nmc', sql: `DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'nextgen_match_cache' AND schemaname = '${S}' AND policyname = 'Service role full access') THEN CREATE POLICY "Service role full access" ON "${S}".nextgen_match_cache FOR ALL USING (auth.role() = 'service_role'); END IF; END $$` },
    { label: 'policy: nuc', sql: `DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'nextgen_user_corrections' AND schemaname = '${S}' AND policyname = 'Service role full access') THEN CREATE POLICY "Service role full access" ON "${S}".nextgen_user_corrections FOR ALL USING (auth.role() = 'service_role'); END IF; END $$` },
    { label: 'policy: hmc', sql: `DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'header_mapping_cache' AND schemaname = '${S}' AND policyname = 'Service role full access') THEN CREATE POLICY "Service role full access" ON "${S}".header_mapping_cache FOR ALL USING (auth.role() = 'service_role'); END IF; END $$` },
    { label: 'policy: cmc', sql: `DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename = 'color_mapping_cache' AND schemaname = '${S}' AND policyname = 'Service role full access') THEN CREATE POLICY "Service role full access" ON "${S}".color_mapping_cache FOR ALL USING (auth.role() = 'service_role'); END IF; END $$` },

    // === Seed admin ===
    { label: 'seed admin', sql: `INSERT INTO "${S}".users (name, email, role, is_active)
        VALUES ('Admin','admin@madison88.com','Admin',true)
        ON CONFLICT (email) DO NOTHING` },
];

async function main() {
    console.log(`Target schema: "${S}"`);

    // Probe: list existing tables in target schema
    const before = await execSql(`SELECT tablename FROM pg_tables WHERE schemaname='${S}' ORDER BY tablename`);
    const beforeNames: string[] = before.map((r: any) => r.tablename || r[0]);
    console.log(`Existing tables in "${S}":`, beforeNames.length ? beforeNames.join(', ') : '(none)');

    // Also check if tables exist in public schema (common mistake)
    const publicTables = await execSql(`SELECT tablename FROM pg_tables WHERE schemaname='public' AND tablename NOT LIKE 'pg_%' ORDER BY tablename`);
    const publicNames: string[] = publicTables.map((r: any) => r.tablename || r[0]);
    if (publicNames.length) {
        console.log(`\n⚠️  Tables also exist in "public" schema: ${publicNames.join(', ')}`);
        console.log(`   The app uses "${S}" schema. Public tables are not used.`);
    }

    let ok = 0, skipped = 0, failed = 0;
    for (const s of STATEMENTS) {
        try {
            await execSql(s.sql);
            ok++;
            console.log(`  OK   ${s.label}`);
        } catch (e) {
            const msg = String(e);
            if (/already exists/i.test(msg)) { skipped++; console.log(`  SKIP ${s.label} (already exists)`); }
            else { failed++; console.log(`  FAIL ${s.label}: ${msg.slice(0, 200)}`); }
        }
    }

    const after = await execSql(`SELECT tablename FROM pg_tables WHERE schemaname='${S}' ORDER BY tablename`);
    const afterNames: string[] = after.map((r: any) => r.tablename || r[0]);
    console.log(`\nFinal tables in "${S}":`, afterNames.join(', ') || '(none)');
    console.log(`\nDone. ok=${ok} skipped=${skipped} failed=${failed}`);
    process.exit(failed ? 1 : 0);
}

main().catch((e) => { console.error(e); process.exit(1); });
