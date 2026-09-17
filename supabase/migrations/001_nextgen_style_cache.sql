-- NextGen Style Cache table
-- Stores style search results from NextGen to avoid re-querying on every extraction.
-- TTL: 24 hours (enforced by application code)
CREATE TABLE IF NOT EXISTS nextgen_style_cache (
    style TEXT PRIMARY KEY,
    info JSONB,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Index for fast lookups
CREATE INDEX IF NOT EXISTS idx_nextgen_style_cache_style ON nextgen_style_cache (style);
CREATE INDEX IF NOT EXISTS idx_nextgen_style_cache_updated ON nextgen_style_cache (updated_at);

-- Enable RLS (Row Level Security)
ALTER TABLE nextgen_style_cache ENABLE ROW LEVEL SECURITY;

-- Policy: allow service role full access
CREATE POLICY "Service role full access" ON nextgen_style_cache
    FOR ALL USING (auth.role() = 'service_role');

-- Comment
COMMENT ON TABLE nextgen_style_cache IS 'Cache for NextGen style search results to avoid repeated API calls';
