-- Add brand config fields to mlo_mapping for shared templates/statuses
ALTER TABLE mlo_mapping
  ADD COLUMN IF NOT EXISTS orders_template TEXT,
  ADD COLUMN IF NOT EXISTS lines_template TEXT,
  ADD COLUMN IF NOT EXISTS valid_statuses TEXT[];

-- Optional: seed examples (remove if not needed)
-- UPDATE mlo_mapping
-- SET orders_template = 'BULK',
--     lines_template = 'BULK',
--     valid_statuses = ARRAY['Confirmed']
-- WHERE lower(brand) IN ('col', 'columbia');

-- UPDATE mlo_mapping
-- SET orders_template = 'BULK',
--     lines_template = 'BULK',
--     valid_statuses = ARRAY['Confirmed','Approved to Submit']
-- WHERE lower(brand) IN ('arcteryx', 'arc''teryx');
