-- Table for learned Buy File header templates
CREATE TABLE IF NOT EXISTS buy_file_templates (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    customer TEXT,
    headers TEXT[] NOT NULL,
    normalized_headers TEXT[] NOT NULL,
    mapping JSONB NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT now(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_buy_file_templates_customer ON buy_file_templates(customer);
CREATE INDEX IF NOT EXISTS idx_buy_file_templates_normalized_headers ON buy_file_templates USING GIN(normalized_headers);
