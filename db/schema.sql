-- PO Cutting Automation Database Schema
-- Run this in Supabase SQL Editor to set up the database

-- Enable UUID extension
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- Users table
CREATE TABLE IF NOT EXISTS users (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL,
    email TEXT NOT NULL UNIQUE,
    role TEXT NOT NULL CHECK (role IN ('Admin', 'PBD Planner', 'Reviewer', 'IT Manager', 'Read-Only')),
    is_active BOOLEAN DEFAULT true,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Factory mapping table
CREATE TABLE IF NOT EXISTS factory_mapping (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    brand TEXT NOT NULL,
    category TEXT NOT NULL,
    product_supplier TEXT NOT NULL,
    updated_by UUID REFERENCES users(id),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    UNIQUE(brand, category)
);

-- Run history table
CREATE TABLE IF NOT EXISTS run_history (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id UUID REFERENCES users(id),
    filename TEXT NOT NULL,
    status TEXT NOT NULL CHECK (status IN ('Processing', 'Validation Failed', 'Pending Review', 'Approved', 'Rejected')),
    error_count INTEGER DEFAULT 0,
    warning_count INTEGER DEFAULT 0,
    orders_rows INTEGER DEFAULT 0,
    lines_rows INTEGER DEFAULT 0,
    order_sizes_rows INTEGER DEFAULT 0,
    reviewed_by UUID REFERENCES users(id),
    review_notes TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    completed_at TIMESTAMP WITH TIME ZONE
);

-- Audit logs table
CREATE TABLE IF NOT EXISTS audit_logs (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    event TEXT NOT NULL,
    user_id UUID REFERENCES users(id),
    run_id UUID REFERENCES run_history(id),
    metadata JSONB,
    ip_address TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Column mapping table
CREATE TABLE IF NOT EXISTS column_mapping (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    customer TEXT NOT NULL,
    buy_file_column TEXT NOT NULL,
    internal_field TEXT NOT NULL,
    notes TEXT,
    updated_by UUID REFERENCES users(id),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    UNIQUE(customer, buy_file_column)
);

-- MLO mapping table
CREATE TABLE IF NOT EXISTS mlo_mapping (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    brand TEXT NOT NULL UNIQUE,
    keyuser1 TEXT,
    keyuser2 TEXT,
    keyuser4 TEXT,
    keyuser5 TEXT,
    orders_template TEXT,
    lines_template TEXT,
    valid_statuses TEXT[],
    updated_by UUID REFERENCES users(id),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Create indexes for better performance
CREATE INDEX IF NOT EXISTS idx_users_email ON users(email);
CREATE INDEX IF NOT EXISTS idx_users_role ON users(role);
CREATE INDEX IF NOT EXISTS idx_run_history_user_id ON run_history(user_id);
CREATE INDEX IF NOT EXISTS idx_run_history_status ON run_history(status);
CREATE INDEX IF NOT EXISTS idx_run_history_created_at ON run_history(created_at);
CREATE INDEX IF NOT EXISTS idx_audit_logs_user_id ON audit_logs(user_id);
CREATE INDEX IF NOT EXISTS idx_audit_logs_event ON audit_logs(event);
CREATE INDEX IF NOT EXISTS idx_audit_logs_created_at ON audit_logs(created_at);
CREATE INDEX IF NOT EXISTS idx_column_mapping_customer ON column_mapping(customer);
CREATE INDEX IF NOT EXISTS idx_factory_mapping_brand ON factory_mapping(brand);

-- Insert default admin user (change password in production)
INSERT INTO users (name, email, role, is_active)
VALUES ('Admin', 'admin@madison88.com', 'Admin', true)
ON CONFLICT (email) DO NOTHING;
