import { createClient } from '@supabase/supabase-js';

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY;
const supabaseServiceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

const isProduction = process.env.NODE_ENV === 'production';

// Detect if we are in Mock Mode (missing or placeholder credentials)
export const isMock = (!supabaseUrl ||
    !supabaseAnonKey ||
    supabaseUrl.includes('[YOUR_') ||
    supabaseAnonKey.includes('[YOUR_')) && !isProduction;

if (isMock) {
    console.warn("⚠️ Database credentials missing. Running in LOCAL MOCK MODE.");
}

if (isProduction && (!supabaseUrl || !supabaseAnonKey)) {
    console.error("❌ CRITICAL ERROR: Supabase credentials missing in PRODUCTION environment!");
}

// Client-side: use anon key
export const supabase = isMock
    ? ({} as any)
    : createClient(supabaseUrl!, supabaseAnonKey!);

// Server-side: use service role key for admin/system actions
export const supabaseAdmin = isMock
    ? ({} as any)
    : createClient(supabaseUrl!, supabaseServiceKey || supabaseAnonKey!, {
        auth: {
            autoRefreshToken: false,
            persistSession: false
        }
    });
