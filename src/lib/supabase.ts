import { createClient } from '@supabase/supabase-js';

export const DEFAULT_SUPABASE_URL = 'https://yydvjgbfaapldtkhlqrh.supabase.co';
export const DEFAULT_SUPABASE_ANON_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl5ZHZqZ2JmYWFwbGR0a2hscXJoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMzMjQ1NTMsImV4cCI6MjA4ODkwMDU1M30.GqUoGviAYXveiEs7YmtN6SE5eZ3ZbiENaZtPUfy8blU';

// Always prioritize the active working Supabase project (yydvjgbfaapldtkhlqrh)
const rawUrl = import.meta.env.VITE_SUPABASE_URL;
const rawKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

const supabaseUrl = (rawUrl && rawUrl.includes('supabase.co') && !rawUrl.includes('placeholder'))
  ? rawUrl
  : DEFAULT_SUPABASE_URL;

const supabaseAnonKey = (rawKey && rawKey.length > 50 && !rawKey.includes('placeholder'))
  ? rawKey
  : DEFAULT_SUPABASE_ANON_KEY;

export const isSupabaseConfigured = true;

export const supabase = createClient(supabaseUrl, supabaseAnonKey, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
});
