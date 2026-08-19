import { createClient } from '@supabase/supabase-js';

// HARDCODED — Active credentials for the SANESUL Supabase project.
// Embedded directly so Vercel always connects to the correct project,
// regardless of environment variable configuration.
const SUPABASE_URL = 'https://yydvjgbfaapldtkhlqrh.supabase.co';
const SUPABASE_ANON_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl5ZHZqZ2JmYWFwbGR0a2hscXJoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMzMjQ1NTMsImV4cCI6MjA4ODkwMDU1M30.GqUoGviAYXveiEs7YmtN6SE5eZ3ZbiENaZtPUfy8blU';

export const isSupabaseConfigured = true;

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: false,
    autoRefreshToken: false,
    detectSessionInUrl: false,
  },
});

