-- Migration: Add uc_mappings table
-- Description: Creates the uc_mappings table in Supabase for synchronizing UC metadata (Gerência, LOCIN, Cidade).

CREATE TABLE IF NOT EXISTS public.uc_mappings (
    uc TEXT PRIMARY KEY,
    gerencia TEXT,
    locin TEXT,
    cidade TEXT,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE
);

-- Enable Row Level Security (RLS)
ALTER TABLE public.uc_mappings ENABLE ROW LEVEL SECURITY;

-- Create Policy to allow authenticated users to perform all operations
CREATE POLICY "Users can manage their own mappings" 
ON public.uc_mappings FOR ALL
TO authenticated
USING (true)
WITH CHECK (true);
