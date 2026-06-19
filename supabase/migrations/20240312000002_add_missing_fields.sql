-- Migration: Add missing fields
-- Description: Adds missing fields from UI and export to bills table

ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS consumo_grupo_b TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_todos_periodos_kw TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS mercado TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS gerencia TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS locin TEXT;
