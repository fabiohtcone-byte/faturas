-- Migration: Add data_vencimento to bills table
-- Description: Adds the data_vencimento column to the bills table.

ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS data_vencimento TEXT;
