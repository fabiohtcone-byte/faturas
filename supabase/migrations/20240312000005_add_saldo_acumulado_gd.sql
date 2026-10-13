-- Migration: Adicionar coluna saldo_acumulado_gd na tabela bills
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS saldo_acumulado_gd TEXT;
