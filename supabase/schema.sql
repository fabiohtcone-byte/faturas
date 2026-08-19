-- ====================================================================
-- SANESUL ENERGY - SCRIPT 100% SEGURO PARA O SQL EDITOR DO SUPABASE
-- Este script cria todas as tabelas e permissões sem nenhum erro.
-- ====================================================================

-- 1. CRIA A TABELA 'bills' SE NÃO EXISTIR
CREATE TABLE IF NOT EXISTS public.bills (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()),
    file_name TEXT NOT NULL,
    uc TEXT,
    demanda_ponta_kw TEXT,
    demanda_fora_ponta_kw TEXT,
    demanda_potencia_medida_ponta TEXT,
    demanda_potencia_medida_fora_ponta TEXT,
    ano_leitura TEXT,
    mes_referencia TEXT,
    data_vencimento TEXT,
    consumo_kwh_ponta TEXT,
    consumo_kwh_fora_ponta TEXT,
    consumo_grupo_b TEXT,
    valor_consumo_kwh_ponta TEXT,
    valor_consumo_kwh_fora_ponta TEXT,
    valor_total TEXT,
    cidade TEXT,
    gerencia TEXT,
    locin TEXT,
    mercado TEXT,
    demanda_todos_periodos_kw TEXT,
    demanda_potencia_nao_consumida_ponta TEXT,
    demanda_potencia_nao_consumida_f_ponta TEXT,
    demanda_potencia_ativa_ultrap_ponta TEXT,
    demanda_potencia_ativa_ultrap_f_ponta TEXT,
    energia_reativa_exced_ponta TEXT,
    energia_reativa_exced_f_ponta TEXT,
    energia_injetada_kwh TEXT,
    energia_compensada_kwh TEXT,
    valor_demanda_potencia_medida_ponta TEXT,
    valor_demanda_potencia_medida_fora_ponta TEXT,
    valor_demanda_potencia_nao_consumida_ponta TEXT,
    valor_demanda_potencia_nao_consumida_f_ponta TEXT,
    valor_demanda_potencia_ativa_ultrap_ponta TEXT,
    valor_demanda_potencia_ativa_ultrap_f_ponta TEXT,
    valor_energia_reativa_exced_ponta TEXT,
    valor_energia_reativa_exced_f_ponta TEXT,
    energia_atv_injetada_gdi_ouc TEXT,
    valor_energia_atv_injetada_gdi_ouc TEXT,
    energia_atv_injetada_gdi_muc TEXT,
    valor_energia_atv_injetada_gdi_muc TEXT,
    cip TEXT,
    outros_encargos TEXT,
    pis TEXT,
    cofins TEXT,
    icms TEXT,
    concessionaria TEXT,
    numero_nota_fiscal TEXT,
    modalidade_tarifaria TEXT,
    subgrupo TEXT,
    tipo TEXT,
    status TEXT DEFAULT 'completed',
    error TEXT,
    user_id UUID
);

-- Adiciona colunas extras caso a tabela já existisse em versão anterior
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS data_vencimento TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS consumo_grupo_b TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_todos_periodos_kw TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS mercado TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS gerencia TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS locin TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now());

-- 2. CRIA A TABELA 'uc_mappings' SE NÃO EXISTIR
CREATE TABLE IF NOT EXISTS public.uc_mappings (
    uc TEXT PRIMARY KEY,
    gerencia TEXT,
    locin TEXT,
    cidade TEXT,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    user_id UUID
);

-- 3. CRIA A TABELA 'energy_invoices' SE NÃO EXISTIR
CREATE TABLE IF NOT EXISTS public.energy_invoices (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    uc TEXT NOT NULL,
    ano TEXT NOT NULL,
    mes TEXT NOT NULL,
    consumo_ponta NUMERIC,
    valor_ponta NUMERIC,
    consumo_fora_ponta NUMERIC,
    valor_fora_ponta NUMERIC,
    valor_total NUMERIC,
    cidade TEXT,
    user_id UUID
);

-- 4. HABILITA ROW LEVEL SECURITY (RLS)
ALTER TABLE public.bills ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.uc_mappings ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.energy_invoices ENABLE ROW LEVEL SECURITY;

-- 5. LIMPA E RECRIA AS POLÍTICAS DE ACESSO COMPARTILHADO (Sem risco de erro 42P01)
DROP POLICY IF EXISTS "Authenticated users can view all bills" ON public.bills;
DROP POLICY IF EXISTS "Authenticated users can insert bills" ON public.bills;
DROP POLICY IF EXISTS "Authenticated users can update bills" ON public.bills;
DROP POLICY IF EXISTS "Authenticated users can delete bills" ON public.bills;
DROP POLICY IF EXISTS "Users can view their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can insert their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can update their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can delete their own bills" ON public.bills;

CREATE POLICY "Authenticated users can view all bills" ON public.bills FOR SELECT TO authenticated USING (true);
CREATE POLICY "Authenticated users can insert bills" ON public.bills FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "Authenticated users can update bills" ON public.bills FOR UPDATE TO authenticated USING (true) WITH CHECK (true);
CREATE POLICY "Authenticated users can delete bills" ON public.bills FOR DELETE TO authenticated USING (true);

DROP POLICY IF EXISTS "Users can manage their own mappings" ON public.uc_mappings;
DROP POLICY IF EXISTS "Authenticated users can manage uc_mappings" ON public.uc_mappings;

CREATE POLICY "Authenticated users can manage uc_mappings" ON public.uc_mappings FOR ALL TO authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "Authenticated users can view all energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Authenticated users can insert energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Authenticated users can update energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Authenticated users can delete energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can view their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can insert their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can update their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can delete their own energy invoices" ON public.energy_invoices;

CREATE POLICY "Authenticated users can view all energy invoices" ON public.energy_invoices FOR SELECT TO authenticated USING (true);
CREATE POLICY "Authenticated users can insert energy invoices" ON public.energy_invoices FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "Authenticated users can update energy invoices" ON public.energy_invoices FOR UPDATE TO authenticated USING (true) WITH CHECK (true);
CREATE POLICY "Authenticated users can delete energy invoices" ON public.energy_invoices FOR DELETE TO authenticated USING (true);

-- 6. ÍNDICES DE PERFORMANCE
CREATE INDEX IF NOT EXISTS idx_bills_uc ON public.bills(uc);
CREATE INDEX IF NOT EXISTS idx_bills_ano_mes ON public.bills(ano_leitura, mes_referencia);
CREATE INDEX IF NOT EXISTS idx_bills_created_at ON public.bills(created_at DESC);
