-- ====================================================================
-- SANESUL ENERGY - SCRIPT OFICIAL COMPLETO (PROJETO: yydvjgbfaapldtkhlqrh)
-- Execute este script no SQL Editor do Supabase para criar/atualizar todas as colunas
-- e liberar o acesso total para os usuários autenticados e anônimos da Sanesul.
-- ====================================================================

-- 1. CRIA OU ATUALIZA A TABELA 'bills'
CREATE TABLE IF NOT EXISTS public.bills (
    id TEXT PRIMARY KEY,
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

-- Garante que todas as colunas existam mesmo se a tabela já foi criada antes
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS file_name TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS uc TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_ponta_kw TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_fora_ponta_kw TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_potencia_medida_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_potencia_medida_fora_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS ano_leitura TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS mes_referencia TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS data_vencimento TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS consumo_kwh_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS consumo_kwh_fora_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS consumo_grupo_b TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_consumo_kwh_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_consumo_kwh_fora_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_total TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS cidade TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS gerencia TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS locin TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS mercado TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_todos_periodos_kw TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_potencia_nao_consumida_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_potencia_nao_consumida_f_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_potencia_ativa_ultrap_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS demanda_potencia_ativa_ultrap_f_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS energia_reativa_exced_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS energia_reativa_exced_f_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS energia_injetada_kwh TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS energia_compensada_kwh TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_demanda_potencia_medida_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_demanda_potencia_medida_fora_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_demanda_potencia_nao_consumida_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_demanda_potencia_nao_consumida_f_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_demanda_potencia_ativa_ultrap_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_demanda_potencia_ativa_ultrap_f_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_energia_reativa_exced_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_energia_reativa_exced_f_ponta TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS energia_atv_injetada_gdi_ouc TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_energia_atv_injetada_gdi_ouc TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS energia_atv_injetada_gdi_muc TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS valor_energia_atv_injetada_gdi_muc TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS cip TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS outros_encargos TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS pis TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS cofins TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS icms TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS concessionaria TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS numero_nota_fiscal TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS modalidade_tarifaria TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS subgrupo TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS tipo TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS status TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS error TEXT;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS user_id UUID;
ALTER TABLE public.bills ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now());

-- 2. CRIA OU ATUALIZA A TABELA 'uc_mappings'
CREATE TABLE IF NOT EXISTS public.uc_mappings (
    uc TEXT PRIMARY KEY,
    gerencia TEXT,
    locin TEXT,
    cidade TEXT,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    user_id UUID
);

-- 3. CRIA OU ATUALIZA A TABELA 'energy_invoices'
CREATE TABLE IF NOT EXISTS public.energy_invoices (
    id TEXT PRIMARY KEY,
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

-- 5. CONFIGURA POLÍTICAS DE ACESSO LIVRE PARA OS MEMBROS DA SANESUL (Anon e Authenticated)
DROP POLICY IF EXISTS "Allow all access to bills" ON public.bills;
DROP POLICY IF EXISTS "Authenticated users can view all bills" ON public.bills;
DROP POLICY IF EXISTS "Authenticated users can insert bills" ON public.bills;
DROP POLICY IF EXISTS "Authenticated users can update bills" ON public.bills;
DROP POLICY IF EXISTS "Authenticated users can delete bills" ON public.bills;
DROP POLICY IF EXISTS "Users can view their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can insert their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can update their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can delete their own bills" ON public.bills;

CREATE POLICY "Allow all access to bills" ON public.bills FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "Allow all access to uc_mappings" ON public.uc_mappings;
DROP POLICY IF EXISTS "Users can manage their own mappings" ON public.uc_mappings;
DROP POLICY IF EXISTS "Authenticated users can manage uc_mappings" ON public.uc_mappings;

CREATE POLICY "Allow all access to uc_mappings" ON public.uc_mappings FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "Allow all access to energy_invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Authenticated users can view all energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Authenticated users can insert energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Authenticated users can update energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Authenticated users can delete energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can view their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can insert their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can update their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can delete their own energy invoices" ON public.energy_invoices;

CREATE POLICY "Allow all access to energy_invoices" ON public.energy_invoices FOR ALL TO anon, authenticated USING (true) WITH CHECK (true);

-- 6. ÍNDICES DE ALTA PERFORMANCE
CREATE INDEX IF NOT EXISTS idx_bills_uc ON public.bills(uc);
CREATE INDEX IF NOT EXISTS idx_bills_ano_mes ON public.bills(ano_leitura, mes_referencia);
CREATE INDEX IF NOT EXISTS idx_bills_created_at ON public.bills(created_at DESC);
