-- Migration: Initial Schema for Energy Invoices App
-- Description: Creates the necessary tables and Row Level Security (RLS) policies.

-- 1. Create the 'bills' table (used for PDF extraction)
CREATE TABLE IF NOT EXISTS public.bills (
    id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    file_name TEXT NOT NULL,
    uc TEXT,
    demanda_ponta_kw TEXT,
    demanda_fora_ponta_kw TEXT,
    demanda_potencia_medida_ponta TEXT,
    demanda_potencia_medida_fora_ponta TEXT,
    ano_leitura TEXT,
    mes_referencia TEXT,
    consumo_kwh_ponta TEXT,
    consumo_kwh_fora_ponta TEXT,
    valor_consumo_kwh_ponta TEXT,
    valor_consumo_kwh_fora_ponta TEXT,
    valor_total TEXT,
    cidade TEXT,
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
    status TEXT DEFAULT 'pending',
    error TEXT,
    user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE NOT NULL
);

-- 2. Create the 'energy_invoices' table (used for CSV uploads and dashboards)
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
    user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE NOT NULL
);

-- 3. Enable Row Level Security (RLS) on both tables
ALTER TABLE public.bills ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.energy_invoices ENABLE ROW LEVEL SECURITY;

-- 4. Create RLS Policies for 'bills'
-- Users can only see their own bills
CREATE POLICY "Users can view their own bills" 
ON public.bills FOR SELECT 
USING (auth.uid() = user_id);

-- Users can only insert their own bills
CREATE POLICY "Users can insert their own bills" 
ON public.bills FOR INSERT 
WITH CHECK (auth.uid() = user_id);

-- Users can update their own bills
CREATE POLICY "Users can update their own bills" 
ON public.bills FOR UPDATE 
USING (auth.uid() = user_id)
WITH CHECK (auth.uid() = user_id);

-- Users can delete their own bills
CREATE POLICY "Users can delete their own bills" 
ON public.bills FOR DELETE 
USING (auth.uid() = user_id);

-- 5. Create RLS Policies for 'energy_invoices'
-- Users can only see their own energy invoices
CREATE POLICY "Users can view their own energy invoices" 
ON public.energy_invoices FOR SELECT 
USING (auth.uid() = user_id);

-- Users can only insert their own energy invoices
CREATE POLICY "Users can insert their own energy invoices" 
ON public.energy_invoices FOR INSERT 
WITH CHECK (auth.uid() = user_id);

-- Users can update their own energy invoices
CREATE POLICY "Users can update their own energy invoices" 
ON public.energy_invoices FOR UPDATE 
USING (auth.uid() = user_id)
WITH CHECK (auth.uid() = user_id);

-- Users can delete their own energy invoices
CREATE POLICY "Users can delete their own energy invoices" 
ON public.energy_invoices FOR DELETE 
USING (auth.uid() = user_id);
