-- Migration: Allow all authenticated users shared access to bills and energy_invoices
-- Description: Enables collaborative team access so that any authenticated employee in Sanesul can view and manage company bills and invoices.

-- 1. Drop old restrictive policies on 'bills'
DROP POLICY IF EXISTS "Users can view their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can insert their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can update their own bills" ON public.bills;
DROP POLICY IF EXISTS "Users can delete their own bills" ON public.bills;

-- 2. Create collaborative policies on 'bills' for authenticated users
CREATE POLICY "Authenticated users can view all bills"
ON public.bills FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "Authenticated users can insert bills"
ON public.bills FOR INSERT
TO authenticated
WITH CHECK (true);

CREATE POLICY "Authenticated users can update bills"
ON public.bills FOR UPDATE
TO authenticated
USING (true)
WITH CHECK (true);

CREATE POLICY "Authenticated users can delete bills"
ON public.bills FOR DELETE
TO authenticated
USING (true);

-- 3. Drop old restrictive policies on 'energy_invoices'
DROP POLICY IF EXISTS "Users can view their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can insert their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can update their own energy invoices" ON public.energy_invoices;
DROP POLICY IF EXISTS "Users can delete their own energy invoices" ON public.energy_invoices;

-- 4. Create collaborative policies on 'energy_invoices' for authenticated users
CREATE POLICY "Authenticated users can view all energy invoices"
ON public.energy_invoices FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "Authenticated users can insert energy invoices"
ON public.energy_invoices FOR INSERT
TO authenticated
WITH CHECK (true);

CREATE POLICY "Authenticated users can update energy invoices"
ON public.energy_invoices FOR UPDATE
TO authenticated
USING (true)
WITH CHECK (true);

CREATE POLICY "Authenticated users can delete energy invoices"
ON public.energy_invoices FOR DELETE
TO authenticated
USING (true);
