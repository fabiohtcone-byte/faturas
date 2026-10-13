import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL || 'https://yydvjgbfaapldtkhlqrh.supabase.co';
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || 'sb_publishable_1P6sirDLXMUXlvD_V9_aLQ_bqJ8cLGv';

export const cloudDb = createClient(supabaseUrl, supabaseAnonKey);
export const supabase = cloudDb;

export const isCloudConfigured = true;
export const isSupabaseConfigured = true;

let quotaExhausted = false;
export const isQuotaExhausted = () => quotaExhausted;
export const resetQuotaCircuitBreaker = () => { quotaExhausted = false; };
