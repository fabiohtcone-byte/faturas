import { createClient } from '@supabase/supabase-js';

// Use placeholders if environment variables are missing to prevent initialization crash.
// The user must configure these in the Settings -> Secrets menu in AI Studio.
const supabaseUrl = process.env.VITE_SUPABASE_URL || 'https://placeholder.supabase.co';
const supabaseAnonKey = process.env.VITE_SUPABASE_ANON_KEY || 'placeholder';

export const isSupabaseConfigured = Boolean(process.env.VITE_SUPABASE_URL && process.env.VITE_SUPABASE_ANON_KEY);

if (!isSupabaseConfigured) {
  console.warn(
    'AVISO: Variáveis de ambiente do Supabase não configuradas.\n' +
    'As funcionalidades de banco de dados e autenticação não funcionarão.\n' +
    'Por favor, configure VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY no menu de Configurações (Secrets) do AI Studio.'
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey, {
  auth: {
    persistSession: false
  }
});
