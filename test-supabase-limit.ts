import { createClient } from '@supabase/supabase-js';
const supabase = createClient(
  'https://yydvjgbfaapldtkhlqrh.supabase.co',
  'sb_publishable_1P6sirDLXMUXlvD_V9_aLQ_bqJ8cLGv'
);
async function test() {
  const { data, count, error } = await supabase.from('bills').select('id', { count: 'exact' });
  console.log("Count:", count);
  console.log("Data length:", data?.length);
}
test();
