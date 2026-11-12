import { createClient } from '@supabase/supabase-js';

// Project URL và publishable key của dự án Supabase. Publishable key được phép đưa lên trình duyệt;
// dữ liệu được bảo vệ bằng RLS trong supabase/schema.sql.
const SUPABASE_URL = 'https://eovueumhcjxfptezqado.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_EBJrnjH5Zj4V7wtTNG7aKQ_O5cySd1P';

export const SB = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
  auth: { persistSession: true, autoRefreshToken: true },
});
