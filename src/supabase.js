import { createClient } from '@supabase/supabase-js';

// Publishable key được phép đưa lên trình duyệt. Dữ liệu được bảo vệ bằng RLS trong supabase/schema.sql.
// Có thể ghi đè bằng biến môi trường VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY nếu cần.
const url = import.meta.env.VITE_SUPABASE_URL || 'https://eovueumhcjxfptezqado.supabase.co';
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY || 'sb_publishable_EBJrnjH5Zj4V7wtTNG7aKQ_O5cySd1P';

export const SB = createClient(url, anonKey, { auth: { persistSession: true, autoRefreshToken: true } });
