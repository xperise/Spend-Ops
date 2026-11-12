import { createClient } from '@supabase/supabase-js';

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

// Khoá anon có thể đưa lên trình duyệt. Dữ liệu được bảo vệ bằng chính sách RLS trong supabase/schema.sql.
export const SB = url && anonKey ? createClient(url, anonKey, { auth: { persistSession: true, autoRefreshToken: true } }) : null;
