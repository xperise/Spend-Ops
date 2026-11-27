// Vercel Serverless Function: cấp / thu hồi quyền đăng nhập ngay từ trang Người dùng, không cần vào Supabase.
// Chỉ Quản trị viên (đang hoạt động trong dữ liệu app) gọi được. Khoá service role chỉ nằm trên máy chủ.
//
// Biến môi trường trên Vercel (Project → Settings → Environment Variables), cài một lần:
//   SUPABASE_SERVICE_ROLE_KEY   Bắt buộc. Supabase → Project Settings → API Keys → service_role / secret key.
//   SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY  Tuỳ chọn, mặc định giống src/supabase.js.
//
// POST /api/users  body: { action: 'grant' | 'revoke' | 'reset', email, oldEmail? }
//   grant  : thêm email vào allowed_emails; nếu chưa có tài khoản đăng nhập thì tạo, trả về mật khẩu tạm.
//   reset  : đặt mật khẩu tạm mới cho tài khoản đã có (đảm bảo có trong allowed_emails).
//   revoke : gỡ email khỏi allowed_emails (người đó không đọc/ghi được dữ liệu nữa).
//   oldEmail (với grant): khi đổi email, gỡ email cũ.
import { createClient } from '@supabase/supabase-js';
import crypto from 'node:crypto';

const SUPABASE_URL = process.env.SUPABASE_URL || 'https://eovueumhcjxfptezqado.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_PUBLISHABLE_KEY || 'sb_publishable_EBJrnjH5Zj4V7wtTNG7aKQ_O5cySd1P';
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function tempPassword(){
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const b = crypto.randomBytes(12); let s = '';
  for (let i = 0; i < 12; i++) s += abc[b[i] % abc.length];
  return s.slice(0, 4) + '-' + s.slice(4, 8) + '-' + s.slice(8);
}
async function findUser(admin, email){
  for (let page = 1; page <= 20; page++) {
    const { data, error } = await admin.auth.admin.listUsers({ page, perPage: 200 });
    if (error) throw error;
    const u = data.users.find(x => (x.email || '').toLowerCase() === email);
    if (u) return u;
    if (data.users.length < 200) return null;
  }
  return null;
}

async function dropEmail(admin, email){   // xoá đúng email, không phân biệt hoa thường (không dùng ilike vì "_" là ký tự đại diện)
  const { data, error } = await admin.from('allowed_emails').select('email');
  if (error) throw error;
  const hit = (data || []).map(x => x.email).filter(e => String(e).toLowerCase() === email);
  if (hit.length) { const d = await admin.from('allowed_emails').delete().in('email', hit); if (d.error) throw d.error; }
}

export default async function handler(req, res){
  if (req.method !== 'POST') return res.status(405).json({ error: 'method' });
  const SERVICE = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!SERVICE) return res.status(503).json({ error: 'not_configured', message: 'Chưa cài SUPABASE_SERVICE_ROLE_KEY trên Vercel (xem README mục 6).' });
  const auth = req.headers.authorization || '';
  if (!auth.startsWith('Bearer ')) return res.status(401).json({ error: 'auth' });

  // 1) Người gọi là ai, có phải Quản trị đang hoạt động không (đọc dữ liệu bằng chính token của họ).
  const ur = await fetch(`${SUPABASE_URL}/auth/v1/user`, { headers: { apikey: SUPABASE_KEY, Authorization: auth } });
  if (!ur.ok) return res.status(401).json({ error: 'auth' });
  const caller = ((await ur.json()).email || '').toLowerCase();
  const sr = await fetch(`${SUPABASE_URL}/rest/v1/app_state?id=eq.main&select=data`, { headers: { apikey: SUPABASE_KEY, Authorization: auth } });
  const rows = sr.ok ? await sr.json() : []; const S = rows && rows[0] && rows[0].data;
  const me = S && (S.users || []).find(u => (u.email || '').toLowerCase() === caller && !u.deleted && u.active);
  if (!me || !me.admin) return res.status(403).json({ error: 'forbidden', message: 'Chỉ Quản trị viên được cấp quyền đăng nhập.' });

  const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
  const action = body.action; const email = String(body.email || '').trim().toLowerCase(); const oldEmail = String(body.oldEmail || '').trim().toLowerCase();
  if (!['grant','revoke','reset'].includes(action) || !EMAIL_RE.test(email)) return res.status(400).json({ error: 'bad_request' });
  if (action === 'revoke' && email === caller) return res.status(400).json({ error: 'self', message: 'Không tự thu hồi quyền của chính mình.' });

  const admin = createClient(SUPABASE_URL, SERVICE, { auth: { persistSession: false, autoRefreshToken: false } });
  try {
    if (action === 'revoke') {
      await dropEmail(admin, email);
      return res.status(200).json({ ok: true, revoked: email });
    }
    const up = await admin.from('allowed_emails').upsert({ email }, { onConflict: 'email' });
    if (up.error) throw up.error;
    if (action === 'grant' && oldEmail && oldEmail !== email && oldEmail !== caller) await dropEmail(admin, oldEmail);

    let u = await findUser(admin, email); let password = null, created = false;
    if (!u) {
      password = tempPassword();
      const { data, error } = await admin.auth.admin.createUser({ email, password, email_confirm: true });
      if (error) throw error; u = data.user; created = true;
    } else if (action === 'reset') {
      password = tempPassword();
      const { error } = await admin.auth.admin.updateUserById(u.id, { password, ban_duration: 'none' });
      if (error) throw error;
    }
    return res.status(200).json({ ok: true, email, created, password });
  } catch (e) {
    return res.status(500).json({ error: 'server', message: String(e.message || e).slice(0, 300) });
  }
}
