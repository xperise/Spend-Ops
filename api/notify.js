// Vercel Serverless Function: gửi email báo người duyệt khi có đề xuất chi / đợt chi chờ duyệt.
// Bảo mật: chỉ người đã đăng nhập và có trong allowed_emails mới gọi được (dữ liệu đọc bằng chính token của họ, RLS chặn người khác).
// Người nhận và nội dung email do máy chủ tự tính từ dữ liệu đã lưu, client chỉ gửi mã phiếu/đợt.
//
// Biến môi trường trên Vercel (Project → Settings → Environment Variables):
//   MAIL_FROM            Bắt buộc. VD: "Xperise Spend Ops <spendops@xperise.com>"
//   Cách 1 — Resend:     RESEND_API_KEY
//   Cách 2 — SMTP:       SMTP_HOST, SMTP_PORT (465 hoặc 587), SMTP_USER, SMTP_PASS
//                        (Google Workspace: smtp.gmail.com, 465, email gửi, App Password 16 ký tự)
//   APP_URL              Tuỳ chọn. Link mở ứng dụng trong email, mặc định lấy theo domain đang chạy.
//   SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY  Tuỳ chọn, mặc định giống src/supabase.js.

const SUPABASE_URL = process.env.SUPABASE_URL || 'https://eovueumhcjxfptezqado.supabase.co';
const SUPABASE_KEY = process.env.SUPABASE_PUBLISHABLE_KEY || 'sb_publishable_EBJrnjH5Zj4V7wtTNG7aKQ_O5cySd1P';
const TZ = 'Asia/Ho_Chi_Minh';

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => Math.round(n || 0).toLocaleString('vi-VN');
const today = () => new Intl.DateTimeFormat('en-CA', {timeZone: TZ}).format(new Date()); // YYYY-MM-DD
const dvn = d => d ? d.split('-').reverse().join('/') : '';

function delegActive(g){ if (g.revokedAt) return false; const d = today(); if (g.start && d < g.start) return false; if (g.end && d > g.end) return false; return true; }

// Người cần nhận: người duyệt bước hiện tại chưa quyết định + người đang được họ uỷ quyền.
function recipients(S, step, exclude){
  const users = S.users || [];
  const ok = id => { const u = users.find(x => x.id === id); return u && u.active && !u.deleted && u.email ? u : null; };
  const pending = (step.approvers || []).filter(id => !(step.decisions || {})[id] && id !== exclude);
  const out = new Map();
  pending.forEach(id => { const u = ok(id); if (u) out.set(u.id, {u, forWho: null}); });
  (S.delegations || []).filter(g => pending.includes(g.from) && delegActive(g) && g.to !== exclude).forEach(g => {
    const u = ok(g.to); const p = users.find(x => x.id === g.from);
    if (u && !out.has(u.id)) out.set(u.id, {u, forWho: p ? p.name : null});
  });
  return [...out.values()];
}

function layout(title, rows, link, button){
  return `<!doctype html><html><body style="margin:0;background:#f4f5f7;font-family:Arial,Helvetica,sans-serif;color:#1f2328">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="padding:24px 12px"><tr><td align="center">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;background:#ffffff;border:1px solid #e3e6ea;border-radius:12px">
    <tr><td style="padding:22px 24px 6px;font-size:13px;color:#59636e">Xperise Spend Ops</td></tr>
    <tr><td style="padding:0 24px 14px;font-size:18px;font-weight:bold;line-height:1.4">${esc(title)}</td></tr>
    <tr><td style="padding:0 24px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-size:14px;line-height:1.5">
      ${rows.filter(r => r[1]).map(r => `<tr><td style="padding:6px 0;color:#59636e;width:150px;vertical-align:top">${esc(r[0])}</td><td style="padding:6px 0;font-weight:${r[2] ? 'bold' : 'normal'}">${esc(r[1])}</td></tr>`).join('')}
    </table></td></tr>
    <tr><td style="padding:20px 24px 24px"><a href="${esc(link)}" style="display:inline-block;background:#0f9d8a;color:#ffffff;text-decoration:none;padding:11px 20px;border-radius:8px;font-weight:bold;font-size:14px">${esc(button)}</a>
      <div style="font-size:12px;color:#8b949e;margin-top:14px">Email tự động, không trả lời email này. Mở ứng dụng để duyệt, trả lại hoặc từ chối.</div></td></tr>
  </table></td></tr></table></body></html>`;
}

function buildReq(S, r, appUrl){
  if (r.status !== 'pending') return null;
  const rd = (r.rounds || [])[r.rounds.length - 1]; const step = rd && rd.chain.find(s => s.status === 'current');
  if (!step) return null;
  const who = (S.users || []).find(u => u.id === r.requester) || {};
  const dept = (S.depts || []).find(d => d.id === r.dept) || {};
  const to = recipients(S, step, S.config && S.config.blockSelf ? r.requester : null);
  const link = `${appUrl}/?open=${encodeURIComponent(r.id)}`;
  return to.map(({u, forWho}) => ({
    to: u.email,
    subject: `[Spend Ops] ${r.code} chờ bạn duyệt: ${r.title} · ${fmt(r.amount)} đ`,
    html: layout(`${forWho ? `Duyệt thay ${forWho}: ` : ''}${r.code} chờ bạn duyệt ${String(step.label || '').toLowerCase()}`, [
      ['Nội dung', r.title, true], ['Số tiền', `${fmt(r.amount)} đ`, true], ['Người đề xuất', who.name],
      ['Phòng chịu chi phí', dept.name], ['Hạng mục', r.cat], ['Cần chi trước', dvn(r.needBy)],
      ['Lần gửi', rd.no > 1 ? `Gửi lại lần ${rd.no}` : ''], ['Ghi chú', r.note]
    ], link, 'Mở đề xuất để duyệt')
  }));
}

function buildBatch(S, b, appUrl){
  if (b.status !== 'review') return null;
  const step = (b.chain || []).find(s => s.status === 'current'); if (!step) return null;
  const reqs = (b.reqIds || []).map(id => (S.reqs || []).find(r => r.id === id)).filter(Boolean);
  const total = reqs.reduce((a, r) => a + (r.amount || 0), 0);
  const by = (S.users || []).find(u => u.id === b.createdBy) || {};
  const to = recipients(S, step, b.createdBy);
  const link = `${appUrl}/?batch=${encodeURIComponent(b.id)}`;
  return to.map(({u, forWho}) => ({
    to: u.email,
    subject: `[Spend Ops] Đợt chi ${b.code} chờ bạn duyệt: ${reqs.length} khoản · ${fmt(total)} đ`,
    html: layout(`${forWho ? `Duyệt thay ${forWho}: ` : ''}Đợt chi ${b.code} chờ bạn duyệt ${String(step.label || '').toLowerCase()}`, [
      ['Tổng tiền', `${fmt(total)} đ`, true], ['Số khoản', String(reqs.length)], ['Hẹn thanh toán', dvn(b.payDate)], ['Kế toán lập', by.name],
      ['Các khoản', reqs.slice(0, 8).map(r => `${r.code} ${r.title} (${fmt(r.amount)} đ)`).join('; ') + (reqs.length > 8 ? `; và ${reqs.length - 8} khoản khác` : '')]
    ], link, 'Mở đợt chi để duyệt')
  }));
}

async function send(mail){
  const from = process.env.MAIL_FROM;
  if (process.env.RESEND_API_KEY) {
    const r = await fetch('https://api.resend.com/emails', { method: 'POST',
      headers: { Authorization: `Bearer ${process.env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ from, to: [mail.to], subject: mail.subject, html: mail.html }) });
    if (!r.ok) throw new Error(`Resend ${r.status}: ${(await r.text()).slice(0, 200)}`);
    return;
  }
  const nodemailer = (await import('nodemailer')).default;
  const port = Number(process.env.SMTP_PORT || 465);
  const tx = nodemailer.createTransport({ host: process.env.SMTP_HOST, port, secure: port === 465,
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS } });
  await tx.sendMail({ from, to: mail.to, subject: mail.subject, html: mail.html });
}

export default async function handler(req, res){
  if (req.method !== 'POST') return res.status(405).json({ error: 'method' });
  if (!process.env.MAIL_FROM || !(process.env.RESEND_API_KEY || (process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS)))
    return res.status(503).json({ error: 'not_configured', message: 'Chưa cấu hình gửi email trên Vercel (MAIL_FROM và RESEND_API_KEY hoặc SMTP_*).' });
  const auth = req.headers.authorization || '';
  if (!auth.startsWith('Bearer ')) return res.status(401).json({ error: 'auth' });

  const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
  const items = (Array.isArray(body.items) ? body.items : []).slice(0, 20)
    .filter(x => x && (x.kind === 'req' || x.kind === 'batch') && typeof x.id === 'string');
  if (!items.length) return res.status(400).json({ error: 'items' });

  // Đọc dữ liệu bằng token của người gọi: RLS chỉ cho người có trong allowed_emails đọc.
  const sr = await fetch(`${SUPABASE_URL}/rest/v1/app_state?id=eq.main&select=data`, { headers: { apikey: SUPABASE_KEY, Authorization: auth } });
  if (!sr.ok) return res.status(401).json({ error: 'auth' });
  const rows = await sr.json(); const S = rows && rows[0] && rows[0].data;
  if (!S) return res.status(403).json({ error: 'forbidden' });
  const nt = (S.config && S.config.notify) || {};
  if (nt.email === false) return res.status(200).json({ sent: [], skipped: 'off' });

  const appUrl = (process.env.APP_URL || `https://${req.headers['x-forwarded-host'] || req.headers.host}`).replace(/\/$/, '');
  const mails = [];
  for (const it of items) {
    if (it.kind === 'req') { const r = (S.reqs || []).find(x => x.id === it.id); const m = r && buildReq(S, r, appUrl); if (m) mails.push(...m); }
    if (it.kind === 'batch' && nt.batch !== false) { const b = (S.batches || []).find(x => x.id === it.id); const m = b && buildBatch(S, b, appUrl); if (m) mails.push(...m); }
  }
  const sent = [], failed = [];
  for (const m of mails) { try { await send(m); sent.push(m.to); } catch (e) { failed.push({ to: m.to, error: String(e.message || e).slice(0, 200) }); } }
  return res.status(failed.length && !sent.length ? 502 : 200).json({ sent, failed });
}
