-- Chạy file này trong Supabase: SQL Editor → New query → Run.
-- Dùng một bảng dữ liệu chung (jsonb) cho toàn bộ ứng dụng, và bảo vệ bằng RLS.

-- 1) Danh sách email được phép truy cập dữ liệu. Chỉ quản trị viên DB sửa bảng này (qua SQL Editor).
create table if not exists public.allowed_emails (
  email text primary key
);

-- 2) Dữ liệu ứng dụng: một dòng id = 'main'.
create table if not exists public.app_state (
  id         text primary key,
  data       jsonb not null,
  version    integer not null default 1,
  updated_at timestamptz not null default now()
);

alter table public.allowed_emails enable row level security;
alter table public.app_state      enable row level security;

-- Hàm kiểm tra email đăng nhập có trong danh sách được phép hay không.
create or replace function public.is_allowed()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.allowed_emails
    where lower(email) = lower(auth.email())
  );
$$;

-- Không có chính sách nào cho allowed_emails: client không đọc được danh sách này.
-- Chỉ người dùng đăng nhập có email trong allowed_emails được đọc/ghi app_state.
drop policy if exists "app_state read"   on public.app_state;
drop policy if exists "app_state insert" on public.app_state;
drop policy if exists "app_state update" on public.app_state;

create policy "app_state read" on public.app_state
  for select to authenticated using (public.is_allowed());

create policy "app_state insert" on public.app_state
  for insert to authenticated with check (public.is_allowed());

create policy "app_state update" on public.app_state
  for update to authenticated using (public.is_allowed()) with check (public.is_allowed());

-- Không tạo chính sách delete: không ai xoá được dòng dữ liệu từ ứng dụng.

-- ===== Việc cần làm sau khi chạy file này =====
-- (a) Tắt đăng ký công khai: Authentication → Sign In / Providers → tắt "Allow new users to sign up".
-- (b) Thêm email quản trị đầu tiên vào danh sách được phép, ví dụ:
--     insert into public.allowed_emails (email) values ('ten.quan.tri@congty.com');
-- (c) Tạo tài khoản người dùng: Authentication → Users → Add user (nhập email + mật khẩu).
--     Thêm email của từng người vào allowed_emails khi cấp quyền truy cập.
