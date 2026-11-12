# Xperise Spend Ops

Ứng dụng đề xuất chi và phê duyệt chi phí. Giao diện chạy trên trình duyệt (Vite + JavaScript thuần), dữ liệu lưu trên Supabase, đăng nhập bằng tài khoản Supabase Auth.

## 1. Chạy trên máy

```bash
npm install
cp .env.example .env          # rồi điền VITE_SUPABASE_URL và VITE_SUPABASE_ANON_KEY
npm run dev                   # mở http://localhost:5173
```

## 2. Tạo cơ sở dữ liệu Supabase

1. Tạo project tại https://supabase.com (chọn region Singapore).
2. Vào **SQL Editor**, dán toàn bộ nội dung `supabase/schema.sql`, bấm **Run**.
3. **Authentication → Sign In / Providers**: tắt **Allow new users to sign up** (không cho đăng ký công khai).
4. **Authentication → Users → Add user**: tạo tài khoản quản trị đầu tiên (email + mật khẩu).
5. **SQL Editor** cấp quyền truy cập cho email đó:
   ```sql
   insert into public.allowed_emails (email) values ('ten.quan.tri@congty.com');
   ```
6. **Project Settings → API**: lấy `Project URL` và `anon public key` điền vào `.env` (máy) và biến môi trường Vercel (xem bước 4).

Lần đăng nhập đầu tiên, khi chưa có người dùng nào trong ứng dụng, tài khoản đó tự thành **quản trị**, được gán vào mọi bước duyệt. Sau đó vào **Người dùng** để thêm nhân sự (mỗi người có email đăng nhập), rồi vào **Cấu hình** để chỉnh ma trận duyệt. Nhớ thêm email từng người vào `allowed_emails` để họ đăng nhập được.

## 3. Đưa lên GitHub

```bash
git init
git add .
git commit -m "Xperise Spend Ops"
git branch -M main
git remote add origin https://github.com/<tai-khoan>/xperise-spend-ops.git
git push -u origin main
```

## 4. Deploy lên Vercel

1. Vào https://vercel.com → **Add New… → Project** → chọn repo vừa push.
2. Framework: **Vite** (tự nhận). Build command: `npm run build`, output: `dist`.
3. **Environment Variables**: thêm `VITE_SUPABASE_URL` và `VITE_SUPABASE_ANON_KEY`.
4. **Deploy**. Sau đó vào Supabase → **Authentication → URL Configuration**, thêm domain Vercel vào *Site URL* và *Redirect URLs*.

## Lưu ý bảo mật và giới hạn hiện tại

- Mọi người được cấp quyền (có email trong `allowed_emails`) đều đọc được toàn bộ dữ liệu trong bảng `app_state`, gồm số tài khoản NCC. Ứng dụng chỉ ẩn bớt trên giao diện theo vai trò. Nếu cần chặn ở tầng cơ sở dữ liệu, phải tách dữ liệu thành nhiều bảng có RLS riêng cho từng vai trò.
- Ghi dữ liệu theo phiên bản: nếu hai người cùng sửa, người lưu sau sẽ được báo tải lại trang, không tự ghi đè.
- File đính kèm (hoá đơn, UNC) hiện chỉ lưu tên file, chưa tải lên Supabase Storage. Cần bổ sung nếu cần lưu file thật.
- Dữ liệu mới bắt đầu trống, không có dữ liệu mẫu của bản prototype.
