# Công cụ NC — bản React (Vite + Tailwind), triển khai qua Vercel

## Chạy thử trên máy (cần Node.js 18+)
```
npm install
npm run dev        # dev server, có hot-reload
npm run build       # build ra thư mục dist/ để deploy
npm run preview     # xem thử bản build
```

## Đưa lên Vercel — cách dễ nhất (qua GitHub)
1. Tạo 1 repo GitHub mới, đẩy (push) toàn bộ code này lên (trừ `node_modules`, `dist` — đã có sẵn trong `.gitignore`).
2. Vào https://vercel.com → đăng nhập bằng tài khoản GitHub.
3. Bấm "Add New" → "Project" → chọn đúng repo vừa tạo.
4. Vercel tự nhận diện đây là project Vite (nhờ file `vercel.json` đã cấu hình sẵn) — chỉ cần bấm "Deploy".
5. Sau ~1 phút có ngay đường link dạng `ten-du-an.vercel.app` — gửi link này cho mọi máy trong xưởng dùng.
6. Mỗi lần bạn push code mới lên GitHub, Vercel tự động build và cập nhật lại link — không cần thao tác gì thêm.

## Cách khác — deploy thẳng từ máy, không cần GitHub
```
npm install -g vercel
vercel login
vercel --prod
```

## Cấu trúc project
```
src/
 |- main.jsx                  (điểm khởi động React)
 |- App.jsx                   (sidebar + điều hướng Trang chủ / tính năng)
 |- index.css                 (Tailwind + font Inter)
 |- components/
 |   |- HomePanel.jsx          (trang chủ)
 |   |- NcCalculator.jsx       (giao diện tính năng — không tính toán ở đây)
 |- lib/
 |   |- utils.js                (hàm tiện ích chung)
 |   |- rules.js                (★ TOÀN BỘ ĐIỀU KIỆN TÍNH TOÁN — đọc comment đầu file)
 |   |- useReferenceData.js     (hook quản lý bảng tra cứu NC编码 + 掺用比例)
 |- data/
     |- nc-code-map.json        (dữ liệu NC编码 mặc định, nhúng sẵn lúc build)
     |- ratio-map.json          (dữ liệu 掺用比例 mặc định, nhúng sẵn lúc build)
```

Logic tính toán y hệt bản Vue trước — chỉ đổi khung giao diện sang React + Tailwind
để tối ưu về mặt thiết kế, và bỏ chế độ single-file offline vì giờ chạy qua Vercel
(https://) không còn bị giới hạn `file://` nữa.

## Về bảng tra cứu và quyền riêng tư
- App vẫn xử lý toàn bộ dữ liệu (đọc Excel, tính toán) ngay trong trình duyệt người dùng —
  Vercel chỉ phát tán code, không nhận/lưu dữ liệu sản xuất hay bảng tra cứu của bạn.
- Cập nhật bảng tra cứu (NC编码/掺用比例): chọn file mới ở app — lưu qua localStorage
  theo từng trình duyệt/máy, giống cơ chế bản trước.
