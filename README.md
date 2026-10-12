# Thông báo phạt nhà xe vận chuyển

Web app tạo file PDF **Thông báo phạt vi phạm hợp đồng** cho nhà xe khi có heo chết trên xe.

File PDF xuất ra gồm, theo thứ tự:

1. **Thông báo phạt vi phạm hợp đồng** (2 trang)
2. **Hình ảnh heo chết** ngày dd/mm/yyyy, số lượng, trọng lượng (chọn hình từ máy)
3. **Bảng giá** (file PDF upload lên, ghép nguyên vẹn vào cuối)

Chạy hoàn toàn trên trình duyệt: hình ảnh và file PDF không được gửi lên server nào.

## Dữ liệu MASTER.xlsx

App đọc file `MASTER.xlsx` đặt cùng thư mục với `index.html`:

| Sheet   | Cột |
|---------|-----|
| `BÊN A` | Tên Lò Giết Mổ, Tên Công Ty, MST, Địa Chỉ, Người Đại Diện, Số Tài Khoản Ngân Hàng, Tên Tài Khoản Ngân Hàng, Tên Ngân Hàng, Nội Dung Chuyển Khoản |
| `BÊN B` | Số Xe, Người Đại Diện, Tên Nhà Cung Cấp, Số Hợp Đồng, Điều Khoản, Ngày Ký Hợp Đồng |

Muốn cập nhật dữ liệu: sửa `MASTER.xlsx` rồi commit lại. Hoặc bấm **Chọn file MASTER** trong app để dùng một file khác (app nhớ file đó trên trình duyệt).

Mã lò trong số thông báo (VD `001/2026-TBVC-CPVN-LA`) khai báo ở đầu `assets/app.js`:

```js
const LO_CODE = { "Sagri": "SG", "Lộc An": "LA", "Cẩm Tú": "CT" };
```

Lò mới chưa có mã sẽ tự lấy chữ cái đầu của tên lò.

## Cách dùng

1. Chọn **lò giết mổ** và **số xe**: các ô lấy từ MASTER tự điền (nhạt màu). Bấm biểu tượng bút để sửa.
2. Nhập số thông báo (`001/2026`), ngày phát sinh (hạn thanh toán mặc định là cuối tháng), số heo, trọng lượng, đơn giá.
3. Chọn hình heo chết và file PDF bảng giá.
4. Bấm **Xuất file PDF**.

## Chạy trên máy

Vì trình duyệt chặn đọc file khi mở trực tiếp `index.html`, cần chạy một server nhỏ:

```bash
python -m http.server 8765
```

Rồi mở http://localhost:8765. (Nếu mở thẳng file `index.html`, app vẫn chạy nhưng phải bấm **Chọn file MASTER**.)

## Đưa lên GitHub Pages

1. Tạo repo trên GitHub và push code lên.
2. Vào **Settings → Pages**, chọn branch `main`, thư mục `/ (root)`.
3. Mở link `https://<tên-tài-khoản>.github.io/<tên-repo>/`.

> Lưu ý: `MASTER.xlsx` chứa số tài khoản ngân hàng, tên người đại diện. Nếu repo để public thì ai cũng xem được. Cân nhắc để repo private, hoặc bỏ `MASTER.xlsx` khỏi repo và dùng nút **Chọn file MASTER**.
