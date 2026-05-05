# Slide deck – Kiến trúc sư số & Không gian sáng tạo (Lớp 6)

## Nội dung

| File | Mô tả |
|------|-------|
| `slides/kien-truc-su-so-lop-6.pptx` | File thuyết trình chính (16:9, Times New Roman) |
| `slides/generate.py` | Script Python sinh lại file PPTX từ đầu |
| `assets/image1.png` | Minh họa HĐ1 – Thư viện hình học quanh em |
| `assets/image2.png` | Minh họa HĐ2 – GeoGebra & lệnh Area |
| `assets/image3.png` | Bảng đối chiếu kết quả thủ công vs. GeoGebra |

## Dàn ý 14 slides

| # | Tiêu đề |
|---|---------|
| 1 | Tiêu đề chủ đề |
| 2 | Mục tiêu (Toán + Tin học) |
| 3 | Năng lực & Phẩm chất |
| 4 | Thiết bị / Học liệu |
| 5 | Tiến trình (timeline 2 tiết) |
| 6 | Khởi động: Flashcards – luật chơi |
| 7 | HĐ1: Thợ săn hình học – nhiệm vụ + ảnh minh họa |
| 8 | HĐ1: Mẫu bảng ghi số liệu |
| 9 | HĐ2: Số hóa bản vẽ – các bước GeoGebra |
| 10 | HĐ2: Minh họa GeoGebra |
| 11 | Đối chiếu kết quả (bảng so sánh) |
| 12 | HĐ3: Thiết kế Gạch Nghệ Thuật |
| 13 | HĐ4: Đấu thầu Thiết kế & Triển lãm |
| 14 | Rubric đánh giá & Chốt thông điệp |

## Cách tái tạo file PPTX

### Yêu cầu

```bash
pip install python-pptx Pillow
```

### Chạy script

```bash
# Từ thư mục gốc của repo
python3 slides/generate.py
```

Script sẽ:
1. Tạo lại 3 ảnh minh họa trong `assets/`
2. Xuất file `slides/kien-truc-su-so-lop-6.pptx`

### Tuỳ chỉnh

- **Màu sắc**: chỉnh các hằng `C_DARK_BLUE`, `C_MED_BLUE`, `C_GOLD` trong phần đầu script.
- **Nội dung**: sửa trực tiếp các chuỗi văn bản trong từng hàm `build_pptx`.
- **Ảnh thật**: thay `assets/image1.png` / `image2.png` / `image3.png` bằng ảnh chụp thực tế, rồi chạy lại script (hoặc chỉ cập nhật file PPTX thủ công qua PowerPoint).

## Thông số kỹ thuật

| Thuộc tính | Giá trị |
|-----------|---------|
| Tỉ lệ | 16:9 (33.87 × 19.05 cm) |
| Font | Times New Roman (toàn bộ) |
| Phong cách | Ít chữ, thuyết trình miệng là chính |
| Bảng màu | Xanh navy / Xanh dương / Trắng / Vàng gold |
