"""
Sinh file PPTX: Kiến trúc sư số – Không gian sáng tạo (Lớp 6)
Yêu cầu: python-pptx, Pillow
Chạy:  python3 slides/generate.py
Kết quả: slides/kien-truc-su-so-lop-6.pptx
"""

import math
import os
from pathlib import Path

# ---------- Pillow – tạo ảnh minh họa ----------
from PIL import Image, ImageDraw, ImageFont

# ---------- python-pptx ----------
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm, Pt, Emu

# ── Đường dẫn gốc repo ──────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).resolve().parent.parent
ASSETS_DIR = REPO_ROOT / "assets"
SLIDES_DIR = REPO_ROOT / "slides"
OUT_PPTX = SLIDES_DIR / "kien-truc-su-so-lop-6.pptx"

ASSETS_DIR.mkdir(exist_ok=True)
SLIDES_DIR.mkdir(exist_ok=True)

# ── Màu sắc ──────────────────────────────────────────────────────────────────
C_DARK_BLUE  = RGBColor(0x1F, 0x4E, 0x79)   # tiêu đề slide
C_MED_BLUE   = RGBColor(0x2E, 0x75, 0xB6)   # accent
C_LIGHT_BLUE = RGBColor(0xD6, 0xE4, 0xF0)   # nền nhạt
C_WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
C_BLACK      = RGBColor(0x1A, 0x1A, 0x1A)
C_GOLD       = RGBColor(0xFF, 0xC0, 0x00)   # điểm nhấn

PIL_DARK_BLUE  = (0x1F, 0x4E, 0x79)
PIL_MED_BLUE   = (0x2E, 0x75, 0xB6)
PIL_LIGHT_BLUE = (0xD6, 0xE4, 0xF0)
PIL_WHITE      = (0xFF, 0xFF, 0xFF)
PIL_GOLD       = (0xFF, 0xC0, 0x00)
PIL_BLACK      = (0x1A, 0x1A, 0x1A)
PIL_RED        = (0xC0, 0x00, 0x00)
PIL_GREEN      = (0x70, 0xAD, 0x47)

FONT = "Times New Roman"

# ─────────────────────────────────────────────────────────────────────────────
# Tạo ảnh minh họa bằng Pillow
# ─────────────────────────────────────────────────────────────────────────────

def _pil_font(size=28, bold=False):
    """Trả về ImageFont (fallback sang default nếu không có TTF)."""
    candidates = [
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"   if bold else
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSerif.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def make_image1(path: Path):
    """
    Collage minh họa HĐ1 – Thợ săn hình học:
    4 ô chứa các hình học từ thực tế trường học.
    """
    W, H = 900, 600
    img = Image.new("RGB", (W, H), PIL_LIGHT_BLUE)
    draw = ImageDraw.Draw(img)

    # Tiêu đề
    font_title = _pil_font(32, bold=True)
    draw.rectangle([0, 0, W, 70], fill=PIL_DARK_BLUE)
    draw.text((W // 2, 35), "Thư viện hình học quanh em",
              font=font_title, fill=PIL_WHITE, anchor="mm")

    font_label = _pil_font(22, bold=True)
    font_sub   = _pil_font(18)

    items = [
        ("Bảng tin",     "Hình chữ nhật", PIL_MED_BLUE,  [(50, 120), (390, 280)]),
        ("Viên gạch",    "Hình vuông",    PIL_GREEN,     [(460, 120), (800, 280)]),
        ("Bồn hoa",      "Hình thoi",     PIL_GOLD,      [(50, 320), (390, 530)]),
        ("Cầu thang",    "Hình thang",    PIL_RED,       [(460, 320), (800, 530)]),
    ]

    shapes = ["rect", "square", "rhombus", "trap"]

    for i, (label, shape_name, color, bbox) in enumerate(items):
        x0, y0, x1, y1 = bbox[0][0], bbox[0][1], bbox[1][0], bbox[1][1]
        cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
        w, h = x1 - x0, y1 - y0

        # Khung nền
        draw.rectangle([x0, y0, x1, y1], fill=PIL_WHITE, outline=color, width=3)

        # Vẽ hình học minh họa
        s = shapes[i]
        pad = 20
        if s == "rect":
            draw.rectangle([x0+pad, y0+pad+30, x1-pad, y1-pad-10],
                           fill=(*color, 180) if len(color)==3 else color,
                           outline=PIL_DARK_BLUE, width=2)
        elif s == "square":
            sq = min(w, h) // 2 - pad
            draw.rectangle([cx-sq, cy-sq+10, cx+sq, cy+sq+10],
                           fill=color, outline=PIL_DARK_BLUE, width=2)
        elif s == "rhombus":
            pts = [(cx, y0+pad+10), (x1-pad, cy+10),
                   (cx, y1-pad-10), (x0+pad, cy+10)]
            draw.polygon(pts, fill=color, outline=PIL_DARK_BLUE, width=2)
        elif s == "trap":
            off = 30
            pts = [(x0+pad+off, y0+pad+20), (x1-pad-off, y0+pad+20),
                   (x1-pad, y1-pad-10),     (x0+pad, y1-pad-10)]
            draw.polygon(pts, fill=color, outline=PIL_DARK_BLUE, width=2)

        # Nhãn
        draw.text((cx, y1-14), shape_name, font=font_sub,
                  fill=PIL_DARK_BLUE, anchor="mm")
        draw.text((cx, y0+12), label, font=font_label,
                  fill=color, anchor="mm")

    # Ghi chú góc dưới
    font_note = _pil_font(16)
    draw.text((W//2, H-8), "📸 Học sinh chụp ảnh + đo thực tế",
              font=font_note, fill=PIL_DARK_BLUE, anchor="mb")

    img.save(path, "PNG")
    print(f"  ✓ {path}")


def make_image2(path: Path):
    """
    Minh họa GeoGebra – hình chữ nhật & hình vuông với lệnh Area.
    """
    W, H = 900, 600
    img = Image.new("RGB", (W, H), (0xF2, 0xF2, 0xF2))
    draw = ImageDraw.Draw(img)

    # Thanh tiêu đề GeoGebra style
    draw.rectangle([0, 0, W, 50], fill=(0x33, 0x33, 0x33))
    font_title = _pil_font(26, bold=True)
    draw.text((W//2, 25), "GeoGebra – Số hóa bản vẽ",
              font=font_title, fill=PIL_WHITE, anchor="mm")

    # Lưới tọa độ
    grid_color = (0xCC, 0xCC, 0xCC)
    step = 40
    for x in range(0, W, step):
        draw.line([(x, 60), (x, H)], fill=grid_color, width=1)
    for y in range(60, H, step):
        draw.line([(0, y), (W, y)], fill=grid_color, width=1)

    # Trục tọa độ
    ax_color = (0x99, 0x99, 0x99)
    draw.line([(W//2, 60), (W//2, H)], fill=ax_color, width=2)
    draw.line([(0, H//2+30), (W, H//2+30)], fill=ax_color, width=2)

    font_lbl = _pil_font(20, bold=True)
    font_val = _pil_font(18)
    font_eq  = _pil_font(22, bold=True)

    # ── Hình chữ nhật (bên trái) ──────────────────────────────────────────
    rx0, ry0, rx1, ry1 = 80, 110, 320, 290
    draw.rectangle([rx0, ry0, rx1, ry1],
                   fill=(0xBD, 0xD7, 0xEE), outline=PIL_MED_BLUE, width=3)
    rcx, rcy = (rx0+rx1)//2, (ry0+ry1)//2
    draw.text((rcx, rcy), "S = 6 × 9 = 54 cm²",
              font=font_eq, fill=PIL_DARK_BLUE, anchor="mm")
    # Kích thước
    draw.text(((rx0+rx1)//2, ry0-12), "9 cm",  font=font_val,
              fill=PIL_MED_BLUE, anchor="mm")
    draw.text((rx0-28, (ry0+ry1)//2), "6 cm", font=font_val,
              fill=PIL_MED_BLUE, anchor="mm")
    draw.text((rcx, ry1+18), "Hình chữ nhật – S = a × b",
              font=font_lbl, fill=PIL_DARK_BLUE, anchor="mm")

    # ── Hình vuông (giữa) ─────────────────────────────────────────────────
    sq = 160
    sx0, sy0 = 380, 140
    sx1, sy1 = sx0+sq, sy0+sq
    scx, scy = (sx0+sx1)//2, (sy0+sy1)//2
    draw.rectangle([sx0, sy0, sx1, sy1],
                   fill=(0xA9, 0xD1, 0x8A), outline=PIL_GREEN, width=3)
    draw.text((scx, scy), "S = 8² = 64 cm²",
              font=font_eq, fill=(0x37, 0x5C, 0x23), anchor="mm")
    draw.text((scx, sy0-12), "8 cm", font=font_val,
              fill=PIL_GREEN, anchor="mm")
    draw.text((scx, sy1+18), "Hình vuông – S = a²",
              font=font_lbl, fill=(0x37, 0x5C, 0x23), anchor="mm")

    # ── Hình thoi (bên phải) ──────────────────────────────────────────────
    dhx, dhy = 750, 210
    d1, d2 = 120, 80
    pts = [(dhx, dhy-d2//2), (dhx+d1//2, dhy),
           (dhx, dhy+d2//2), (dhx-d1//2, dhy)]
    draw.polygon(pts, fill=(0xFF, 0xE0, 0x99), outline=PIL_GOLD, width=3)
    draw.text((dhx, dhy), "S = ½d₁d₂",
              font=font_eq, fill=(0x80, 0x60, 0x00), anchor="mm")
    draw.text((dhx, dhy+d2//2+18), "Hình thoi – S = ½×d₁×d₂",
              font=font_lbl, fill=(0x80, 0x60, 0x00), anchor="mm")

    # ── Hộp kết quả Area ──────────────────────────────────────────────────
    draw.rectangle([40, 380, 860, 560], fill=PIL_WHITE,
                   outline=PIL_DARK_BLUE, width=2)
    draw.rectangle([40, 380, 860, 420], fill=PIL_DARK_BLUE)
    draw.text((450, 400), "Lệnh Area (GeoGebra) – So sánh kết quả",
              font=font_lbl, fill=PIL_WHITE, anchor="mm")

    cols = ["Vật thể", "Tính tay (cm²)", "GeoGebra Area (cm²)", "Chênh lệch"]
    xs   = [130, 330, 570, 770]
    for x, c in zip(xs, cols):
        draw.text((x, 440), c, font=font_val, fill=PIL_DARK_BLUE, anchor="mm")

    rows = [
        ("Bảng tin",  "54",  "54.2", "≈ 0.4%"),
        ("Viên gạch", "144", "144",  "0%"),
        ("Bồn hoa",   "96",  "97.1", "≈ 1.1%"),
    ]
    ys = [470, 505, 540]
    for y, row in zip(ys, rows):
        for x, val in zip(xs, row):
            draw.text((x, y), val, font=font_val,
                      fill=PIL_BLACK, anchor="mm")

    img.save(path, "PNG")
    print(f"  ✓ {path}")


def make_image3(path: Path):
    """
    Bảng đối chiếu kết quả: thủ công vs GeoGebra.
    """
    W, H = 960, 520
    img = Image.new("RGB", (W, H), PIL_WHITE)
    draw = ImageDraw.Draw(img)

    font_h = _pil_font(24, bold=True)
    font_b = _pil_font(20, bold=True)
    font_n = _pil_font(18)

    # Tiêu đề
    draw.rectangle([0, 0, W, 60], fill=PIL_DARK_BLUE)
    draw.text((W//2, 30), "Bảng đối chiếu kết quả – HĐ2",
              font=font_h, fill=PIL_WHITE, anchor="mm")

    headers = ["STT", "Vật thể", "Hình dạng", "CT tính tay",
               "KQ tay", "KQ GGB", "Ghi chú"]
    col_w   = [50, 140, 140, 160, 100, 100, 200]
    col_x   = [20]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)

    # Header row
    draw.rectangle([10, 70, W-10, 110], fill=PIL_MED_BLUE)
    for x, h, w in zip(col_x, headers, col_w):
        draw.text((x + w//2, 90), h, font=font_b, fill=PIL_WHITE, anchor="mm")

    rows = [
        ("1", "Kệ sách",      "Hình thoi",    "S=½d₁d₂",     "96",  "97.1", "Sai số đo d"),
        ("2", "Bảng tin",     "Hình chữ nhật","S=a×b",        "54",  "54.2", "Gần khớp"),
        ("3", "Viên gạch hoa","Hình vuông",   "S=a²",         "144", "144",  "Khớp hoàn toàn"),
        ("4", "Cái bàn",      "Hình thang",   "S=(a+b)h/2",   "120", "122",  "Cần đo h chính xác"),
        ("5", "Cửa sổ",       "Hình chữ nhật","S=a×b",        "80",  "80.5", "Sai số đơn vị"),
    ]
    row_colors = [PIL_LIGHT_BLUE, PIL_WHITE]
    for i, row in enumerate(rows):
        y0 = 115 + i * 70
        y1 = y0 + 68
        draw.rectangle([10, y0, W-10, y1], fill=row_colors[i % 2])
        for x, val, w in zip(col_x, row, col_w):
            draw.text((x + w//2, (y0+y1)//2), val,
                      font=font_n, fill=PIL_BLACK, anchor="mm")

    # Đường kẻ lưới
    for x, w in zip(col_x[1:], col_w):
        draw.line([(x, 70), (x, H-20)], fill=PIL_MED_BLUE, width=1)
    for i in range(len(rows)+1):
        y = 115 + i * 70
        draw.line([(10, y), (W-10, y)], fill=PIL_MED_BLUE, width=1)
    draw.rectangle([10, 70, W-10, H-20], outline=PIL_DARK_BLUE, width=2)

    img.save(path, "PNG")
    print(f"  ✓ {path}")


# ─────────────────────────────────────────────────────────────────────────────
# Tiện ích tạo slide
# ─────────────────────────────────────────────────────────────────────────────

def _add_slide(prs, layout_idx=6):
    layout = prs.slide_layouts[layout_idx]
    return prs.slides.add_slide(layout)


def _txb(slide, text, x, y, w, h,
         font_size=24, bold=False, italic=False,
         color=None, align=PP_ALIGN.LEFT,
         font_name=FONT):
    """Thêm text box vào slide."""
    from pptx.util import Pt
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame
    tf.word_wrap = True
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name      = font_name
    run.font.size      = Pt(font_size)
    run.font.bold      = bold
    run.font.italic    = italic
    if color:
        run.font.color.rgb = color
    return txb


def _bg(slide, color: RGBColor, prs):
    """Đặt màu nền slide."""
    from pptx.oxml.ns import qn
    from lxml import etree
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _heading(slide, text, prs, y_cm=1.2, font_size=36):
    """Tiêu đề slide – dải màu đậm trên cùng."""
    W = prs.slide_width
    bar_h = Cm(1.8)
    bar = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        0, Cm(0), W, bar_h
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_DARK_BLUE
    bar.line.fill.background()

    _txb(slide, text,
         Cm(0.5), Cm(0.1), W - Cm(1), bar_h,
         font_size=font_size, bold=True,
         color=C_WHITE, align=PP_ALIGN.LEFT)


def _bullet_box(slide, items, x, y, w, h,
                font_size=22, bullet="▪ ", color=None, prs=None):
    """Thêm danh sách bullet."""
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = bullet + item
        run.font.name = FONT
        run.font.size = Pt(font_size)
        if color:
            run.font.color.rgb = color


def _add_image(slide, img_path, x, y, w, h):
    slide.shapes.add_picture(str(img_path), x, y, w, h)


def _accent_box(slide, x, y, w, h, color=None):
    """Hộp nền màu nhạt."""
    color = color or C_LIGHT_BLUE
    box = slide.shapes.add_shape(1, x, y, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = color
    box.line.color.rgb = C_MED_BLUE
    return box


# ─────────────────────────────────────────────────────────────────────────────
# 14 slides
# ─────────────────────────────────────────────────────────────────────────────

def build_pptx(img1, img2, img3):
    prs = Presentation()
    prs.slide_width  = Cm(33.867)   # 16:9
    prs.slide_height = Cm(19.05)

    W = prs.slide_width
    H = prs.slide_height

    def blank():
        return prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # ── Slide 1: Tiêu đề ─────────────────────────────────────────────────────
    sl = blank()
    # Nền gradient giả (hình nền màu đậm)
    bg = sl.shapes.add_shape(1, 0, 0, W, H)
    bg.fill.solid(); bg.fill.fore_color.rgb = C_DARK_BLUE; bg.line.fill.background()

    # Dải trắng giữa
    stripe = sl.shapes.add_shape(1, 0, Cm(5.5), W, Cm(8))
    stripe.fill.solid(); stripe.fill.fore_color.rgb = C_WHITE; stripe.line.fill.background()

    # Điểm nhấn vàng
    acc = sl.shapes.add_shape(1, 0, Cm(5.5), Cm(0.4), Cm(8))
    acc.fill.solid(); acc.fill.fore_color.rgb = C_GOLD; acc.line.fill.background()

    _txb(sl, "KIẾN TRÚC SƯ SỐ",
         Cm(1), Cm(6.0), W - Cm(2), Cm(3),
         font_size=48, bold=True, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)
    _txb(sl, "& KHÔNG GIAN SÁNG TẠO",
         Cm(1), Cm(8.2), W - Cm(2), Cm(2),
         font_size=32, bold=True, color=C_MED_BLUE, align=PP_ALIGN.CENTER)
    _txb(sl, "Hoạt động giáo dục theo chủ đề – Lớp 6  |  2 tiết",
         Cm(1), Cm(10.5), W - Cm(2), Cm(1.5),
         font_size=20, color=C_BLACK, align=PP_ALIGN.CENTER)
    _txb(sl, "Môn: Toán & Tin học  |  Tích hợp GeoGebra",
         Cm(1), Cm(17.2), W - Cm(2), Cm(1.5),
         font_size=18, italic=True, color=C_LIGHT_BLUE, align=PP_ALIGN.CENTER)

    # ── Slide 2: Mục tiêu ────────────────────────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "Mục tiêu chủ đề", prs)

    # Cột Toán
    _accent_box(sl, Cm(0.5), Cm(2.2), Cm(15.5), Cm(15.5))
    _txb(sl, "📐  TOÁN HỌC",
         Cm(0.8), Cm(2.4), Cm(15), Cm(1.2),
         font_size=22, bold=True, color=C_DARK_BLUE)
    _bullet_box(sl,
        ["Vận dụng công thức diện tích hình phẳng vào thực tiễn",
         "Đo đạc & tính diện tích bề mặt đồ vật thực tế",
         "Bước đầu làm quen với ước lượng diện tích"],
        Cm(0.8), Cm(3.8), Cm(15), Cm(10),
        font_size=20, color=C_BLACK)

    # Cột Tin
    _accent_box(sl, Cm(17), Cm(2.2), Cm(16.3), Cm(15.5), color=RGBColor(0xE2, 0xEF, 0xDA))
    _txb(sl, "💻  TIN HỌC",
         Cm(17.3), Cm(2.4), Cm(15.5), Cm(1.2),
         font_size=22, bold=True, color=RGBColor(0x37, 0x5C, 0x23))
    _bullet_box(sl,
        ["Sử dụng GeoGebra / Scratch để mô phỏng hình học",
         "Dựng hình & tính diện tích bằng công cụ số",
         "Kiểm chứng kết quả thủ công qua công cụ số"],
        Cm(17.3), Cm(3.8), Cm(15.5), Cm(10),
        font_size=20, color=C_BLACK)

    # ── Slide 3: Năng lực & Phẩm chất ────────────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "Năng lực & Phẩm chất", prs)

    _txb(sl, "🧠  Năng lực",
         Cm(0.5), Cm(2.2), Cm(32), Cm(1),
         font_size=22, bold=True, color=C_DARK_BLUE)
    _bullet_box(sl,
        ["Mô hình hoá toán học: chuyển vật thể thực → hình học",
         "Giải quyết vấn đề: lựa chọn công thức & công cụ phù hợp",
         "Sử dụng công cụ & phương tiện toán học (GeoGebra, Scratch…)"],
        Cm(0.5), Cm(3.3), Cm(32.5), Cm(6),
        font_size=20, color=C_BLACK)

    div = sl.shapes.add_shape(1, Cm(0.5), Cm(9.2), W - Cm(1), Cm(0.07))
    div.fill.solid(); div.fill.fore_color.rgb = C_MED_BLUE; div.line.fill.background()

    _txb(sl, "💎  Phẩm chất",
         Cm(0.5), Cm(9.5), Cm(32), Cm(1),
         font_size=22, bold=True, color=C_DARK_BLUE)
    _bullet_box(sl,
        ["Yêu nước: giữ gìn & làm đẹp cảnh quan trường học",
         "Trách nhiệm: hoàn thành nhiệm vụ nhóm; tiết kiệm nguyên vật liệu",
         "Chăm chỉ: kiên trì đo đạc; tìm tòi tối ưu hoá công việc"],
        Cm(0.5), Cm(10.7), Cm(32.5), Cm(7),
        font_size=20, color=C_BLACK)

    # ── Slide 4: Thiết bị / Học liệu ─────────────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "Thiết bị & Học liệu", prs)

    items = [
        ("🖥️", "Máy tính xách tay\n& Tivi thông minh"),
        ("📐", "Thước dây\n& Thước kẻ dài"),
        ("📱", "Máy tính cầm tay\n(Casio)"),
        ("🎨", "Giấy A0, A4\nBút dạ, Băng dính"),
        ("💾", "GeoGebra / Scratch\n(phần mềm miễn phí)"),
    ]
    cols = 5
    box_w = Cm(6)
    box_h = Cm(9)
    gap   = Cm(0.5)
    start_x = Cm(0.4)
    for i, (icon, label) in enumerate(items):
        bx = start_x + i * (box_w + gap)
        _accent_box(sl, bx, Cm(2.3), box_w, box_h)
        _txb(sl, icon, bx, Cm(3.0), box_w, Cm(3),
             font_size=38, align=PP_ALIGN.CENTER)
        _txb(sl, label, bx, Cm(6.0), box_w, Cm(5),
             font_size=18, align=PP_ALIGN.CENTER, color=C_DARK_BLUE)

    # ── Slide 5: Tiến trình ───────────────────────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "Tiến trình tổ chức (2 tiết)", prs)

    stages = [
        ("Mở đầu",      "15'", C_GOLD,               "Flashcards – Khởi động"),
        ("Khám phá",    "30'", C_MED_BLUE,            "HĐ1: Thợ săn hình học"),
        ("Rèn luyện",   "25'", RGBColor(0x70,0xAD,0x47), "HĐ2: Số hóa bản vẽ"),
        ("Vận dụng",    "15'", RGBColor(0xC0,0x00,0x00), "HĐ3 & HĐ4: Thiết kế & Đấu thầu"),
        ("Đánh giá",    "5'",  C_DARK_BLUE,           "Rubric + Chốt thông điệp"),
    ]
    bar_w = (W - Cm(1)) / len(stages)
    timeline_y = Cm(4)
    for i, (name, dur, color, desc) in enumerate(stages):
        bx = Cm(0.5) + i * bar_w
        # Hộp màu
        box = sl.shapes.add_shape(1, bx, timeline_y, bar_w - Cm(0.15), Cm(3.5))
        box.fill.solid(); box.fill.fore_color.rgb = color; box.line.fill.background()
        _txb(sl, name, bx, timeline_y + Cm(0.3), bar_w - Cm(0.15), Cm(1.2),
             font_size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        _txb(sl, dur,  bx, timeline_y + Cm(1.5),  bar_w - Cm(0.15), Cm(1.2),
             font_size=28, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        _txb(sl, desc, bx, Cm(8.5), bar_w - Cm(0.15), Cm(4),
             font_size=16, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)
        # Mũi tên
        if i < len(stages) - 1:
            arr = sl.shapes.add_shape(
                13,  # right arrow
                bx + bar_w - Cm(0.35), timeline_y + Cm(1.2), Cm(0.5), Cm(1)
            )
            arr.fill.solid(); arr.fill.fore_color.rgb = C_BLACK
            arr.line.fill.background()

    # Tổng thời gian
    _txb(sl, "⏱  Tổng: 90 phút  (2 tiết × 45 phút)",
         Cm(0.5), Cm(16.5), W - Cm(1), Cm(1.5),
         font_size=20, bold=True, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)

    # ── Slide 6: Khởi động – Flashcards ──────────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "Khởi động: Trò chơi Flashcards  (15')", prs)

    _txb(sl, "🃏  Luật chơi",
         Cm(0.5), Cm(2.3), Cm(32), Cm(1),
         font_size=24, bold=True, color=C_DARK_BLUE)
    _bullet_box(sl,
        ["Lớp chia nhóm 4–6 em theo tổ",
         "Mỗi bức ảnh hiện trong 5–10 giây",
         "Nhóm giơ tay nhanh + gọi đúng tên hình học → +1 điểm",
         "Đoán sai → mất quyền trả lời tấm thẻ tiếp theo"],
        Cm(0.5), Cm(3.5), Cm(19), Cm(10),
        font_size=21, color=C_BLACK)

    # Card minh họa
    for j, (icon, label) in enumerate([("🔷","Hình thoi"), ("⬛","Hình vuông"),
                                        ("▬","Chữ nhật"),  ("🔺","Hình thang")]):
        cx = Cm(21) + j * Cm(3.1)
        _accent_box(sl, cx, Cm(3.5), Cm(2.8), Cm(4))
        _txb(sl, icon,  cx, Cm(4.0), Cm(2.8), Cm(2),
             font_size=30, align=PP_ALIGN.CENTER)
        _txb(sl, label, cx, Cm(5.8), Cm(2.8), Cm(1.5),
             font_size=14, align=PP_ALIGN.CENTER, color=C_DARK_BLUE)

    _txb(sl, 'Nhịp độ 1 (Dễ): ảnh rõ ràng\nNhịp độ 2 – "Mắt Cú": ảnh cắt lớp / làm mờ',
         Cm(0.5), Cm(14.5), Cm(32), Cm(3.5),
         font_size=20, italic=True, color=C_MED_BLUE)
    _txb(sl, "💰  Điểm tích lũy = vốn cho các hoạt động sau!",
         Cm(0.5), Cm(17.2), Cm(32), Cm(1.5),
         font_size=22, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)

    # ── Slide 7: HĐ1 – Thợ săn hình học + image1 ─────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "HĐ1: Thợ săn hình học  (30')", prs)

    _txb(sl, "📋  Nhiệm vụ nhóm",
         Cm(0.5), Cm(2.2), Cm(16), Cm(1),
         font_size=22, bold=True, color=C_DARK_BLUE)
    _bullet_box(sl,
        ["Di chuyển quanh khuôn viên trường (10–15 phút)",
         "Chụp ≥ 4 đồ vật đại diện 4 hình khác nhau",
         "Đo kích thước thực tế bằng thước dây",
         "Ghi số liệu vào sổ tay / điện thoại"],
        Cm(0.5), Cm(3.3), Cm(16), Cm(10),
        font_size=20, color=C_BLACK)

    _add_image(sl, img1, Cm(17), Cm(2.2), Cm(16), Cm(12))

    _txb(sl, "🎯  Mỗi hình học đúng = 1 chiến lợi phẩm cho Thư viện hình học số",
         Cm(0.5), Cm(16.8), W - Cm(1), Cm(1.5),
         font_size=18, italic=True, color=C_MED_BLUE, align=PP_ALIGN.CENTER)

    # ── Slide 8: HĐ1 – Mẫu bảng ghi số liệu ────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "HĐ1: Ghi số liệu thực tế", prs)

    _txb(sl, "📊  Mẫu bảng – Thư viện hình học quanh em",
         Cm(0.5), Cm(2.2), Cm(32), Cm(1),
         font_size=22, bold=True, color=C_DARK_BLUE)

    # Vẽ bảng
    headers_tbl = ["Đồ vật", "Hình dạng", "Chiều dài (cm)", "Chiều rộng (cm)", "Ghi chú"]
    col_ws = [Cm(7), Cm(6.5), Cm(5.5), Cm(5.5), Cm(8)]
    xs = [Cm(0.5)]
    for w in col_ws[:-1]:
        xs.append(xs[-1] + w)

    hdr_y = Cm(3.5)
    box = sl.shapes.add_shape(1, Cm(0.5), hdr_y, W - Cm(1), Cm(1.5))
    box.fill.solid(); box.fill.fore_color.rgb = C_MED_BLUE; box.line.fill.background()
    for x, h, cw in zip(xs, headers_tbl, col_ws):
        _txb(sl, h, x, hdr_y + Cm(0.1), cw, Cm(1.3),
             font_size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    sample_rows = [
        ("Bảng tin",   "Hình chữ nhật", "180",  "90",  "Đơn vị: cm"),
        ("Viên gạch",  "Hình vuông",    "12",   "12",  "Đường chéo 17"),
        ("Bồn hoa",    "Hình thoi",     "d₁=16","d₂=12","Chụp từ trên"),
        ("Cửa sổ",     "Hình chữ nhật", "120",  "80",  ""),
    ]
    row_h = Cm(2.8)
    for i, row in enumerate(sample_rows):
        ry = hdr_y + Cm(1.5) + i * row_h
        bg_c = C_LIGHT_BLUE if i % 2 == 0 else C_WHITE
        rb = sl.shapes.add_shape(1, Cm(0.5), ry, W - Cm(1), row_h - Cm(0.05))
        rb.fill.solid(); rb.fill.fore_color.rgb = bg_c; rb.line.color.rgb = C_MED_BLUE
        for x, val, cw in zip(xs, row, col_ws):
            _txb(sl, val, x + Cm(0.1), ry + Cm(0.3), cw - Cm(0.2), row_h,
                 font_size=18, color=C_BLACK, align=PP_ALIGN.CENTER)

    _txb(sl, "✏️  Nhóm ghi thêm các đồ vật nhóm mình quan sát được",
         Cm(0.5), Cm(17.2), W - Cm(1), Cm(1.5),
         font_size=18, italic=True, color=C_MED_BLUE)

    # ── Slide 9: HĐ2 – Số hóa bản vẽ (GeoGebra steps) ───────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "HĐ2: Số hóa bản vẽ – GeoGebra  (25')", prs)

    steps = [
        ("NV1", "Chèn ảnh vật thể làm nền vào GeoGebra"),
        ("NV2", "Đánh dấu đỉnh → nối thành hình (Đoạn thẳng / Đa giác)"),
        ("NV3", "Chỉnh tỉ lệ khớp với số liệu đo thực tế"),
        ("NV4", "Tính tay diện tích theo công thức phù hợp"),
        ("NV5", "Dùng lệnh Area để GeoGebra tự tính diện tích"),
        ("NV6", "So sánh 2 kết quả – phân tích sai số (nếu có)"),
    ]
    cols2 = 2
    sw = (W - Cm(1.5)) / cols2
    sh = Cm(3.8)
    colors_step = [C_MED_BLUE, C_DARK_BLUE, RGBColor(0x70,0xAD,0x47),
                   C_GOLD, RGBColor(0xC0,0x00,0x00), C_DARK_BLUE]
    for i, (tag, desc) in enumerate(steps):
        col = i % cols2
        row = i // cols2
        bx = Cm(0.5) + col * sw
        by = Cm(2.3) + row * (sh + Cm(0.2))
        box = sl.shapes.add_shape(1, bx, by, sw - Cm(0.2), sh)
        box.fill.solid(); box.fill.fore_color.rgb = colors_step[i]
        box.line.fill.background()
        _txb(sl, tag,  bx + Cm(0.3), by + Cm(0.2), Cm(2), Cm(1),
             font_size=20, bold=True, color=C_WHITE)
        _txb(sl, desc, bx + Cm(2.5), by + Cm(0.6), sw - Cm(3), Cm(2.8),
             font_size=18, color=C_WHITE)

    _txb(sl, "🔑  Công thức: Hình thoi S = ½d₁d₂  |  Hình thang S = (a+b)h/2  |  Hình vuông S = a²",
         Cm(0.5), Cm(17), W - Cm(1), Cm(1.5),
         font_size=18, bold=True, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)

    # ── Slide 10: HĐ2 – Minh họa GeoGebra + image2 ───────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "HĐ2: Minh họa GeoGebra", prs)

    _add_image(sl, img2, Cm(0.3), Cm(2.2), Cm(22), Cm(13))

    _txb(sl, "💡  Lệnh Area tự động tính\ndiện tích → so sánh\nvới kết quả tính tay",
         Cm(23.2), Cm(3.5), Cm(10), Cm(6),
         font_size=20, color=C_DARK_BLUE)

    _txb(sl, "❓  Công cụ số giúp ích gì\ntrong kiểm tra tính toán?",
         Cm(23.2), Cm(10.5), Cm(10), Cm(5),
         font_size=20, italic=True, color=C_MED_BLUE)

    # ── Slide 11: Đối chiếu kết quả + image3 ─────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "Đối chiếu kết quả", prs)

    _add_image(sl, img3, Cm(0.3), Cm(2.2), Cm(25), Cm(13.5))

    _txb(sl, "📌  Nhận xét",
         Cm(26.2), Cm(2.5), Cm(7.3), Cm(1),
         font_size=20, bold=True, color=C_DARK_BLUE)
    _bullet_box(sl,
        ["Sai số do góc chụp",
         "Sai số đo đơn vị",
         "Thao tác điểm chưa chính xác",
         "Kết quả gần như trùng khớp = thành công!"],
        Cm(26.2), Cm(3.8), Cm(7.3), Cm(12),
        font_size=17, color=C_BLACK)

    # ── Slide 12: HĐ3 – Thiết kế gạch nghệ thuật ──────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "HĐ3: Thiết kế Gạch Nghệ Thuật  (15')", prs)

    _txb(sl, "📐  Phôi gạch hình vuông – cạnh 12 cm",
         Cm(0.5), Cm(2.2), Cm(20), Cm(1),
         font_size=22, bold=True, color=C_DARK_BLUE)
    _bullet_box(sl,
        ["NV1: Quan sát & điền phiếu học tập (tính S gạch, S hình thoi, S cung tròn)",
         "NV2: Thiết kế họa tiết (tam giác / thoi / chữ nhật) – kỹ thuật khăn trải bàn",
         "Yêu cầu: Có tính đối xứng (≥ 1 trục)",
         "Diện tích họa tiết < 40% diện tích viên gạch"],
        Cm(0.5), Cm(3.4), Cm(20), Cm(10),
        font_size=20, color=C_BLACK)

    # Minh họa viên gạch vẽ bằng code
    tile_x, tile_y, tile_s = Cm(22), Cm(3), Cm(13)
    tile_box = sl.shapes.add_shape(1, tile_x, tile_y, tile_s, tile_s)
    tile_box.fill.solid(); tile_box.fill.fore_color.rgb = RGBColor(0xF2,0xF2,0xF2)
    tile_box.line.color.rgb = C_DARK_BLUE

    # Hình thoi bên trong
    from pptx.util import Emu
    cx_emu = tile_x + tile_s / 2
    cy_emu = tile_y + tile_s / 2
    d1_emu = tile_s * 0.6
    d2_emu = tile_s * 0.6

    # Dùng shape arrow – không vẽ được polygon trực tiếp trong pptx dễ
    # Thay bằng một hình thoi giả bằng shape type=6 (diamond)
    diamond = sl.shapes.add_shape(
        4,  # diamond
        cx_emu - d1_emu/2, cy_emu - d2_emu/2,
        d1_emu, d2_emu
    )
    diamond.fill.solid(); diamond.fill.fore_color.rgb = C_MED_BLUE
    diamond.line.color.rgb = C_DARK_BLUE

    _txb(sl, "4 trục đối xứng ✓\nHọa tiết < 40% ✓",
         tile_x, tile_y + tile_s + Cm(0.3), tile_s, Cm(2),
         font_size=18, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)

    _txb(sl, "📝  S viên gạch = 12² = 144 cm²   |   40% × 144 = 57.6 cm²",
         Cm(0.5), Cm(17), W - Cm(1), Cm(1.5),
         font_size=20, bold=True, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)

    # ── Slide 13: HĐ4 – Đấu thầu thiết kế ───────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "HĐ4: Đấu thầu Thiết kế  (15')", prs)

    steps13 = [
        ("Bước 1 – 3'",  "Xem lại diện tích khu vực đã đo (GeoGebra)"),
        ("Bước 2 – 10'", "Vẽ & tính số viên gạch:\nS khu vực ÷ S 1 viên = số viên cần dùng"),
        ("Bước 3 – 5'",  "Triển lãm: dán mẫu gạch lên bảng\nGhi: khu vực + số lượng gạch"),
    ]
    for i, (title, desc) in enumerate(steps13):
        by = Cm(2.5) + i * Cm(5.3)
        box = sl.shapes.add_shape(1, Cm(0.5), by, W - Cm(1), Cm(5))
        box.fill.solid()
        box.fill.fore_color.rgb = [C_MED_BLUE, RGBColor(0x70,0xAD,0x47), C_GOLD][i]
        box.line.fill.background()
        _txb(sl, title, Cm(0.8), by + Cm(0.3), Cm(8), Cm(1.2),
             font_size=20, bold=True, color=C_WHITE)
        _txb(sl, desc, Cm(9), by + Cm(0.6), W - Cm(10), Cm(4),
             font_size=20, color=C_WHITE)

    _txb(sl, "🏆  Nhóm thiết kế đẹp + tính toán chính xác sẽ được tôn vinh!",
         Cm(0.5), Cm(18.2), W - Cm(1), Cm(0.8),
         font_size=18, bold=True, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)

    # ── Slide 14: Rubric + Chốt thông điệp ───────────────────────────────────
    sl = blank()
    _bg(sl, C_WHITE, prs)
    _heading(sl, "Đánh giá & Chốt thông điệp", prs)

    criteria = [
        ("Mô hình toán học",    "Đúng ≥ 4/4 hình, đo chính xác, đúng CT"),
        ("Sử dụng GeoGebra",    "Nhập ảnh, dựng hình, lệnh Area đúng"),
        ("Sáng tạo – Thẩm mỹ", "Đối xứng, phối màu đẹp, họa tiết < 40%"),
        ("Hợp tác & Thuyết trình","Phân công, trình bày tự tin, trả lời câu hỏi"),
        ("Chính xác – Trình bày","Số liệu đúng, bảng đẹp, font chuẩn"),
    ]
    hdr_y = Cm(2.3)
    hdr = sl.shapes.add_shape(1, Cm(0.5), hdr_y, W - Cm(1), Cm(1.2))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = C_DARK_BLUE; hdr.line.fill.background()
    for x_off, label, cw in [
        (Cm(0.5), "Tiêu chí", Cm(11)),
        (Cm(12), "Xuất sắc (3đ)", Cm(7)),
        (Cm(19.5), "Hoàn thành (2đ)", Cm(7)),
        (Cm(27), "Cần cải thiện (1đ)", Cm(6)),
    ]:
        _txb(sl, label, x_off + Cm(0.1), hdr_y + Cm(0.1), cw, Cm(1),
             font_size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    pts_descs = [
        ("Đúng & đầy đủ", "Sai sót nhỏ",   "Chưa rõ"),
        ("Thành thạo",     "Cơ bản",         "Chưa biết dùng"),
        ("Độc đáo, đẹp",  "Khả thi",         "Đơn giản"),
        ("Chủ động",       "Đúng nhiệm vụ",  "Thiếu phối hợp"),
        ("Chính xác, đẹp", "Rõ ràng",        "Nhiều lỗi"),
    ]
    row_h14 = Cm(2.8)
    for i, ((crit, _), (p3, p2, p1)) in enumerate(zip(criteria, pts_descs)):
        ry = hdr_y + Cm(1.2) + i * row_h14
        bg_c = C_LIGHT_BLUE if i % 2 == 0 else C_WHITE
        rb = sl.shapes.add_shape(1, Cm(0.5), ry, W - Cm(1), row_h14 - Cm(0.05))
        rb.fill.solid(); rb.fill.fore_color.rgb = bg_c; rb.line.color.rgb = C_MED_BLUE
        for x_off, label, cw in [
            (Cm(0.5), crit, Cm(11)),
            (Cm(12), p3,    Cm(7)),
            (Cm(19.5), p2,  Cm(7)),
            (Cm(27), p1,    Cm(6)),
        ]:
            _txb(sl, label, x_off + Cm(0.1), ry + Cm(0.3), cw, row_h14,
                 font_size=15, color=C_BLACK, align=PP_ALIGN.CENTER)

    # Thông điệp cuối
    msg_box = sl.shapes.add_shape(1, Cm(0.5), Cm(17.1), W - Cm(1), Cm(1.7))
    msg_box.fill.solid(); msg_box.fill.fore_color.rgb = C_DARK_BLUE
    msg_box.line.fill.background()
    _txb(sl, "\"Toán học không chỉ là những con số – hôm nay, các em là những Kiến trúc sư thực thụ!\"",
         Cm(0.5), Cm(17.2), W - Cm(1), Cm(1.5),
         font_size=18, bold=True, italic=True,
         color=C_WHITE, align=PP_ALIGN.CENTER)

    prs.save(str(OUT_PPTX))
    print(f"\n✅  Đã lưu: {OUT_PPTX}")


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Tạo ảnh minh họa…")
    img1 = ASSETS_DIR / "image1.png"
    img2 = ASSETS_DIR / "image2.png"
    img3 = ASSETS_DIR / "image3.png"
    make_image1(img1)
    make_image2(img2)
    make_image3(img3)

    print("\nSinh file PPTX…")
    build_pptx(img1, img2, img3)
