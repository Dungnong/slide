"""
Sinh file PPTX: Thuật Toán A* – Tìm Đường trong Lưới 2D
Yêu cầu: python-pptx, Pillow
Chạy:  python3 slides/generate_astar.py
Kết quả: slides/astar-slides.pptx

Format: Mở đầu → Cơ sở lý thuyết → Thiết kế → Thực nghiệm → Kết luận
"""

import os
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm, Pt, Emu

# ── Đường dẫn ────────────────────────────────────────────────────────────────
REPO_ROOT  = Path(__file__).resolve().parent.parent
ASSETS_DIR = REPO_ROOT / "assets"
SLIDES_DIR = REPO_ROOT / "slides"
OUT_PPTX   = SLIDES_DIR / "astar-slides.pptx"

ASSETS_DIR.mkdir(exist_ok=True)
SLIDES_DIR.mkdir(exist_ok=True)

# ── Bảng màu (Canva-style: deep teal + orange accent) ────────────────────────
C_DARK     = RGBColor(0x0D, 0x2B, 0x45)   # nền tối
C_TEAL     = RGBColor(0x00, 0x7B, 0x83)   # accent chính
C_ORANGE   = RGBColor(0xFF, 0x6B, 0x35)   # điểm nhấn
C_LIGHT    = RGBColor(0xF0, 0xF7, 0xF8)   # nền sáng
C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
C_GRAY     = RGBColor(0x5A, 0x6A, 0x72)
C_YELLOW   = RGBColor(0xFF, 0xD1, 0x66)

PIL_DARK    = (0x0D, 0x2B, 0x45)
PIL_TEAL    = (0x00, 0x7B, 0x83)
PIL_ORANGE  = (0xFF, 0x6B, 0x35)
PIL_LIGHT   = (0xF0, 0xF7, 0xF8)
PIL_WHITE   = (0xFF, 0xFF, 0xFF)
PIL_YELLOW  = (0xFF, 0xD1, 0x66)
PIL_GREEN   = (0x00, 0xAF, 0x91)
PIL_RED     = (0xE5, 0x3E, 0x3E)
PIL_PURPLE  = (0x88, 0x4E, 0xA0)
PIL_GRAY    = (0x5A, 0x6A, 0x72)

# Widescreen 16:9
SLIDE_W = Cm(33.87)
SLIDE_H = Cm(19.05)


# ─────────────────────────────────────────────────────────────────────────────
# Font helpers
# ─────────────────────────────────────────────────────────────────────────────

def _pil_font(size=28, bold=False):
    candidates = [
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"    if bold else
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


# ─────────────────────────────────────────────────────────────────────────────
# Asset images
# ─────────────────────────────────────────────────────────────────────────────

def make_grid_image(path: Path):
    """Vẽ lưới 2D minh họa A* với màu sắc: Start, End, Open, Closed, Path."""
    W, H = 800, 600
    img = Image.new("RGB", (W, H), PIL_DARK)
    draw = ImageDraw.Draw(img)

    COLS, ROWS = 12, 9
    cw = W // COLS
    ch = H // ROWS

    BARRIER = {(2,1),(2,2),(2,3),(2,4),(2,5),(2,6),
               (5,2),(5,3),(5,4),(5,5),(5,6),(5,7),
               (8,0),(8,1),(8,2),(8,3),(8,4),(8,5)}
    CLOSED  = {(1,0),(1,1),(1,2),(3,0),(3,1),(4,0),(4,1),(4,2),
               (6,0),(6,1),(6,2),(7,0),(7,1)}
    OPEN    = {(3,2),(3,3),(4,3),(6,3),(7,2),(7,3)}
    PATH    = {(0,0),(1,0),(2,0),(3,0),(4,0),(5,0),(6,0),(7,0),
               (7,1),(7,2),(7,3),(7,4),(7,5),(7,6),(7,7),(7,8)}
    START   = (0, 0)
    END     = (11, 8)

    for r in range(ROWS):
        for c in range(COLS):
            x0, y0 = c * cw, r * ch
            x1, y1 = x0 + cw - 2, y0 + ch - 2
            cell = (c, r)
            if cell == START:
                color = PIL_ORANGE
            elif cell == END:
                color = (0x00, 0x80, 0xFF)
            elif cell in BARRIER:
                color = (0x22, 0x22, 0x22)
            elif cell in PATH:
                color = PIL_PURPLE
            elif cell in CLOSED:
                color = PIL_RED
            elif cell in OPEN:
                color = PIL_GREEN
            else:
                color = (0x1A, 0x3A, 0x4A)
            draw.rectangle([x0, y0, x1, y1], fill=color)

    # Legend
    legend = [
        (PIL_ORANGE,  "Start"),
        ((0x00, 0x80, 0xFF), "End"),
        (PIL_GREEN,   "Open List"),
        (PIL_RED,     "Closed List"),
        (PIL_PURPLE,  "Đường đi"),
        ((0x22,0x22,0x22), "Vật cản"),
    ]
    lx, ly = 10, H - 36
    fn = _pil_font(18, bold=False)
    for col, label in legend:
        draw.rectangle([lx, ly, lx+18, ly+18], fill=col)
        draw.text((lx + 22, ly), label, font=fn, fill=PIL_WHITE)
        lx += 130

    img.save(path)


def make_formula_image(path: Path):
    """Hình minh họa công thức f(n) = g(n) + h(n)."""
    W, H = 900, 400
    img = Image.new("RGB", (W, H), PIL_DARK)
    draw = ImageDraw.Draw(img)

    fn_big  = _pil_font(52, bold=True)
    fn_med  = _pil_font(32, bold=False)
    fn_sm   = _pil_font(24, bold=False)

    # Công thức chính
    draw.text((W//2, 60), "f(n)  =  g(n)  +  h(n)", font=fn_big,
              fill=PIL_YELLOW, anchor="mm")

    # Giải thích 3 phần
    boxes = [
        (150, 180, "g(n)", "Chi phí thực tế\ntừ Start → n", PIL_GREEN),
        (450, 180, "h(n)", "Heuristic ước lượng\nn → Goal", PIL_ORANGE),
        (750, 180, "f(n)", "Tổng chi phí\ndự kiến qua n", PIL_TEAL),
    ]
    for cx, cy, lbl, desc, color in boxes:
        # Circle
        r = 55
        draw.ellipse([cx-r, cy-r, cx+r, cy+r], fill=color)
        draw.text((cx, cy), lbl, font=fn_med, fill=PIL_WHITE, anchor="mm")
        # Description below
        for i, line in enumerate(desc.split("\n")):
            draw.text((cx, cy + r + 20 + i*30), line,
                      font=fn_sm, fill=PIL_WHITE, anchor="mm")

    # Arrow from g to +
    draw.line([(230, 180), (340, 180)], fill=PIL_YELLOW, width=3)
    draw.line([(560, 180), (670, 180)], fill=PIL_YELLOW, width=3)

    img.save(path)


def make_heuristic_image(path: Path):
    """So sánh Manhattan vs Euclidean trên lưới."""
    W, H = 900, 480
    img = Image.new("RGB", (W, H), PIL_DARK)
    draw = ImageDraw.Draw(img)

    fn_title = _pil_font(30, bold=True)
    fn_label = _pil_font(22, bold=False)
    fn_sm    = _pil_font(18, bold=False)

    COLS, ROWS = 7, 6
    SZ = 60
    PAD = 30

    def draw_grid(ox, oy, path_cells, color, title):
        draw.text((ox + COLS*SZ//2, oy - 24), title,
                  font=fn_title, fill=PIL_YELLOW, anchor="mm")
        for r in range(ROWS):
            for c in range(COLS):
                x0, y0 = ox + c*SZ, oy + r*SZ
                x1, y1 = x0 + SZ - 3, y0 + SZ - 3
                cell = (c, r)
                if cell == (0, 0):
                    fill = PIL_ORANGE
                elif cell == (6, 5):
                    fill = (0x00, 0x80, 0xFF)
                elif cell in path_cells:
                    fill = color
                else:
                    fill = (0x1A, 0x3A, 0x4A)
                draw.rectangle([x0, y0, x1, y1], fill=fill)
        # Labels
        draw.text((ox, oy), "S", font=fn_sm, fill=PIL_WHITE)
        draw.text((ox + 6*SZ + 4, oy + 5*SZ + 4), "G",
                  font=fn_sm, fill=PIL_WHITE)

    # Manhattan path (L-shaped: go right then down)
    man_path = {(c, 0) for c in range(7)} | {(6, r) for r in range(6)}
    draw_grid(PAD, 80, man_path, PIL_GREEN, "Manhattan (4 hướng)")

    # Euclidean path (diagonal)
    euc_path = {(i, i) for i in range(6)} | {(6, 5)}
    draw_grid(PAD + 7*SZ + 80, 80, euc_path, PIL_ORANGE, "Euclidean (đường chim bay)")

    # Formulas below
    draw.text((PAD + COLS*SZ//2, 80 + ROWS*SZ + 30),
              "h = |x₁-x₂| + |y₁-y₂|",
              font=fn_label, fill=PIL_GREEN, anchor="mm")
    draw.text((PAD + 7*SZ + 80 + COLS*SZ//2, 80 + ROWS*SZ + 30),
              "h = √((x₁-x₂)² + (y₁-y₂)²)",
              font=fn_label, fill=PIL_ORANGE, anchor="mm")

    img.save(path)


def make_algo_flow_image(path: Path):
    """Sơ đồ luồng thuật toán A*."""
    W, H = 900, 550
    img = Image.new("RGB", (W, H), PIL_DARK)
    draw = ImageDraw.Draw(img)

    fn = _pil_font(22, bold=False)
    fn_sm = _pil_font(18, bold=False)

    steps = [
        (450, 50,  "Thêm Start vào Open List",      PIL_ORANGE),
        (450, 150, "Lấy nút n có f(n) nhỏ nhất",    PIL_TEAL),
        (450, 250, "n là Goal?",                     PIL_YELLOW),
        (700, 350, "Truy hồi đường đi ✓",            PIL_GREEN),
        (300, 350, "Duyệt láng giềng m của n",       PIL_TEAL),
        (300, 450, "Cập nhật g(m), f(m); thêm vào Open List", PIL_TEAL),
    ]

    box_w, box_h = 320, 55
    for cx, cy, text, color in steps:
        x0 = cx - box_w//2
        y0 = cy - box_h//2
        draw.rounded_rectangle([x0, y0, x0+box_w, y0+box_h],
                                radius=10, fill=color)
        draw.text((cx, cy), text, font=fn_sm, fill=PIL_WHITE,
                  anchor="mm")

    # Arrows
    def arrow(x1, y1, x2, y2):
        draw.line([(x1, y1), (x2, y2)], fill=PIL_WHITE, width=2)
        # arrowhead
        dx, dy = x2 - x1, y2 - y1
        length = math.sqrt(dx*dx + dy*dy)
        if length == 0:
            return
        ux, uy = dx/length, dy/length
        ax1 = int(x2 - 10*ux + 6*uy)
        ay1 = int(y2 - 10*uy - 6*ux)
        ax2 = int(x2 - 10*ux - 6*uy)
        ay2 = int(y2 - 10*uy + 6*ux)
        draw.polygon([(x2, y2), (ax1, ay1), (ax2, ay2)], fill=PIL_WHITE)

    arrow(450, 77, 450, 122)
    arrow(450, 177, 450, 222)
    arrow(450, 277, 700, 322)   # YES → Goal
    arrow(450, 277, 300, 322)   # NO  → neighbors
    draw.text((580, 296), "CÓ", font=fn_sm, fill=PIL_GREEN, anchor="mm")
    draw.text((370, 296), "KHÔNG", font=fn_sm, fill=PIL_RED, anchor="mm")
    arrow(300, 377, 300, 422)
    # back arrow
    draw.line([(160, 450), (160, 150)], fill=PIL_GRAY, width=2)
    arrow(160, 150, 290, 150)
    draw.text((120, 300), "Lặp", font=fn_sm, fill=PIL_GRAY, anchor="mm")

    img.save(path)


def make_result_summary_image(path: Path):
    """Bảng đánh giá 3 kịch bản."""
    W, H = 900, 420
    img = Image.new("RGB", (W, H), PIL_DARK)
    draw = ImageDraw.Draw(img)

    fn_h  = _pil_font(28, bold=True)
    fn_hd = _pil_font(22, bold=True)
    fn    = _pil_font(20, bold=False)

    headers = ["Kịch bản", "Không gian", "Vùng duyệt", "Kết quả"]
    rows = [
        ["1 – Trống",   "Không vật cản",     "Lớn (chi phí bằng nhau)", "✓ Tìm được, tối ưu"],
        ["2 – Chia cắt","Tường hình ✚ lớn", "Lan rộng tìm khe hở",      "✓ Vượt qua, tối ưu"],
        ["3 – Mê cung", "Ngõ cụt chằng chịt","Vào nhánh, không bị lạc", "✓ Thoát mê cung"],
    ]
    col_w = [150, 200, 280, 250]
    col_x = [20]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)

    row_h = 70
    header_y = 30
    colors_col = [PIL_TEAL, PIL_LIGHT, PIL_LIGHT, PIL_GREEN]

    # Header row
    for i, (hd, cx, cw) in enumerate(zip(headers, col_x, col_w)):
        draw.rectangle([cx, header_y, cx+cw-4, header_y+50], fill=PIL_TEAL)
        draw.text((cx + cw//2, header_y + 25), hd,
                  font=fn_hd, fill=PIL_WHITE, anchor="mm")

    # Data rows
    row_colors = [PIL_DARK, (0x0A, 0x2A, 0x3A)]
    for ri, row in enumerate(rows):
        y = header_y + 54 + ri * row_h
        for ci, (cell, cx, cw) in enumerate(zip(row, col_x, col_w)):
            bg = row_colors[ri % 2]
            draw.rectangle([cx, y, cx+cw-4, y+row_h-4], fill=bg)
            text_color = PIL_ORANGE if ci == 0 else (PIL_GREEN if ci == 3 else PIL_WHITE)
            draw.text((cx + cw//2, y + row_h//2), cell,
                      font=fn, fill=text_color, anchor="mm")

    img.save(path)


# ─────────────────────────────────────────────────────────────────────────────
# PPTX helpers
# ─────────────────────────────────────────────────────────────────────────────

def new_prs():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    return prs


def blank_layout(prs):
    return prs.slide_layouts[6]   # hoàn toàn trống


def add_rect(slide, x, y, w, h, color):
    from pptx.util import Emu
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_text(slide, text, x, y, w, h,
             font_size=24, bold=False, color=None,
             align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size   = Pt(font_size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = color if color else C_WHITE
    return txb


def add_image(slide, img_path, x, y, w, h):
    slide.shapes.add_picture(str(img_path), x, y, w, h)


def section_bar(slide, label, color=None):
    """Thanh nhỏ ở đầu slide hiển thị tên phần."""
    color = color or C_TEAL
    add_rect(slide, Cm(0), Cm(0), SLIDE_W, Cm(0.8), color)
    add_text(slide, label, Cm(0.4), Cm(0), Cm(20), Cm(0.8),
             font_size=11, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT)


def slide_number(slide, num):
    add_text(slide, str(num), SLIDE_W - Cm(1.5), SLIDE_H - Cm(0.9),
             Cm(1.2), Cm(0.7), font_size=11, color=C_GRAY, align=PP_ALIGN.RIGHT)


def full_bg(slide, color):
    add_rect(slide, Cm(0), Cm(0), SLIDE_W, SLIDE_H, color)


def bullet_block(slide, items, x, y, w, h, font_size=20, color=None, indent="• "):
    """Thêm danh sách bullet vào slide."""
    color = color or C_WHITE
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = indent + item
        run.font.size  = Pt(font_size)
        run.font.color.rgb = color


# ─────────────────────────────────────────────────────────────────────────────
# Individual slides
# ─────────────────────────────────────────────────────────────────────────────

def slide_01_title(prs):
    """Slide 1 – Title (Mở đầu)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)

    # Accent top bar
    add_rect(slide, Cm(0), Cm(0), SLIDE_W, Cm(1.2), C_TEAL)

    # Orange side accent
    add_rect(slide, Cm(0), Cm(1.2), Cm(0.6), SLIDE_H - Cm(1.2), C_ORANGE)

    # Title
    add_text(slide, "Thuật Toán A* (A-Star)",
             Cm(1.2), Cm(2.5), Cm(25), Cm(3),
             font_size=44, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT)

    # Subtitle
    add_text(slide, "Tìm Đường Ngắn Nhất trong Lưới 2D có Vật Cản",
             Cm(1.2), Cm(5.8), Cm(25), Cm(2),
             font_size=24, bold=False, color=C_YELLOW, align=PP_ALIGN.LEFT)

    # Decorative grid illustration (text art)
    add_text(slide, "🔷  →  ■  →  ★",
             Cm(1.2), Cm(8.5), Cm(14), Cm(1.5),
             font_size=28, color=C_TEAL, align=PP_ALIGN.LEFT)

    # Info bottom
    add_rect(slide, Cm(0), SLIDE_H - Cm(3), SLIDE_W, Cm(3), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "Môn học: Trí Tuệ Nhân Tạo  |  Bài Tập Lớn",
             Cm(1.2), SLIDE_H - Cm(2.8), Cm(30), Cm(1),
             font_size=18, color=C_LIGHT, align=PP_ALIGN.LEFT)
    add_text(slide, "Python  +  Pygame  |  Heuristic: Manhattan Distance",
             Cm(1.2), SLIDE_H - Cm(1.8), Cm(30), Cm(1),
             font_size=16, color=C_GRAY, align=PP_ALIGN.LEFT)


def slide_02_problem(prs):
    """Slide 2 – Đặt vấn đề (Mở đầu)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "MỞ ĐẦU")

    add_text(slide, "Đặt Vấn Đề",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    add_rect(slide, Cm(1), Cm(3.2), Cm(14.5), Cm(13), RGBColor(0x07, 0x1E, 0x32))
    bullet_block(slide, [
        "Robot kho bãi, drone, nhân vật game cần di chuyển tự động",
        "Di chuyển = đối mặt với chướng ngại vật phức tạp",
        "Sai lộ trình → va chạm, thiệt hại tài sản, đình trệ sản xuất",
        "Yêu cầu: tìm đường khả thi VÀ tối ưu nhất",
    ], Cm(1.5), Cm(3.5), Cm(13.5), Cm(12), font_size=20)

    # Right panel – applications
    add_rect(slide, Cm(16.5), Cm(3.2), Cm(16), Cm(5.5), C_TEAL)
    add_text(slide, "Ứng dụng thực tiễn",
             Cm(16.8), Cm(3.4), Cm(15), Cm(1),
             font_size=16, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
    apps = ["🤖 Robot logistics & kho bãi", "🚁 Drone tự hành", "🎮 AI trong game",
            "🗺️ GPS & bản đồ số", "🚨 Cứu hộ nguy hiểm"]
    bullet_block(slide, apps, Cm(16.8), Cm(4.6), Cm(15.2), Cm(3.8),
                 font_size=17, color=C_DARK, indent="")

    add_rect(slide, Cm(16.5), Cm(9.2), Cm(16), Cm(6.8), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "Thách thức cốt lõi",
             Cm(16.8), Cm(9.4), Cm(15), Cm(1),
             font_size=16, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    add_text(slide, "Làm thế nào để xác định lộ trình\nNGẮN NHẤT và KHÉO LÉO NHẤT\ntrong không gian lưới 2D đầy\nchướng ngại vật?",
             Cm(16.8), Cm(10.5), Cm(15.2), Cm(5),
             font_size=18, color=C_WHITE, align=PP_ALIGN.LEFT)

    slide_number(slide, 2)


def slide_03_objective(prs):
    """Slide 3 – Mục tiêu (Mở đầu)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "MỞ ĐẦU")

    add_text(slide, "Mục Tiêu & Giải Pháp",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    # Three objective boxes
    boxes = [
        (Cm(1),    "🎯 Mục tiêu",      C_TEAL,
         "Nghiên cứu & hiện thực hóa thuật toán A* giải bài toán tìm đường trong lưới 2D có vật cản"),
        (Cm(11.5), "⚙️ Phương pháp",   C_ORANGE,
         "Kết hợp hàm Heuristic để giới hạn phạm vi tìm kiếm, khắc phục duyệt rỗng của Dijkstra"),
        (Cm(22),   "💻 Triển khai",     C_TEAL,
         "Python + Pygame: giao diện trực quan, vẽ vật cản bằng chuột, mô phỏng từng bước"),
    ]
    for bx, title, color, desc in boxes:
        add_rect(slide, bx, Cm(3.5), Cm(10), Cm(8), color)
        add_text(slide, title, bx + Cm(0.3), Cm(3.8), Cm(9.4), Cm(1.2),
                 font_size=20, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
        add_text(slide, desc, bx + Cm(0.3), Cm(5.3), Cm(9.4), Cm(5.8),
                 font_size=18, color=C_DARK, align=PP_ALIGN.LEFT)

    # Bottom
    add_text(slide, "Đóng góp: Công cụ trực quan hóa A* – chuyển lý thuyết toán học thành hình ảnh dễ hiểu",
             Cm(1), Cm(13), Cm(31), Cm(1.5),
             font_size=19, italic=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

    slide_number(slide, 3)


def slide_04_outline(prs):
    """Slide 4 – Bố cục (Mở đầu)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "MỞ ĐẦU")

    add_text(slide, "Nội Dung Trình Bày",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    sections = [
        ("01", "MỞ ĐẦU",          C_ORANGE, "Đặt vấn đề – Mục tiêu – Định hướng"),
        ("02", "CƠ SỞ LÝ THUYẾT", C_TEAL,   "Đồ thị – Bài toán tìm đường – Thuật toán A*"),
        ("03", "THIẾT KẾ",        C_YELLOW,  "Môi trường – Kiến trúc – Cài đặt"),
        ("04", "THỰC NGHIỆM",     C_ORANGE,  "3 kịch bản – Đánh giá hiệu suất"),
        ("05", "KẾT LUẬN",        C_TEAL,   "Thành tựu – Hạn chế – Hướng phát triển"),
    ]
    for i, (num, title, color, desc) in enumerate(sections):
        y = Cm(3.5 + i * 2.8)
        add_rect(slide, Cm(1), y, Cm(2.5), Cm(2.2), color)
        add_text(slide, num, Cm(1), y, Cm(2.5), Cm(2.2),
                 font_size=28, bold=True, color=C_DARK, align=PP_ALIGN.CENTER)
        add_rect(slide, Cm(3.8), y, Cm(28), Cm(2.2), RGBColor(0x07, 0x1E, 0x32))
        add_text(slide, title, Cm(4.2), y + Cm(0.2), Cm(12), Cm(1),
                 font_size=18, bold=True, color=color, align=PP_ALIGN.LEFT)
        add_text(slide, desc, Cm(4.2), y + Cm(1.1), Cm(27), Cm(1),
                 font_size=16, color=C_LIGHT, align=PP_ALIGN.LEFT)

    slide_number(slide, 4)


def slide_05_graph(prs):
    """Slide 5 – Đồ thị (Cơ sở lý thuyết)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "CƠ SỞ LÝ THUYẾT")

    add_text(slide, "Cấu Trúc Đồ Thị (Graph)",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    add_text(slide, "G = (V, E)",
             Cm(1), Cm(3.2), Cm(10), Cm(1.5),
             font_size=28, bold=True, color=C_YELLOW, align=PP_ALIGN.LEFT)

    defs = [
        "V – Vertices (đỉnh): vị trí hoặc trạng thái trong không gian",
        "E – Edges (cạnh): liên kết / đường đi giữa các đỉnh",
    ]
    bullet_block(slide, defs, Cm(1), Cm(4.8), Cm(20), Cm(3), font_size=20)

    types = [
        ("Vô hướng",  "Cạnh 2 chiều đối xứng",         C_TEAL),
        ("Có hướng",  "Cạnh 1 chiều (cung)",             C_ORANGE),
        ("Có trọng số","Chi phí/khoảng cách trên mỗi cạnh", C_YELLOW),
    ]
    for i, (t, d, col) in enumerate(types):
        x = Cm(1 + i * 10.8)
        add_rect(slide, x, Cm(8), Cm(10), Cm(3.5), col)
        add_text(slide, t, x + Cm(0.3), Cm(8.3), Cm(9.4), Cm(1),
                 font_size=18, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
        add_text(slide, d, x + Cm(0.3), Cm(9.4), Cm(9.4), Cm(1.8),
                 font_size=16, color=C_DARK, align=PP_ALIGN.LEFT)

    add_text(slide, "→ Bài toán tìm đường dùng Đồ thị có trọng số trên lưới 2D",
             Cm(1), Cm(12.3), Cm(31), Cm(1.2),
             font_size=20, italic=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    slide_number(slide, 5)


def slide_06_pathfinding(prs, grid_img: Path):
    """Slide 6 – Bài toán tìm đường (Cơ sở lý thuyết)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "CƠ SỞ LÝ THUYẾT")

    add_text(slide, "Bài Toán Tìm Đường của Nhóm",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    # Left: problem definition
    add_rect(slide, Cm(1), Cm(3.2), Cm(15), Cm(12.5), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "Đặc điểm bài toán",
             Cm(1.3), Cm(3.4), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    items = [
        "Không gian lưới 2D: 40×40 ô (tọa độ x, y)",
        "Vật cản (Barrier): ô không thể đi qua",
        "Điểm Start (cam) → Goal (xanh dương)",
        "Heuristic: Khoảng cách Manhattan",
        "Di chuyển: 4 hướng (Lên/Xuống/Trái/Phải)",
    ]
    bullet_block(slide, items, Cm(1.3), Cm(4.7), Cm(14.4), Cm(10), font_size=18)

    # Right: grid image
    grid_img = ASSETS_DIR / "astar_grid.png"
    if grid_img.exists():
        add_image(slide, grid_img, Cm(17), Cm(3.2), Cm(15.5), Cm(12.5))

    slide_number(slide, 6)


def slide_07_astar_intro(prs):
    """Slide 7 – Giới thiệu A* (Cơ sở lý thuyết)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "CƠ SỞ LÝ THUYẾT")

    add_text(slide, "Thuật Toán A* (A-Star)",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    # Comparison: Blind vs Informed
    add_rect(slide, Cm(1), Cm(3.2), Cm(14.5), Cm(7.5), RGBColor(0x4A, 0x15, 0x15))
    add_text(slide, "❌ Tìm kiếm Mù (Blind Search)",
             Cm(1.3), Cm(3.4), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    blind = ["BFS: duyệt toàn bộ đồng đều", "Dijkstra: chính xác nhưng lãng phí tài nguyên",
             "Duyệt cả vùng không cần thiết"]
    bullet_block(slide, blind, Cm(1.3), Cm(4.8), Cm(14), Cm(5.5), font_size=18, color=C_LIGHT)

    add_rect(slide, Cm(17), Cm(3.2), Cm(15.5), Cm(7.5), RGBColor(0x07, 0x2E, 0x1A))
    add_text(slide, "✅ A* – Tìm kiếm Có Thông Tin",
             Cm(17.3), Cm(3.4), Cm(15), Cm(1),
             font_size=18, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)
    smart = ["Dùng Heuristic định hướng đến đích", "Thu hẹp đáng kể không gian tìm kiếm",
             "Vẫn đảm bảo tìm đường ngắn nhất"]
    bullet_block(slide, smart, Cm(17.3), Cm(4.8), Cm(15), Cm(5.5), font_size=18, color=C_LIGHT)

    # VS label
    add_rect(slide, Cm(15.5), Cm(5.5), Cm(1), Cm(3), C_ORANGE)
    add_text(slide, "VS", Cm(15.5), Cm(6.3), Cm(1), Cm(1.4),
             font_size=16, bold=True, color=C_DARK, align=PP_ALIGN.CENTER)

    add_text(slide, "A* = Chi phí thực tế g(n)  +  Ước lượng Heuristic h(n)",
             Cm(1), Cm(11.5), Cm(31), Cm(1.2),
             font_size=22, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

    add_text(slide, "Thuật toán A* được công bố bởi Hart, Nilsson & Raphael (1968) – IEEE Transactions",
             Cm(1), Cm(13), Cm(31), Cm(1),
             font_size=15, italic=True, color=C_GRAY, align=PP_ALIGN.CENTER)

    slide_number(slide, 7)


def slide_08_formula(prs, formula_img: Path):
    """Slide 8 – Công thức f(n)=g(n)+h(n) (Cơ sở lý thuyết)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "CƠ SỞ LÝ THUYẾT")

    add_text(slide, "Hàm Đánh Giá  f(n) = g(n) + h(n)",
             Cm(1), Cm(1.2), Cm(31), Cm(1.8),
             font_size=32, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    if formula_img.exists():
        add_image(slide, formula_img, Cm(1), Cm(3.2), Cm(31.5), Cm(11.5))

    add_text(slide, "Điều kiện tối ưu: h(n) ≤ chi phí thực tế → A* đảm bảo tìm đường NGẮN NHẤT",
             Cm(1), Cm(15.2), Cm(31), Cm(1.2),
             font_size=18, italic=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

    slide_number(slide, 8)


def slide_09_heuristic(prs, heur_img: Path):
    """Slide 9 – Hàm Heuristic (Cơ sở lý thuyết)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "CƠ SỞ LÝ THUYẾT")

    add_text(slide, "Các Hàm Heuristic",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    if heur_img.exists():
        add_image(slide, heur_img, Cm(1), Cm(3), Cm(31.5), Cm(10.5))

    # Comparison note
    items = [
        "Chebyshev: di chuyển 8 hướng, chi phí chéo = chi phí ngang",
        "Octile: tối ưu nhất cho 8 hướng khi chi phí chéo = √2",
        "→ Nhóm dùng Manhattan vì môi trường di chuyển 4 hướng",
    ]
    bullet_block(slide, items, Cm(1), Cm(13.8), Cm(31), Cm(3.8),
                 font_size=17, color=C_LIGHT)

    slide_number(slide, 9)


def slide_10_algo_flow(prs, flow_img: Path):
    """Slide 10 – Quy trình A* (Cơ sở lý thuyết)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "CƠ SỞ LÝ THUYẾT")

    add_text(slide, "Quy Trình Hoạt Động của A*",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    if flow_img.exists():
        add_image(slide, flow_img, Cm(1), Cm(3), Cm(19), Cm(13.5))

    # Right side: key concepts
    add_rect(slide, Cm(21), Cm(3), Cm(11.5), Cm(5.5), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "Open List (Tập mở)",
             Cm(21.3), Cm(3.2), Cm(11), Cm(1),
             font_size=16, bold=True, color=C_GREEN if False else C_TEAL, align=PP_ALIGN.LEFT)
    add_text(slide, "Các nút đã phát hiện,\nchưa được duyệt xong",
             Cm(21.3), Cm(4.3), Cm(11), Cm(2), font_size=15, color=C_LIGHT, align=PP_ALIGN.LEFT)

    add_rect(slide, Cm(21), Cm(9), Cm(11.5), Cm(5.5), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "Closed List (Tập đóng)",
             Cm(21.3), Cm(9.2), Cm(11), Cm(1),
             font_size=16, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    add_text(slide, "Các nút đã duyệt xong.\nKhông xét lại để tránh\nvòng lặp vô tận",
             Cm(21.3), Cm(10.3), Cm(11), Cm(3), font_size=15, color=C_LIGHT, align=PP_ALIGN.LEFT)

    slide_number(slide, 10)


def slide_11_design_env(prs):
    """Slide 11 – Môi trường thiết kế (Thiết kế)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "THIẾT KẾ")

    add_text(slide, "Môi Trường & Thiết Lập",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_YELLOW, align=PP_ALIGN.LEFT)

    cards = [
        ("🐍 Python",      C_TEAL,   "Xử lý thuật toán\nlinh hoạt, nhanh"),
        ("🎮 Pygame",      C_ORANGE, "Đồ họa lưới 2D\n& sự kiện chuột"),
        ("🗺️ Lưới 40×40", C_YELLOW, "800×800 px\n1600 Node"),
        ("📐 Manhattan",   C_TEAL,   "Heuristic 4 hướng\ntối ưu cho lưới"),
    ]
    for i, (title, color, desc) in enumerate(cards):
        x = Cm(1 + i * 8.2)
        add_rect(slide, x, Cm(3.5), Cm(7.7), Cm(5.5), color)
        add_text(slide, title, x + Cm(0.3), Cm(3.7), Cm(7.1), Cm(1.5),
                 font_size=20, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
        add_text(slide, desc, x + Cm(0.3), Cm(5.5), Cm(7.1), Cm(3),
                 font_size=17, color=C_DARK, align=PP_ALIGN.LEFT)

    # Color convention
    add_text(slide, "Quy ước màu sắc hiển thị",
             Cm(1), Cm(9.7), Cm(15), Cm(1),
             font_size=20, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT)

    colors_legend = [
        ("🟠 Cam – Start",   "Điểm xuất phát"),
        ("🔵 Xanh – End",    "Điểm đích"),
        ("⬛ Đen – Barrier", "Vật cản"),
        ("🟢 Xanh lá – Open","Đang xếp hàng"),
        ("🔴 Đỏ – Closed",  "Đã duyệt xong"),
        ("🟣 Tím – Path",   "Đường đi tối ưu"),
    ]
    for i, (lbl, desc) in enumerate(colors_legend):
        col = i % 3
        row = i // 3
        x = Cm(1 + col * 10.8)
        y = Cm(11 + row * 2.2)
        add_rect(slide, x, y, Cm(10.3), Cm(1.9), RGBColor(0x07, 0x1E, 0x32))
        add_text(slide, lbl, x + Cm(0.3), y + Cm(0.1), Cm(5.8), Cm(0.9),
                 font_size=15, bold=True, color=C_LIGHT, align=PP_ALIGN.LEFT)
        add_text(slide, desc, x + Cm(0.3), y + Cm(0.9), Cm(9.7), Cm(0.9),
                 font_size=14, color=C_GRAY, align=PP_ALIGN.LEFT)

    slide_number(slide, 11)


def slide_12_design_arch(prs):
    """Slide 12 – Kiến trúc cài đặt (Thiết kế)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "THIẾT KẾ")

    add_text(slide, "Kiến Trúc & Cấu Trúc Dữ Liệu",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_YELLOW, align=PP_ALIGN.LEFT)

    # Priority Queue box
    add_rect(slide, Cm(1), Cm(3.2), Cm(15), Cm(6.5), C_TEAL)
    add_text(slide, "Priority Queue",
             Cm(1.3), Cm(3.4), Cm(14.4), Cm(1.2),
             font_size=22, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
    add_text(slide, "Lấy nút có f(n) nhỏ nhất trong O(log n)\nĐảm bảo luôn chọn nút tốt nhất để mở rộng",
             Cm(1.3), Cm(4.8), Cm(14.4), Cm(4.5),
             font_size=18, color=C_DARK, align=PP_ALIGN.LEFT)

    # Hash Set box
    add_rect(slide, Cm(17.5), Cm(3.2), Cm(15), Cm(6.5), C_ORANGE)
    add_text(slide, "Hash Set (open_set_hash)",
             Cm(17.8), Cm(3.4), Cm(14.4), Cm(1.2),
             font_size=22, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
    add_text(slide, "Kiểm tra nút trong Open List trong O(1)\nLoại bỏ trùng lặp siêu nhanh, tránh duyệt lại",
             Cm(17.8), Cm(4.8), Cm(14.4), Cm(4.5),
             font_size=18, color=C_DARK, align=PP_ALIGN.LEFT)

    # Interaction design
    add_rect(slide, Cm(1), Cm(10.2), Cm(31.5), Cm(5.8), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "Giao diện người dùng (UI)",
             Cm(1.3), Cm(10.4), Cm(15), Cm(1),
             font_size=18, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    ui_items = [
        "Click trái: vẽ Start → End → Barrier (vật cản)",
        "Click phải: xóa ô đã vẽ",
        "Phím SPACE: bắt đầu chạy thuật toán A*",
        "Hoạt ảnh (animation) cập nhật từng khung hình – quan sát ra quyết định thời gian thực",
    ]
    bullet_block(slide, ui_items, Cm(1.3), Cm(11.6), Cm(31), Cm(4),
                 font_size=18, color=C_LIGHT)

    slide_number(slide, 12)


def slide_13_exp1(prs, grid_img: Path):
    """Slide 13 – Kịch bản 1 (Thực nghiệm)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "THỰC NGHIỆM")

    add_text(slide, "Kịch Bản 1: Không Gian Trống",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    add_rect(slide, Cm(1), Cm(3.2), Cm(14.5), Cm(13), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "📋 Mô tả",
             Cm(1.3), Cm(3.4), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    add_text(slide, "Start: góc trên trái\nGoal: góc dưới phải\nKhông có vật cản",
             Cm(1.3), Cm(4.7), Cm(14), Cm(2.5), font_size=18, color=C_LIGHT, align=PP_ALIGN.LEFT)
    add_text(slide, "📊 Nhận xét",
             Cm(1.3), Cm(7.8), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)
    add_text(slide, "Nhiều đường cùng chi phí tối ưu → thuật toán duyệt vùng rộng (màu đỏ)\n\nĐường đi cuối cùng (màu tím) men theo mép trên và mép phải không gian",
             Cm(1.3), Cm(9), Cm(14), Cm(6.5), font_size=17, color=C_LIGHT, align=PP_ALIGN.LEFT)

    if grid_img.exists():
        add_image(slide, grid_img, Cm(16.5), Cm(3.2), Cm(16), Cm(13))

    slide_number(slide, 13)


def slide_14_exp2(prs, grid_img: Path):
    """Slide 14 – Kịch bản 2 (Thực nghiệm)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "THỰC NGHIỆM")

    add_text(slide, "Kịch Bản 2: Không Gian Bị Chia Cắt",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    add_rect(slide, Cm(1), Cm(3.2), Cm(14.5), Cm(13), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "📋 Mô tả",
             Cm(1.3), Cm(3.4), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    add_text(slide, "Tường hình ✚ lớn chia cắt\nStart và Goal\nChỉ có 1 lối thoát hẹp ở mép phải",
             Cm(1.3), Cm(4.7), Cm(14), Cm(3), font_size=18, color=C_LIGHT, align=PP_ALIGN.LEFT)
    add_text(slide, "📊 Nhận xét",
             Cm(1.3), Cm(8.2), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)
    add_text(slide, "A* lan rộng tìm khe hở → ngay khi phát hiện hành lang, hội tụ tức thì\n\nĐường đi (tím) vòng qua rìa tường chính xác",
             Cm(1.3), Cm(9.5), Cm(14), Cm(6), font_size=17, color=C_LIGHT, align=PP_ALIGN.LEFT)

    if grid_img.exists():
        add_image(slide, grid_img, Cm(16.5), Cm(3.2), Cm(16), Cm(13))

    slide_number(slide, 14)


def slide_15_exp3(prs, grid_img: Path):
    """Slide 15 – Kịch bản 3 (Thực nghiệm)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "THỰC NGHIỆM")

    add_text(slide, "Kịch Bản 3: Mê Cung Phức Tạp",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    add_rect(slide, Cm(1), Cm(3.2), Cm(14.5), Cm(13), RGBColor(0x07, 0x1E, 0x32))
    add_text(slide, "📋 Mô tả",
             Cm(1.3), Cm(3.4), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    add_text(slide, "Nhiều tường chằng chịt\nTạo ngõ cụt (dead-end)\nHành lang hẹp dạng ziczac",
             Cm(1.3), Cm(4.7), Cm(14), Cm(3), font_size=18, color=C_LIGHT, align=PP_ALIGN.LEFT)
    add_text(slide, "📊 Nhận xét",
             Cm(1.3), Cm(8.2), Cm(14), Cm(1),
             font_size=18, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)
    add_text(slide, "A* vào các nhánh phụ nhưng KHÔNG bị lạc\n\nPhản hồi gần như tức thời nhờ PriorityQueue + Hash Set\n\nĐường đi (tím) len lỏi qua khúc cua hoàn hảo",
             Cm(1.3), Cm(9.5), Cm(14), Cm(6), font_size=17, color=C_LIGHT, align=PP_ALIGN.LEFT)

    if grid_img.exists():
        add_image(slide, grid_img, Cm(16.5), Cm(3.2), Cm(16), Cm(13))

    slide_number(slide, 15)


def slide_16_eval(prs, result_img: Path):
    """Slide 16 – Đánh giá kết quả (Thực nghiệm)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "THỰC NGHIỆM")

    add_text(slide, "Đánh Giá Kết Quả Chung",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)

    if result_img.exists():
        add_image(slide, result_img, Cm(1), Cm(3.2), Cm(31.5), Cm(9))

    items = [
        "Tính chính xác: Luôn tìm được đường đi ngắn nhất, truy vết came_from không lỗi",
        "Hiệu năng: PriorityQueue O(log n) + Hash Set O(1) → phản hồi tức thời dù 1600 Node",
        "Giao diện: Hoạt ảnh mượt mà, thao tác click trực quan, dễ quan sát cơ chế A*",
    ]
    bullet_block(slide, items, Cm(1), Cm(12.6), Cm(31.5), Cm(5.5), font_size=18)

    slide_number(slide, 16)


def slide_17_conclusion(prs):
    """Slide 17 – Tổng kết (Kết luận)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "KẾT LUẬN")

    add_text(slide, "Tổng Kết Kết Quả Đạt Được",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    achievements = [
        ("📐 Lý thuyết",   C_TEAL,
         "Hệ thống hóa đầy đủ nguyên lý A*\nHàm f(n)=g(n)+h(n) và hàm Heuristic Manhattan\nSo sánh với Dijkstra, BFS"),
        ("💻 Ứng dụng",    C_ORANGE,
         "Chương trình Python + Pygame hoàn chỉnh\nGiao diện trực quan, tương tác chuột\nMô phỏng từng bước lan tỏa"),
        ("⚡ Hiệu suất",   C_YELLOW,
         "PriorityQueue + Hash Set kép\nPhản hồi tức thời trong mọi kịch bản\nĐường đi tối ưu 100% độ chính xác"),
    ]
    for i, (title, color, desc) in enumerate(achievements):
        x = Cm(1 + i * 10.8)
        add_rect(slide, x, Cm(3.5), Cm(10.3), Cm(10), color)
        add_text(slide, title, x + Cm(0.3), Cm(3.7), Cm(9.7), Cm(1.2),
                 font_size=20, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
        add_text(slide, desc, x + Cm(0.3), Cm(5.2), Cm(9.7), Cm(7.8),
                 font_size=17, color=C_DARK, align=PP_ALIGN.LEFT)

    add_text(slide,
             "✅ Đề tài hoàn thành mục tiêu: Công cụ trực quan hóa A* – lý thuyết toán học → hình ảnh dễ hiểu",
             Cm(1), Cm(14.2), Cm(31.5), Cm(1.5),
             font_size=18, italic=True, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

    slide_number(slide, 17)


def slide_18_limitation(prs):
    """Slide 18 – Hạn chế & Hướng phát triển (Kết luận)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "KẾT LUẬN")

    add_text(slide, "Hạn Chế & Hướng Phát Triển",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    # Left: Limitations
    add_rect(slide, Cm(1), Cm(3.2), Cm(14.5), Cm(10.5), RGBColor(0x4A, 0x15, 0x15))
    add_text(slide, "⚠️ Hạn chế hiện tại",
             Cm(1.3), Cm(3.4), Cm(14), Cm(1),
             font_size=20, bold=True, color=C_ORANGE, align=PP_ALIGN.LEFT)
    limits = [
        "Chỉ di chuyển 4 hướng → đường bậc thang cứng nhắc",
        "Môi trường tĩnh – chưa xử lý vật cản di động",
        "Lưới kích thước cố định 40×40",
    ]
    bullet_block(slide, limits, Cm(1.3), Cm(4.8), Cm(14), Cm(7.5), font_size=18, color=C_LIGHT)

    # Right: Future
    add_rect(slide, Cm(17), Cm(3.2), Cm(15.5), Cm(10.5), RGBColor(0x07, 0x2E, 0x1A))
    add_text(slide, "🚀 Hướng phát triển",
             Cm(17.3), Cm(3.4), Cm(15), Cm(1),
             font_size=20, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)
    future = [
        "Mở rộng 8 hướng: Chebyshev / Octile heuristic",
        "Path Smoothing: làm mịn đường đi tự nhiên hơn",
        "D* / LPA*: xử lý vật cản di động thời gian thực",
        "Tùy chỉnh kích thước lưới & import bản đồ ngoài",
    ]
    bullet_block(slide, future, Cm(17.3), Cm(4.8), Cm(15), Cm(7.5), font_size=18, color=C_LIGHT)

    add_text(slide,
             "Mục tiêu xa hơn: ứng dụng cho robot kho bãi, drone tự hành và hệ thống tự hành thực tế",
             Cm(1), Cm(14.5), Cm(31.5), Cm(1.5),
             font_size=17, italic=True, color=C_YELLOW, align=PP_ALIGN.CENTER)

    slide_number(slide, 18)


def slide_19_references(prs):
    """Slide 19 – Tài liệu tham khảo (Kết luận)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)
    section_bar(slide, "KẾT LUẬN")

    add_text(slide, "Tài Liệu Tham Khảo",
             Cm(1), Cm(1.2), Cm(30), Cm(1.8),
             font_size=34, bold=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    refs = [
        "[1] Từ Minh Phương (2015). Giáo trình Nhập môn Trí tuệ nhân tạo. PTIT.",
        "[2] Hart, Nilsson & Raphael (1968). A Formal Basis for Heuristic Determination of Min Cost Paths. IEEE.",
        "[3] Stuart Russell & Peter Norvig (2020). Artificial Intelligence: A Modern Approach (4th Ed.). Pearson.",
        "[4] Pygame Community. Pygame Documentation. https://www.pygame.org/docs/",
        "[5] Amit Patel. Introduction to A* Algorithm. Red Blob Games.\n     https://www.redblobgames.com/pathfinding/a-star/introduction.html",
        "[6] Magnus Lie Hetland (2014). Python Algorithms. Apress.",
        "[7] Mã nguồn thực nghiệm: https://github.com/Vu52Hz/Thuat_Toan_A_Star.git",
    ]
    for i, ref in enumerate(refs):
        y = Cm(3.2 + i * 2.05)
        add_rect(slide, Cm(1), y, Cm(31.5), Cm(1.85), RGBColor(0x07, 0x1E, 0x32))
        add_text(slide, ref, Cm(1.3), y + Cm(0.1), Cm(31), Cm(1.6),
                 font_size=14, color=C_LIGHT, align=PP_ALIGN.LEFT)

    slide_number(slide, 19)


def slide_20_thanks(prs):
    """Slide 20 – Cảm ơn (Kết luận)"""
    slide = prs.slides.add_slide(blank_layout(prs))
    full_bg(slide, C_DARK)

    # Top accent
    add_rect(slide, Cm(0), Cm(0), SLIDE_W, Cm(1.5), C_TEAL)
    add_rect(slide, Cm(0), Cm(1.5), Cm(0.8), SLIDE_H - Cm(1.5), C_ORANGE)

    add_text(slide, "Xin Cảm Ơn!",
             Cm(1.5), Cm(3), Cm(30), Cm(3.5),
             font_size=56, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT)

    add_text(slide, "Thank You for Listening",
             Cm(1.5), Cm(6.8), Cm(20), Cm(1.5),
             font_size=24, italic=True, color=C_TEAL, align=PP_ALIGN.LEFT)

    add_rect(slide, Cm(1.5), Cm(9), Cm(22), Cm(0.15), C_ORANGE)

    add_text(slide, "📌  Mọi thắc mắc xin vui lòng đặt câu hỏi",
             Cm(1.5), Cm(9.8), Cm(28), Cm(1.2),
             font_size=22, color=C_YELLOW, align=PP_ALIGN.LEFT)

    # Summary tags
    tags = ["A* Algorithm", "Python + Pygame", "Heuristic", "Pathfinding", "AI"]
    for i, tag in enumerate(tags):
        x = Cm(1.5 + i * 5.5)
        add_rect(slide, x, Cm(12.5), Cm(5), Cm(1.5), C_TEAL)
        add_text(slide, tag, x, Cm(12.5), Cm(5), Cm(1.5),
                 font_size=16, bold=True, color=C_DARK, align=PP_ALIGN.CENTER)

    add_text(slide, "Môn học: Trí Tuệ Nhân Tạo  |  Thuật toán A* – Tìm đường trong lưới 2D",
             Cm(1.5), SLIDE_H - Cm(1.8), Cm(30), Cm(1.2),
             font_size=15, color=C_GRAY, align=PP_ALIGN.LEFT)


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("Generating assets...")
    grid_img    = ASSETS_DIR / "astar_grid.png"
    formula_img = ASSETS_DIR / "astar_formula.png"
    heur_img    = ASSETS_DIR / "astar_heuristic.png"
    flow_img    = ASSETS_DIR / "astar_flow.png"
    result_img  = ASSETS_DIR / "astar_results.png"

    make_grid_image(grid_img)
    make_formula_image(formula_img)
    make_heuristic_image(heur_img)
    make_algo_flow_image(flow_img)
    make_result_summary_image(result_img)
    print("  ✓ Assets created")

    print("Building PPTX...")
    prs = new_prs()

    slide_01_title(prs)
    slide_02_problem(prs)
    slide_03_objective(prs)
    slide_04_outline(prs)
    slide_05_graph(prs)
    slide_06_pathfinding(prs, grid_img)
    slide_07_astar_intro(prs)
    slide_08_formula(prs, formula_img)
    slide_09_heuristic(prs, heur_img)
    slide_10_algo_flow(prs, flow_img)
    slide_11_design_env(prs)
    slide_12_design_arch(prs)
    slide_13_exp1(prs, grid_img)
    slide_14_exp2(prs, grid_img)
    slide_15_exp3(prs, grid_img)
    slide_16_eval(prs, result_img)
    slide_17_conclusion(prs)
    slide_18_limitation(prs)
    slide_19_references(prs)
    slide_20_thanks(prs)

    prs.save(str(OUT_PPTX))
    print(f"  ✓ Saved: {OUT_PPTX}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
