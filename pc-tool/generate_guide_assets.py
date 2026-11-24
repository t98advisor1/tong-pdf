# -*- coding: utf-8 -*-
"""
Capture real TongPDF full-screen screenshots via headless Chrome,
annotate them with high-visibility boxes, arrows, numbered badges, and Korean labels,
and generate fast-view animated GIFs for guide.html.
"""

import os
import time
import math
import threading
import subprocess
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from PIL import Image, ImageDraw, ImageFont

ROOT = r"C:\AI Agent\SynologyDrive\tong-pdf"
OUT_DIR = os.path.join(ROOT, "assets", "guide")
os.makedirs(OUT_DIR, exist_ok=True)

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
if not os.path.exists(CHROME):
    CHROME = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

FONT_BOLD = r"C:\Windows\Fonts\malgunbd.ttf"
FONT_REG = r"C:\Windows\Fonts\malgun.ttf"


def get_font(size, bold=True):
    path = FONT_BOLD if bold else FONT_REG
    return ImageFont.truetype(path, size)


class QuietHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def log_message(self, format, *args):
        pass


def start_server():
    server = ThreadingHTTPServer(("127.0.0.1", 0), QuietHandler)
    port = server.server_address[1]
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()
    return server, port


def capture(url, out_png, width=1280, height=920):
    cmd = [
        CHROME,
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        f"--window-size={width},{height}",
        "--virtual-time-budget=3500",
        f"--screenshot={out_png}",
        url,
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def draw_highlight_box(img, box, color=(220, 38, 38), fill_alpha=32, width=5, radius=14):
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    x1, y1, x2, y2 = box
    od.rounded_rectangle([x1, y1, x2, y2], radius=radius, fill=(*color, fill_alpha), outline=(*color, 255), width=width)
    return Image.alpha_composite(img.convert("RGBA"), overlay)


def draw_arrow(draw, start, end, color=(220, 38, 38), width=5, head_len=18):
    x1, y1 = start
    x2, y2 = end
    # White halo for contrast
    draw.line([x1, y1, x2, y2], fill=(255, 255, 255, 230), width=width + 4)
    draw.line([x1, y1, x2, y2], fill=(*color, 255), width=width)
    angle = math.atan2(y2 - y1, x2 - x1)
    left = (x2 - head_len * math.cos(angle - math.pi / 6), y2 - head_len * math.sin(angle - math.pi / 6))
    right = (x2 - head_len * math.cos(angle + math.pi / 6), y2 - head_len * math.sin(angle + math.pi / 6))
    draw.polygon([(x2, y2), left, right], fill=(*color, 255))


def draw_callout(img, badge_num, title, sub, pos, arrow_to=None, color=(220, 38, 38)):
    """Draws a crisp rounded label card with numbered badge + arrow pointing to target."""
    img = img.convert("RGBA")
    draw = ImageDraw.Draw(img)
    f_num = get_font(19, bold=True)
    f_title = get_font(19, bold=True)
    f_sub = get_font(15, bold=False)

    tb = draw.textbbox((0, 0), title, font=f_title)
    tw = tb[2] - tb[0]
    sw = 0
    if sub:
        sb = draw.textbbox((0, 0), sub, font=f_sub)
        sw = sb[2] - sb[0]

    card_w = max(tw, sw) + 64
    card_h = 62 if sub else 44
    cx, cy = pos
    # Keep inside image bounds
    cx = max(12, min(cx, img.width - card_w - 12))
    cy = max(12, min(cy, img.height - card_h - 12))

    if arrow_to is not None:
        # Compute best anchor point on card border toward arrow_to
        mid_x = cx + card_w / 2
        mid_y = cy + card_h / 2
        tx, ty = arrow_to
        if abs(tx - mid_x) > abs(ty - mid_y):
            sx = cx + card_w if tx > mid_x else cx
            sy = mid_y
        else:
            sx = mid_x
            sy = cy + card_h if ty > mid_y else cy
        draw_arrow(draw, (sx, sy), (tx, ty), color=color, width=5)

    # Shadow
    draw.rounded_rectangle([cx + 3, cy + 4, cx + card_w + 3, cy + card_h + 4], radius=12, fill=(15, 23, 42, 70))
    # Card background
    draw.rounded_rectangle([cx, cy, cx + card_w, cy + card_h], radius=12, fill=(255, 255, 255, 248), outline=(*color, 255), width=3)

    # Number circle badge
    bx, by = cx + 12, cy + (card_h - 32) // 2
    draw.ellipse([bx, by, bx + 32, by + 32], fill=(*color, 255))
    nb = draw.textbbox((0, 0), str(badge_num), font=f_num)
    nw, nh = nb[2] - nb[0], nb[3] - nb[1]
    draw.text((bx + (32 - nw) / 2, by + (32 - nh) / 2 - 2), str(badge_num), font=f_num, fill=(255, 255, 255, 255))

    # Text
    tx_start = cx + 52
    if sub:
        draw.text((tx_start, cy + 8), title, font=f_title, fill=(15, 23, 42, 255))
        draw.text((tx_start, cy + 34), sub, font=f_sub, fill=(71, 85, 105, 255))
    else:
        draw.text((tx_start, cy + 10), title, font=f_title, fill=(15, 23, 42, 255))

    return img


def add_top_caption_banner(img, step_label, caption, step_idx, total_steps, accent=(29, 78, 216)):
    """Adds a clean top caption bar for fast-view GIF frames."""
    banner_h = 74
    w, h = img.size
    canvas = Image.new("RGBA", (w, h + banner_h), (241, 245, 249, 255))
    canvas.paste(img.convert("RGBA"), (0, banner_h))
    draw = ImageDraw.Draw(canvas)

    draw.rectangle([0, 0, w, banner_h], fill=(15, 23, 42, 255))
    # Progress bar along bottom of banner
    prog_w = int(w * (step_idx / total_steps))
    draw.rectangle([0, banner_h - 6, w, banner_h], fill=(51, 65, 85, 255))
    draw.rectangle([0, banner_h - 6, prog_w, banner_h], fill=(250, 204, 21, 255))

    f_badge = get_font(18, bold=True)
    f_cap = get_font(23, bold=True)

    # Pill badge
    draw.rounded_rectangle([20, 16, 150, 52], radius=18, fill=(*accent, 255))
    draw.text((34, 22), step_label, font=f_badge, fill=(255, 255, 255, 255))
    draw.text((168, 20), caption, font=f_cap, fill=(255, 255, 255, 255))

    return canvas


def build_pc_workflow_gif(out_gif):
    """Generates a clean visual 4-step animated GIF showing the Windows portable EXE workflow."""
    w, h = 960, 540
    steps = [
        ("STEP 1 / 4", "TongPDF-PC.exe 파일 1개 다운로드 (설치 없음)",
         "웹사이트 상단 노란색 버튼에서 TongPDF-PC.exe를 받아 바탕화면에 둡니다.",
         "⬇  TongPDF-PC.exe (무설치 포터블 · 32MB)", (29, 78, 216)),
        ("STEP 2 / 4", "바탕화면에서 더블클릭 실행",
         "첫 실행 시 Windows 'PC 보호' 창이 뜨면 [추가 정보] → [실행]을 누릅니다.",
         "🖱️  더블클릭 → 통PDF 전용 창 즉시 실행 (관리자 권한 불필요)", (2, 132, 199)),
        ("STEP 3 / 4", "한글·한쇼·PPT 파일 또는 폴더를 통째로 드래그!",
         ".hwp / .hwpx / .show / .ppt / .pptx / .docx 수십 개를 한꺼번에 끌어다 놓습니다.",
         "📂  시험지.hwpx + 수업자료.show + 발표.pptx  ➜  창으로 드래그", (5, 150, 105)),
        ("STEP 4 / 4", "[⚡ 일괄 PDF 변환 시작] 클릭 → 원본 100% PDF 완성!",
         "PC의 한컴오피스·파워포인트가 백그라운드에서 자동 변환 (외부 전송 없음)",
         "✅  수식 · 표 · 결재란 · 한쇼 슬라이드 원본 100% 그대로 PDF 저장 완료!", (220, 38, 38)),
    ]
    frames = []
    f_tag = get_font(20, bold=True)
    f_title = get_font(30, bold=True)
    f_desc = get_font(20, bold=False)
    f_box = get_font(22, bold=True)
    f_small = get_font(16, bold=False)

    for idx, (tag, title, desc, highlight, color) in enumerate(steps, start=1):
        im = Image.new("RGB", (w, h), (248, 250, 252))
        d = ImageDraw.Draw(im)

        # Top header bar
        d.rectangle([0, 0, w, 72], fill=(15, 23, 42))
        d.text((28, 20), "통PDF (TongPDF-PC.exe) · 교무실 Windows PC 빠른 사용법", font=get_font(22, True), fill=(255, 255, 255))

        # Progress dots
        for j in range(1, 5):
            cx = 760 + (j - 1) * 44
            fill = (250, 204, 21) if j <= idx else (71, 85, 105)
            d.ellipse([cx, 24, cx + 24, 48], fill=fill)
            d.text((cx + 7, 26), str(j), font=get_font(15, True), fill=(15, 23, 42) if j <= idx else (203, 213, 225))

        # Main card
        d.rounded_rectangle([42, 108, w - 42, h - 46], radius=20, fill=(255, 255, 255), outline=(203, 213, 225), width=2)
        d.rounded_rectangle([76, 142, 236, 184], radius=21, fill=color)
        d.text((100, 150), tag, font=f_tag, fill=(255, 255, 255))

        d.text((76, 206), title, font=f_title, fill=(15, 23, 42))
        d.text((76, 260), desc, font=f_desc, fill=(71, 85, 105))

        # Highlight action box
        d.rounded_rectangle([76, 324, w - 76, 412], radius=14, fill=(239, 246, 255), outline=color, width=3)
        d.text((106, 353), highlight, font=f_box, fill=color)

        # Bottom tip
        d.text((76, 444), "※ 한글 보안 팝업이 뜨면 프로그램 우측 상단 [한글 보안 팝업 해제 설정] 또는 [모두 허용]을 1회 눌러주세요.", font=f_small, fill=(100, 116, 139))

        frames.append(im.convert("P", palette=Image.ADAPTIVE))

    frames[0].save(out_gif, save_all=True, append_images=frames[1:], duration=2400, loop=0, optimize=True)


def main():
    server, port = start_server()
    base_url = f"http://127.0.0.1:{port}/index.html"
    raw_files = {}

    try:
        for tab in ("doc2pdf", "imgPdf", "mergeSplit", "pcTool"):
            raw_path = os.path.join(OUT_DIR, f"_raw_{tab}.png")
            capture(f"{base_url}?tab={tab}", raw_path, width=1280, height=920)
            raw_files[tab] = raw_path
    finally:
        server.shutdown()

    # 1) Annotate Step 1: Doc -> PDF
    im1 = Image.open(raw_files["doc2pdf"]).convert("RGBA")
    im1 = draw_highlight_box(im1, (90, 100, 1190, 248), color=(29, 78, 216), fill_alpha=18)
    im1 = draw_highlight_box(im1, (114, 444, 1166, 666), color=(220, 38, 38), fill_alpha=18)
    im1 = draw_highlight_box(im1, (114, 678, 1166, 748), color=(5, 150, 105), fill_alpha=20)
    im1 = draw_callout(im1, 1, "교무실 PC 완벽 변환기(.exe)", "한쇼(.show)·복잡한 한글은 여기서 다운로드", (760, 260), arrow_to=(940, 200), color=(29, 78, 216))
    im1 = draw_callout(im1, 2, "문서 파일 여러 개 드래그!", ".hwpx · .hwp · .pptx · .docx 동시 선택", (760, 464), arrow_to=(620, 540), color=(220, 38, 38))
    im1 = draw_callout(im1, 3, "용지·글자 선택 후 일괄 변환", "Ctrl+Enter 변환 · Ctrl+Shift+D ZIP 저장", (730, 774), arrow_to=(860, 742), color=(5, 150, 105))
    step1_out = os.path.join(OUT_DIR, "step1_doc2pdf.png")
    im1.convert("RGB").save(step1_out, optimize=True)

    # 2) Annotate Step 2: Image <-> PDF
    im2 = Image.open(raw_files["imgPdf"]).convert("RGBA")
    im2 = draw_highlight_box(im2, (114, 440, 632, 742), color=(220, 38, 38), fill_alpha=18)
    im2 = draw_highlight_box(im2, (648, 440, 1166, 742), color=(2, 132, 199), fill_alpha=18)
    im2 = draw_callout(im2, 1, "이미지 여러 장 → PDF 1개", "학습지 사진·스캔본을 A4 PDF로 묶기", (130, 768), arrow_to=(350, 738), color=(220, 38, 38))
    im2 = draw_callout(im2, 2, "PDF → 페이지별 고화질 PNG", "해상도 선택 후 ZIP 파일로 한 번에 저장", (690, 768), arrow_to=(910, 738), color=(2, 132, 199))
    step2_out = os.path.join(OUT_DIR, "step2_img_pdf.png")
    im2.convert("RGB").save(step2_out, optimize=True)

    # 3) Annotate Step 3: PDF Merge & Split
    im3 = Image.open(raw_files["mergeSplit"]).convert("RGBA")
    im3 = draw_highlight_box(im3, (114, 440, 632, 742), color=(220, 38, 38), fill_alpha=18)
    im3 = draw_highlight_box(im3, (648, 440, 1166, 742), color=(5, 150, 105), fill_alpha=18)
    im3 = draw_callout(im3, 1, "PDF 여러 개 순서대로 합치기", "▲▼ 버튼으로 순서 변경 후 1개로 병합", (125, 768), arrow_to=(350, 738), color=(220, 38, 38))
    im3 = draw_callout(im3, 2, "원하는 쪽만 추출 / 1장씩 분리", "예: 1-3, 5 입력 또는 전체 1쪽씩 ZIP 저장", (680, 768), arrow_to=(910, 738), color=(5, 150, 105))
    step3_out = os.path.join(OUT_DIR, "step3_merge_split.png")
    im3.convert("RGB").save(step3_out, optimize=True)

    # 4) Annotate Step 4: PC EXE Tool Tab
    im4 = Image.open(raw_files["pcTool"]).convert("RGBA")
    im4 = draw_highlight_box(im4, (944, 376, 1166, 422), color=(220, 38, 38), fill_alpha=24)
    im4 = draw_highlight_box(im4, (114, 442, 1166, 610), color=(29, 78, 216), fill_alpha=14)
    im4 = draw_highlight_box(im4, (114, 650, 1166, 902), color=(5, 150, 105), fill_alpha=12)
    im4 = draw_callout(im4, 1, "무설치 실행파일 다운로드", "바탕화면·USB 어디서나 바로 실행", (545, 294), arrow_to=(944, 398), color=(220, 38, 38))
    im4 = draw_callout(im4, 2, "3단계 간편 사용법 확인", "다운로드 → 더블클릭 → 폴더째 드래그", (760, 560), arrow_to=(620, 520), color=(29, 78, 216))
    im4 = draw_callout(im4, 3, "웹 vs PC용 포맷 지원 비교", "한쇼(.show)·구형 .ppt는 PC용 .exe 권장", (740, 826), arrow_to=(740, 790), color=(5, 150, 105))
    step4_out = os.path.join(OUT_DIR, "step4_pc_exe.png")
    im4.convert("RGB").save(step4_out, optimize=True)

    # 5) Build Fast-View Web GIF from annotated frames (resized to crisp 960px wide)
    gif_frames = []
    gif_meta = [
        (step1_out, "STEP 1 / 4", "문서(.hwpx·.hwp·.pptx·.docx) 드래그 → 일괄 PDF 변환", (220, 38, 38)),
        (step2_out, "STEP 2 / 4", "이미지 여러 장 → PDF 합치기 · PDF → 고화질 PNG 저장", (2, 132, 199)),
        (step3_out, "STEP 3 / 4", "PDF 여러 개 합치기 · 원하는 페이지만 쏙쏙 나누기", (5, 150, 105)),
        (step4_out, "STEP 4 / 4", "한쇼(.show)·복잡한 한글은 PC용 무설치(.exe)로 100% 변환", (29, 78, 216)),
    ]
    for i, (path, lbl, cap, acc) in enumerate(gif_meta, start=1):
        fr = Image.open(path).convert("RGBA")
        new_w = 960
        new_h = int(fr.height * (new_w / fr.width))
        fr = fr.resize((new_w, new_h), Image.Resampling.LANCZOS)
        fr = add_top_caption_banner(fr, lbl, cap, i, len(gif_meta), accent=acc)
        gif_frames.append(fr.convert("RGB").convert("P", palette=Image.ADAPTIVE))

    web_gif = os.path.join(OUT_DIR, "quick_demo_web.gif")
    gif_frames[0].save(web_gif, save_all=True, append_images=gif_frames[1:], duration=2400, loop=0, optimize=True)

    # 6) Build Fast-View PC EXE Workflow GIF
    pc_gif = os.path.join(OUT_DIR, "quick_demo_pc.gif")
    build_pc_workflow_gif(pc_gif)

    # Cleanup temporary raw captures
    for rp in raw_files.values():
        if os.path.exists(rp):
            os.remove(rp)

    print("GENERATED GUIDE ASSETS OK:")
    for fn in sorted(os.listdir(OUT_DIR)):
        fp = os.path.join(OUT_DIR, fn)
        print(f" - {fn} ({os.path.getsize(fp):,} bytes)")


if __name__ == "__main__":
    main()
