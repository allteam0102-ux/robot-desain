# -*- coding: utf-8 -*-
"""
ROBOT DESAIN KONTEN OTOMATIS  (versi 5 — output pptx (Canva) + svg (Figma))
==========================================================================
Alur: Notion (Status "Siap Desain") -> ambil gambar Pexels per slide ->
rakit desain -> kirim ke Telegram. Menghasilkan:
  1. .pptx  -> tarik/import ke Canva  (teks & gambar bisa diedit)
  2. .svg   -> tarik ke Figma / Illustrator (teks bisa diedit, gambar jadi fill)

Ringan, gratis, tanpa Aspose. Semua slide digabung jadi SATU file .svg
berjejer kiri->kanan.
"""

import os
import re
import io
import html
import time
import base64
import random
import tempfile
import traceback

import requests
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


# ======================================================================
# >>> PENGATURAN BRAND (boleh kamu ubah sesukamu) <<<
# ======================================================================
BRAND_TAGLINE = os.environ.get("BRAND_TAGLINE") or "Balancing Your Love"
FOOTER_TEXT   = os.environ.get("FOOTER_TEXT")   or "nikahinstitute.com  |  Kelas & Konseling Pranikah"
CTA_TEXT      = os.environ.get("CTA_TEXT")      or "GESER \u2192"
ACCENT_COLOR  = (os.environ.get("ACCENT_COLOR") or "7C3AED").lstrip("#")
LOGO_URL      = os.environ.get("LOGO_URL")      or ""
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Poppins"
BODY_FONT     = os.environ.get("BODY_FONT")     or "Poppins"
STYLE_HINT    = os.environ.get("STYLE_HINT")    or "candid lifestyle editorial aesthetic cinematic"

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG")  or "true").lower() == "true"

# Output "full AI" (opsional, BERBAYAR/kuota) — butuh API key
OUTPUT_AI      = (os.environ.get("OUTPUT_AI") or "auto").lower()   # true / false / auto (auto = nyala kalau ada key)
AI_PROVIDER    = (os.environ.get("AI_PROVIDER") or "gemini").lower()   # "gemini" atau "openai"
# --- OpenAI ---
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
AI_MODEL       = os.environ.get("AI_MODEL") or "gpt-image-1"
AI_SIZE        = os.environ.get("AI_SIZE") or "1024x1536"          # portrait; nanti dipotong ke rasio kanvas
AI_QUALITY     = os.environ.get("AI_QUALITY") or "medium"          # low / medium / high (makin tinggi makin mahal)
# --- Gemini (Nano Banana) ---
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL   = os.environ.get("GEMINI_MODEL") or "gemini-2.5-flash-image"

def ai_enabled():
    if OUTPUT_AI == "false":
        return False
    if AI_PROVIDER == "openai":
        return bool(OPENAI_API_KEY)
    return bool(GEMINI_API_KEY)


# ======================================================================
# 1. PENGATURAN TEKNIS (dari Secrets)
# ======================================================================
def env(name, default=None, required=False):
    val = os.environ.get(name, default)
    if required and (val is None or str(val).strip() == ""):
        raise SystemExit(f"[SETUP ERROR] Environment variable '{name}' belum diisi. Cek README.")
    return val

NOTION_TOKEN        = env("NOTION_TOKEN", required=True)
NOTION_DATABASE_ID  = env("NOTION_DATABASE_ID", required=True)
PEXELS_API_KEY      = env("PEXELS_API_KEY", required=True)
TELEGRAM_BOT_TOKEN  = env("TELEGRAM_BOT_TOKEN", required=True)
TELEGRAM_CHAT_ID    = env("TELEGRAM_CHAT_ID", required=True)

TITLE_PROPERTY   = env("TITLE_PROPERTY", "Judul")
STATUS_PROPERTY  = env("STATUS_PROPERTY", "Status")
FORMAT_PROPERTY  = env("FORMAT_PROPERTY", "Format")
CONTENT_PROPERTY = env("CONTENT_PROPERTY", "Konten")
KEYWORD_PROPERTY = env("KEYWORD_PROPERTY", "Kata Kunci Gambar")

STATUS_READY = env("STATUS_READY", "Siap Desain")
STATUS_DONE  = env("STATUS_DONE",  "Terkirim")
STATUS_ERROR = env("STATUS_ERROR", "Gagal")
STATUS_TYPE  = env("STATUS_TYPE", "select").strip().lower()

PEXELS_ORIENTATION = "portrait" if CANVAS_H > CANVAS_W else "square"

def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

ACCENT = hex_rgb(ACCENT_COLOR)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

EMU_W = CANVAS_W * 9525
EMU_H = CANVAS_H * 9525
def fx(f): return Emu(int(f * EMU_W))
def fy(f): return Emu(int(f * EMU_H))


# ======================================================================
# 2. HELPER NOTION
# ======================================================================
NOTION_BASE = "https://api.notion.com/v1"
NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_TOKEN}",
    "Notion-Version": "2022-06-28",
    "Content-Type": "application/json",
}

def notion_find_ready():
    url = f"{NOTION_BASE}/databases/{NOTION_DATABASE_ID}/query"
    if STATUS_TYPE == "status":
        flt = {"property": STATUS_PROPERTY, "status": {"equals": STATUS_READY}}
    else:
        flt = {"property": STATUS_PROPERTY, "select": {"equals": STATUS_READY}}
    resp = requests.post(url, headers=NOTION_HEADERS, json={"filter": flt}, timeout=60)
    if resp.status_code != 200:
        raise SystemExit(f"[NOTION ERROR] {resp.status_code}: {resp.text}")
    return resp.json().get("results", [])

def notion_set_status(page_id, status_value, note=None):
    url = f"{NOTION_BASE}/pages/{page_id}"
    if STATUS_TYPE == "status":
        props = {STATUS_PROPERTY: {"status": {"name": status_value}}}
    else:
        props = {STATUS_PROPERTY: {"select": {"name": status_value}}}
    props_with_note = dict(props)
    if note:
        props_with_note["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
    r = requests.patch(url, headers=NOTION_HEADERS, json={"properties": props_with_note}, timeout=60)
    if r.status_code != 200 and note:
        requests.patch(url, headers=NOTION_HEADERS, json={"properties": props}, timeout=60)

def _plain_text(rich_list):
    return "".join(part.get("plain_text", "") for part in (rich_list or []))

def read_property(props, name, kind):
    p = props.get(name)
    if not p:
        return ""
    if kind == "title":
        return _plain_text(p.get("title", [])).strip()
    if kind == "rich_text":
        return _plain_text(p.get("rich_text", [])).strip()
    if kind == "select":
        return (p.get("select") or {}).get("name", "").strip()
    if kind == "status":
        return (p.get("status") or {}).get("name", "").strip()
    return ""


# ======================================================================
# 3. PARSING ISI KONTEN
# ======================================================================
def split_slides(content_text):
    raw = (content_text or "").replace("\r\n", "\n").replace("\r", "\n")
    blocks = re.split(r"(?m)^\s*---\s*$", raw)
    return [b.strip() for b in blocks if b.strip()]

def headline_and_body(block):
    lines = block.split("\n")
    idx = next((i for i, l in enumerate(lines) if l.strip()), None)
    if idx is None:
        return "", ""
    return lines[idx].strip(), "\n".join(lines[idx + 1:]).strip()

def keyword_for(index, keyword_blocks, headline, body):
    if index < len(keyword_blocks) and keyword_blocks[index].strip():
        q = keyword_blocks[index].strip()
    elif headline:
        q = headline
    elif body:
        q = body
    else:
        q = "aesthetic minimal"
    q = re.sub(r"==", "", q)
    words = re.sub(r"[^\w\s]", " ", q).split()
    return " ".join(words[:4]) if words else "aesthetic minimal"

def parse_highlights(text):
    out = []
    for part in re.split(r"(==.+?==)", text):
        if len(part) >= 4 and part.startswith("==") and part.endswith("=="):
            out.append((part[2:-2], True))
        elif part != "":
            out.append((part, False))
    return out


# ======================================================================
# 4. GAMBAR PEXELS
# ======================================================================
def pexels_pick(query, used_ids=None):
    used_ids = used_ids or set()
    def _search(q):
        r = requests.get("https://api.pexels.com/v1/search",
                         headers={"Authorization": PEXELS_API_KEY},
                         params={"query": q, "per_page": 15, "orientation": PEXELS_ORIENTATION}, timeout=60)
        r.raise_for_status()
        return r.json().get("photos", [])
    photos = _search(f"{query} {STYLE_HINT}".strip()) or _search(query) or _search("aesthetic minimal calm")
    if not photos:
        return None, None, None
    fresh = [p for p in photos if p.get("id") not in used_ids]   # hindari gambar yang sudah dipakai
    pool = fresh if fresh else photos
    photo = random.choice(pool[:8])
    img_url = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    return requests.get(img_url, timeout=60).content, f'Foto: {photo.get("photographer", "-")} (Pexels)', photo.get("id")

def crop_to_canvas(im):
    target = CANVAS_W / CANVAS_H
    w, h = im.size
    if w / h > target:
        nw = int(h * target); x = (w - nw) // 2; im = im.crop((x, 0, x + nw, h))
    else:
        nh = int(w / target); y = (h - nh) // 2; im = im.crop((0, y, w, y + nh))
    return im.resize((CANVAS_W, CANVAS_H), Image.LANCZOS)

def prepare_image(img_bytes, out_path):
    im = crop_to_canvas(Image.open(io.BytesIO(img_bytes)).convert("RGB"))
    overlay = Image.new("RGBA", (CANVAS_W, CANVAS_H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    top_h = int(CANVAS_H * 0.16)
    for y in range(0, top_h):
        draw.line([(0, y), (CANVAS_W, y)], fill=(0, 0, 0, int(120 * (1 - y / top_h))))
    start = int(CANVAS_H * 0.42)
    for y in range(start, CANVAS_H):
        draw.line([(0, y), (CANVAS_W, y)], fill=(0, 0, 0, int(230 * (y - start) / (CANVAS_H - start))))
    Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB").save(out_path, "PNG")
    return out_path


# ======================================================================
# 4b. DESAIN FULL AI (Gemini / OpenAI) — opsional, berbayar/kuota
# ======================================================================
def ai_build_prompt(headline, body):
    return (
        f'High-end Instagram post design, portrait 4:5 vertical composition, for "{BRAND_TAGLINE}", a premium '
        f'marriage and relationship counseling brand. Cinematic, candid, editorial photography with muted '
        f'film tones and soft natural light. Elegant modern minimalist layout with sophisticated typography. '
        f'Prominent headline text, rendered exactly and spelled correctly: "{headline}". '
        f'Supporting subtext: "{body[:120]}". Purple (#{ACCENT_COLOR}) accent color. Soft dark gradient at the '
        f'bottom so the text stays readable. Clean, premium, social-media ready. All visible text in Indonesian.'
    )

def _ai_openai(headline, body, out_path):
    if not OPENAI_API_KEY:
        return None
    try:
        r = requests.post(
            "https://api.openai.com/v1/images/generations",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={"model": AI_MODEL, "prompt": ai_build_prompt(headline, body),
                  "size": AI_SIZE, "quality": AI_QUALITY, "n": 1},
            timeout=180,
        )
        if not r.ok:
            print(f"    ! OpenAI tolak ({r.status_code}): {r.text[:400]}")
            return None
        b64 = (r.json().get("data") or [{}])[0].get("b64_json")
        if not b64:
            print("    ! OpenAI tidak mengembalikan gambar.")
            return None
        im = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGB")
        crop_to_canvas(im).save(out_path, "PNG")
        return out_path
    except Exception as e:
        print("    ! OpenAI generate gagal:", e)
        return None

def _ai_gemini(headline, body, out_path):
    if not GEMINI_API_KEY:
        return None
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
        r = requests.post(
            url,
            params={"key": GEMINI_API_KEY},
            headers={"Content-Type": "application/json"},
            json={
                "contents": [{"parts": [{"text": ai_build_prompt(headline, body)}]}],
                "generationConfig": {"responseModalities": ["IMAGE"]},
            },
            timeout=180,
        )
        if not r.ok:
            print(f"    ! Gemini tolak ({r.status_code}): {r.text[:400]}")
            return None
        data = r.json()
        parts = (((data.get("candidates") or [{}])[0].get("content") or {}).get("parts")) or []
        b64 = None
        for p in parts:
            inline = p.get("inlineData") or p.get("inline_data")
            if inline and inline.get("data"):
                b64 = inline["data"]
                break
        if not b64:
            print("    ! Gemini tidak mengembalikan gambar. Resp:", str(data)[:300])
            return None
        im = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGB")
        crop_to_canvas(im).save(out_path, "PNG")
        return out_path
    except Exception as e:
        print("    ! Gemini generate gagal:", e)
        return None

def ai_generate_design(headline, body, out_path):
    if AI_PROVIDER == "openai":
        return _ai_openai(headline, body, out_path)
    return _ai_gemini(headline, body, out_path)


# ======================================================================
# 5a. RAKIT .PPTX (untuk Canva)
# ======================================================================
def text_box(slide, left, top, width, height, text, size_pt,
             bold=False, font=None, align=PP_ALIGN.LEFT, color=None, highlight=False):
    font = font or BODY_FONT
    color = color or WHITE
    box = slide.shapes.add_textbox(fx(left), fy(top), fx(width), fy(height))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in text.split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        segs = parse_highlights(line) if highlight else [(line, False)]
        for seg, is_hl in (segs or [("", False)]):
            run = p.add_run()
            run.text = seg
            f = run.font
            f.size = Pt(size_pt)
            f.bold = bold or is_hl
            f.name = font
            f.color.rgb = ACCENT if is_hl else color
    return box

def add_cta_pptx(slide):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, fx(0.66), fy(0.915), fx(0.28), fy(0.05))
    shp.fill.solid()
    shp.fill.fore_color.rgb = WHITE
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = CTA_TEXT
    f = run.font
    f.size = Pt(15); f.bold = True; f.name = BODY_FONT; f.color.rgb = ACCENT

def download_logo_path():
    if not LOGO_URL:
        return None
    try:
        data = requests.get(LOGO_URL, timeout=30).content
        path = os.path.join(tempfile.gettempdir(), f"logo_{random.randint(1, 999999)}.png")
        with open(path, "wb") as fp:
            fp.write(data)
        Image.open(path).verify()
        return path
    except Exception as e:
        print("  ! logo gagal diambil:", e)
        return None

def build_pptx(slides_data, out_path, logo_path):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W)
    prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]

    for s in slides_data:
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=prs.slide_height)
        if logo_path:
            try:
                slide.shapes.add_picture(logo_path, fx(0.06), fy(0.045), height=fy(0.045))
            except Exception:
                pass
        if BRAND_TAGLINE:
            text_box(slide, 0.52, 0.045, 0.42, 0.08, BRAND_TAGLINE, 14, bold=True, align=PP_ALIGN.RIGHT)
        if s["headline"]:
            text_box(slide, 0.07, 0.52, 0.86, 0.17, s["headline"], 38, bold=True, font=HEADLINE_FONT, highlight=True)
        if s["body"]:
            text_box(slide, 0.07, 0.70, 0.86, 0.19, s["body"], 21, bold=False, font=BODY_FONT, highlight=True)
        if FOOTER_TEXT:
            text_box(slide, 0.06, 0.93, 0.58, 0.055, FOOTER_TEXT, 11, bold=False, align=PP_ALIGN.LEFT)
        if s["total"] > 1 and CTA_TEXT:
            add_cta_pptx(slide)

    prs.save(out_path)
    return out_path


# ======================================================================
# 5b. RAKIT .SVG (untuk Figma / Illustrator) — SATU file, semua slide berjejer
# ======================================================================
def _img_data_uri(path, as_jpeg=True):
    im = Image.open(path).convert("RGB")
    buf = io.BytesIO()
    if as_jpeg:
        im.save(buf, "JPEG", quality=85)
        mime = "jpeg"
    else:
        im.save(buf, "PNG")
        mime = "png"
    b = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/{mime};base64,{b}"

def _svg_escape(t):
    return html.escape(t or "", quote=True)

def _wrap(text, max_chars):
    lines = []
    for para in (text or "").split("\n"):
        words = para.split()
        if not words:
            continue
        cur = ""
        for w in words:
            if len(cur) + len(w) + 1 <= max_chars:
                cur = (cur + " " + w).strip()
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
    return lines

def _svg_text(text, x_px, y_px, width_px, font_size, bold,
              font=None, anchor="start", highlight=False, line_gap=1.25):
    font = font or BODY_FONT
    weight = "700" if bold else "400"
    max_chars = max(6, int(width_px / (font_size * 0.55)))
    lines = _wrap(text, max_chars)
    out = []
    for i, line in enumerate(lines):
        baseline = y_px + font_size + int(i * font_size * line_gap)
        spans = ""
        segs = parse_highlights(line) if highlight else [(line, False)]
        for seg, is_hl in segs:
            fill = ACCENT_COLOR if is_hl else "FFFFFF"
            spans += f'<tspan fill="#{fill}">{_svg_escape(seg)}</tspan>'
        out.append(
            f'<text x="{x_px}" y="{baseline}" text-anchor="{anchor}" '
            f'font-family="{_svg_escape(font)}, Poppins, Arial, sans-serif" '
            f'font-size="{font_size}" font-weight="{weight}">{spans}</text>'
        )
    return "\n".join(out)

def build_svg_single(slides_data, out_path, logo_path):
    total = len(slides_data)
    W = CANVAS_W * total
    H = CANVAS_H
    logo_uri = _img_data_uri(logo_path, as_jpeg=False) if logo_path else None

    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">']
    for idx, s in enumerate(slides_data):
        xo = idx * CANVAS_W
        # background
        parts.append(
            f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{CANVAS_H}" '
            f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(s["image_path"])}"/>'
        )
        # logo
        if logo_uri:
            parts.append(
                f'<image x="{xo + int(0.06 * CANVAS_W)}" y="{int(0.05 * CANVAS_H)}" '
                f'height="{int(0.05 * CANVAS_H)}" href="{logo_uri}"/>'
            )
        # tagline (kanan atas)
        if BRAND_TAGLINE:
            parts.append(_svg_text(BRAND_TAGLINE, xo + int(0.94 * CANVAS_W), int(0.05 * CANVAS_H),
                                   int(0.4 * CANVAS_W), 26, True, anchor="end"))
        # headline
        if s["headline"]:
            parts.append(_svg_text(s["headline"], xo + int(0.07 * CANVAS_W), int(0.50 * CANVAS_H),
                                   int(0.86 * CANVAS_W), 60, True, font=HEADLINE_FONT, highlight=True))
        # body
        if s["body"]:
            parts.append(_svg_text(s["body"], xo + int(0.07 * CANVAS_W), int(0.70 * CANVAS_H),
                                   int(0.86 * CANVAS_W), 34, False, highlight=True))
        # footer
        if FOOTER_TEXT:
            parts.append(_svg_text(FOOTER_TEXT, xo + int(0.06 * CANVAS_W), int(0.925 * CANVAS_H),
                                   int(0.6 * CANVAS_W), 20, False))
        # CTA pill (carousel)
        if total > 1 and CTA_TEXT:
            px = xo + int(0.66 * CANVAS_W); py = int(0.90 * CANVAS_H)
            pw = int(0.28 * CANVAS_W); ph = int(0.06 * CANVAS_H)
            parts.append(f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" rx="{ph // 2}" fill="#FFFFFF"/>')
            parts.append(
                f'<text x="{px + pw // 2}" y="{py + int(ph * 0.66)}" text-anchor="middle" '
                f'font-family="{_svg_escape(BODY_FONT)}, Poppins, Arial, sans-serif" '
                f'font-size="26" font-weight="700" fill="#{ACCENT_COLOR}">{_svg_escape(CTA_TEXT)}</text>'
            )
    parts.append('</svg>')
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return out_path


# ======================================================================
# 6. TELEGRAM
# ======================================================================
TG_BASE = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

def _tg_call(method, data, files=None, label=""):
    try:
        r = requests.post(f"{TG_BASE}/{method}", data=data, files=files, timeout=300)
        try:
            j = r.json()
        except Exception:
            j = {}
        time.sleep(0.5)  # jeda kecil biar tidak kena batas kirim Telegram
        if not r.ok or not j.get("ok", False):
            print(f"    ! Telegram {method} {label} GAGAL (HTTP {r.status_code}): {r.text[:300]}")
            return False
        return True
    except Exception as e:
        print(f"    ! Telegram {method} {label} ERROR: {e}")
        return False

def tg_message(text):
    return _tg_call("sendMessage", {"chat_id": TELEGRAM_CHAT_ID, "text": text[:4000]}, label="msg")

def tg_photo(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendPhoto", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                        {"photo": f}, label="photo")

def tg_document(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendDocument", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                        {"document": f}, label="doc")


# ======================================================================
# 7. PROSES SATU KONTEN
# ======================================================================
def process_page(page, workdir):
    props = page["properties"]
    title = read_property(props, TITLE_PROPERTY, "title") or "Tanpa Judul"
    fmt = read_property(props, FORMAT_PROPERTY, "select") or "Single Post"
    content = read_property(props, CONTENT_PROPERTY, "rich_text")
    keywords = read_property(props, KEYWORD_PROPERTY, "rich_text")

    if not content.strip():
        raise ValueError("Kolom 'Konten' kosong.")

    slide_blocks = split_slides(content)
    if "single" in fmt.lower():
        slide_blocks = slide_blocks[:1]
    keyword_blocks = split_slides(keywords)

    print(f"  -> '{title}' | {fmt} | {len(slide_blocks)} slide")

    slides_data, credits, preview_paths = [], [], []
    used_ids = set()
    total = len(slide_blocks)
    for i, block in enumerate(slide_blocks, start=1):
        headline, body = headline_and_body(block)
        q = keyword_for(i - 1, keyword_blocks, headline, body)
        img_bytes, credit, pid = pexels_pick(q, used_ids)
        if not img_bytes:
            raise ValueError(f"Tidak ada gambar Pexels untuk kata kunci '{q}'.")
        if pid:
            used_ids.add(pid)          # supaya slide berikutnya tidak pakai gambar yang sama
        img_path = os.path.join(workdir, f"slide_{i}.png")
        prepare_image(img_bytes, img_path)
        preview_paths.append(img_path)
        if credit:
            credits.append(f"Slide {i}: {credit}")
        slides_data.append({"headline": headline, "body": body,
                            "image_path": img_path, "index": i, "total": total})

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    logo_path = download_logo_path()

    # ringkasan caption
    caption_lines = [f"🎨 {title}  ({fmt}, {total} slide)", ""]
    for i, s in enumerate(slides_data, start=1):
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:160] + "…") if len(s["body"]) > 160 else s["body"]
        caption_lines.append(f"— Slide {i}: {h}")
        if b:
            caption_lines.append(f"  {b}")
    caption_lines += ["", *credits]
    tg_message("\n".join(caption_lines))

    # preview gambar
    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")

    # OUTPUT 1: pptx (Canva)
    if OUTPUT_PPTX:
        pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
        build_pptx(slides_data, pptx_path, logo_path)
        ok = tg_document(pptx_path, caption=f"{title} — .pptx: import ke Canva ✨")
        if not ok:
            tg_message("⚠️ File .pptx gagal dikirim (cek log).")

    # OUTPUT 2: svg (Figma / Illustrator) — 1 file semua slide
    if OUTPUT_SVG:
        svg_path = os.path.join(workdir, f"{safe_name}.svg")
        build_svg_single(slides_data, svg_path, logo_path)
        size_mb = os.path.getsize(svg_path) / 1024 / 1024
        print(f"  Ukuran SVG: {size_mb:.1f} MB")
        ok = tg_document(svg_path, caption=f"{title} — .svg: tarik ke Figma (teks bisa diedit) ✨")
        if not ok:
            tg_message("⚠️ File .svg gagal dikirim (cek log).")

    # OUTPUT 3: desain full AI (Gemini / OpenAI) — 1 gambar per konten
    if ai_enabled():
        first = slides_data[0]
        ai_path = os.path.join(workdir, f"{safe_name}_ai.png")
        prov = "Gemini" if AI_PROVIDER != "openai" else "ChatGPT"
        if ai_generate_design(first["headline"], first["body"], ai_path):
            tg_photo(ai_path, caption=f"{title} — 🤖 versi FULL AI ({prov}), {CANVAS_W}x{CANVAS_H}")
        else:
            tg_message(f"ℹ️ Versi AI ({prov}) gagal dibuat (cek log / kuota & akses API).")

    return total


# ======================================================================
# 8. MAIN
# ======================================================================
def main():
    print(f"== Robot Desain Konten (v6, {CANVAS_W}x{CANVAS_H}) | pptx={OUTPUT_PPTX} svg={OUTPUT_SVG} ai={ai_enabled()} ==")
    if tg_message("✅ Tes koneksi Telegram — robot mulai jalan."):
        print("  Telegram OK (pesan tes terkirim).")
    else:
        print("  !! Telegram BERMASALAH — cek TELEGRAM_BOT_TOKEN & TELEGRAM_CHAT_ID di Secrets.")
    pages = notion_find_ready()
    print(f"Ditemukan {len(pages)} konten berstatus '{STATUS_READY}'.")
    if not pages:
        print("Tidak ada yang perlu diproses. Selesai.")
        return

    for page in pages:
        page_id = page["id"]
        workdir = tempfile.mkdtemp()
        try:
            process_page(page, workdir)
            notion_set_status(page_id, STATUS_DONE)
            print("  v Sukses & status diubah jadi:", STATUS_DONE)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            print("  x GAGAL:", err)
            traceback.print_exc()
            try:
                tg_message(f"⚠️ Gagal memproses satu konten.\n{err}")
            except Exception:
                pass
            notion_set_status(page_id, STATUS_ERROR, note=err)

    print("== Selesai ==")


if __name__ == "__main__":
    main()
