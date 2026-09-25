# -*- coding: utf-8 -*-
"""
ROBOT DESAIN KONTEN OTOMATIS  (versi 3 — 1080x1350 + gambar aesthetic)
=====================================================================
Alur: Notion (Status "Siap Desain") -> ambil gambar Pexels per slide ->
rakit jadi .pptx bergaya brand -> kirim ke Telegram -> import ke Canva -> edit.

Kalau mau ganti tagline/footer/warna/gaya gambar, cukup ubah teks di
bagian ">>> PENGATURAN BRAND <<<" di bawah ini. Tidak perlu sentuh yang lain.
"""

import os
import re
import io
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
BRAND_TAGLINE = os.environ.get("BRAND_TAGLINE") or "Balancing Your Love"     # kanan atas
FOOTER_TEXT   = os.environ.get("FOOTER_TEXT")   or "nikahinstitute.com  |  Kelas & Konseling Pranikah"  # kiri bawah
CTA_TEXT      = os.environ.get("CTA_TEXT")      or "GESER \u2192"            # tombol kanan bawah (khusus carousel)
ACCENT_COLOR  = (os.environ.get("ACCENT_COLOR") or "7C3AED").lstrip("#")     # warna aksen (hex tanpa #)
LOGO_URL      = os.environ.get("LOGO_URL")      or ""                        # URL logo PNG transparan (opsional)
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Poppins"
BODY_FONT     = os.environ.get("BODY_FONT")     or "Poppins"

# Pengarah gaya gambar: kata ini otomatis ditambahkan ke pencarian Pexels
# supaya hasilnya lebih aesthetic/editorial, bukan stok promosi yang kaku.
STYLE_HINT    = os.environ.get("STYLE_HINT")    or "aesthetic candid cinematic natural light"

# Ukuran kanvas (px). Default 1080x1350 (portrait 4:5, feed IG).
CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)


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

# Orientasi Pexels: portrait kalau kanvas lebih tinggi dari lebar, selain itu square
PEXELS_ORIENTATION = "portrait" if CANVAS_H > CANVAS_W else "square"

def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

ACCENT = hex_rgb(ACCENT_COLOR)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

# Konversi px -> EMU (1 px @96dpi = 9525 EMU)
EMU_W = CANVAS_W * 9525
EMU_H = CANVAS_H * 9525
def fx(f): return Emu(int(f * EMU_W))   # posisi/lebar (arah horizontal)
def fy(f): return Emu(int(f * EMU_H))   # posisi/tinggi (arah vertikal)


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
# 4. GAMBAR PEXELS (dibuat lebih aesthetic)
# ======================================================================
def pexels_pick(query):
    def _search(q):
        r = requests.get("https://api.pexels.com/v1/search",
                         headers={"Authorization": PEXELS_API_KEY},
                         params={"query": q, "per_page": 10, "orientation": PEXELS_ORIENTATION}, timeout=60)
        r.raise_for_status()
        return r.json().get("photos", [])
    # 1) coba kata kunci + pengarah gaya  2) kata kunci polos  3) cadangan generik
    photos = _search(f"{query} {STYLE_HINT}".strip()) or _search(query) or _search("aesthetic minimal calm")
    if not photos:
        return None, None
    photo = random.choice(photos[:6])
    img_url = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    return requests.get(img_url, timeout=60).content, f'Foto: {photo.get("photographer", "-")} (Pexels)'

def prepare_image(img_bytes, out_path):
    """Potong ke rasio kanvas + gelapkan atas (logo) & bawah (teks)."""
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    target = CANVAS_W / CANVAS_H
    w, h = im.size
    if w / h > target:                       # terlalu lebar -> potong sisi kiri-kanan
        nw = int(h * target); x = (w - nw) // 2; im = im.crop((x, 0, x + nw, h))
    else:                                    # terlalu tinggi -> potong atas-bawah
        nh = int(w / target); y = (h - nh) // 2; im = im.crop((0, y, w, y + nh))
    im = im.resize((CANVAS_W, CANVAS_H), Image.LANCZOS)

    overlay = Image.new("RGBA", (CANVAS_W, CANVAS_H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    top_h = int(CANVAS_H * 0.16)             # scrim atas tipis
    for y in range(0, top_h):
        draw.line([(0, y), (CANVAS_W, y)], fill=(0, 0, 0, int(120 * (1 - y / top_h))))
    start = int(CANVAS_H * 0.42)             # gradasi bawah tebal
    for y in range(start, CANVAS_H):
        draw.line([(0, y), (CANVAS_W, y)], fill=(0, 0, 0, int(230 * (y - start) / (CANVAS_H - start))))
    Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB").save(out_path, "PNG")
    return out_path


# ======================================================================
# 5. RAKIT .PPTX BERGAYA BRAND
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

def add_cta(slide):
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

def download_logo():
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

def build_pptx(slides_data, out_path):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W)
    prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]
    logo_path = download_logo()

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
            add_cta(slide)

    prs.save(out_path)
    return out_path


# ======================================================================
# 6. TELEGRAM
# ======================================================================
TG_BASE = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

def tg_message(text):
    requests.post(f"{TG_BASE}/sendMessage",
                  data={"chat_id": TELEGRAM_CHAT_ID, "text": text[:4000]}, timeout=60)

def tg_photo(path, caption=""):
    with open(path, "rb") as f:
        requests.post(f"{TG_BASE}/sendPhoto",
                      data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                      files={"photo": f}, timeout=120)

def tg_document(path, caption=""):
    with open(path, "rb") as f:
        requests.post(f"{TG_BASE}/sendDocument",
                      data={"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]},
                      files={"document": f}, timeout=120)


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
    total = len(slide_blocks)
    for i, block in enumerate(slide_blocks, start=1):
        headline, body = headline_and_body(block)
        q = keyword_for(i - 1, keyword_blocks, headline, body)
        img_bytes, credit = pexels_pick(q)
        if not img_bytes:
            raise ValueError(f"Tidak ada gambar Pexels untuk kata kunci '{q}'.")
        img_path = os.path.join(workdir, f"slide_{i}.png")
        prepare_image(img_bytes, img_path)
        preview_paths.append(img_path)
        if credit:
            credits.append(f"Slide {i}: {credit}")
        slides_data.append({"headline": headline, "body": body,
                            "image_path": img_path, "index": i, "total": total})

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
    build_pptx(slides_data, pptx_path)

    caption_lines = [f"🎨 {title}  ({fmt}, {total} slide)", ""]
    for i, s in enumerate(slides_data, start=1):
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:160] + "…") if len(s["body"]) > 160 else s["body"]
        caption_lines.append(f"— Slide {i}: {h}")
        if b:
            caption_lines.append(f"  {b}")
    caption_lines += ["", *credits, "",
                      "➡️ Buka Canva → Upload → pilih file .pptx ini → tiap slide jadi halaman yang bisa diedit."]
    tg_message("\n".join(caption_lines))

    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")
    tg_document(pptx_path, caption=f"{title} — import file ini ke Canva ✨")
    return total


# ======================================================================
# 8. MAIN
# ======================================================================
def main():
    print(f"== Robot Desain Konten (v3, {CANVAS_W}x{CANVAS_H}): cek Notion ==")
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
