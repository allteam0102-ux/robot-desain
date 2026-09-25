# -*- coding: utf-8 -*-
"""
ROBOT DESAIN KONTEN OTOMATIS  (versi 4 — +output PSD teks-live via Aspose)
=========================================================================
Menghasilkan DUA output sekaligus:
  1. .pptx  -> untuk diedit di Canva (gratis)
  2. .psd   -> untuk diedit di Photoshop, TEKS MASIH BISA DIKETIK ULANG
              (pakai library Aspose.PSD — berbayar; mode gratis ada watermark)

Kalau Aspose tidak terpasang / gagal, output .psd otomatis dilewati dan
.pptx tetap terkirim (jaring pengaman).

Atur output lewat env: OUTPUT_PPTX (default true), OUTPUT_PSD (default true).
Watermark hilang jika ASPOSE_METERED_PUBLIC & ASPOSE_METERED_PRIVATE diisi.
"""

import os
import re
import io
import random
import zipfile
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
STYLE_HINT    = os.environ.get("STYLE_HINT")    or "aesthetic candid cinematic natural light"

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

# Output mana yang dibuat
OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_PSD  = (os.environ.get("OUTPUT_PSD")  or "true").lower() == "true"


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
# 1b. ASPOSE (untuk PSD). Diimpor dengan aman — kalau gagal, PSD dilewati.
# ======================================================================
ASPOSE_OK = False
if OUTPUT_PSD:
    try:
        from aspose.psd import Image as AsImage, Rectangle as AsRect, Color as AsColor, Graphics as AsGraphics
        from aspose.psd.fileformats.psd import PsdImage
        ASPOSE_OK = True
    except Exception as _e:
        print("  ! Aspose.PSD tidak tersedia, output PSD dilewati:", _e)

def apply_aspose_license():
    if not ASPOSE_OK:
        return
    # 1) File license (Temporary / Full) dari ASPOSE_LICENSE_PATH
    lic_path = os.environ.get("ASPOSE_LICENSE_PATH", "")
    if lic_path and os.path.exists(lic_path) and os.path.getsize(lic_path) > 0:
        try:
            from aspose.psd import License
            License().set_license(lic_path)
            print("  Aspose: file license aktif.")
            return
        except Exception as e:
            print("  ! Gagal set file license:", e)
    # 2) Metered keys (pay-as-you-use)
    pub = os.environ.get("ASPOSE_METERED_PUBLIC", "")
    priv = os.environ.get("ASPOSE_METERED_PRIVATE", "")
    if pub and priv:
        try:
            from aspose.psd import Metered
            Metered().set_metered_key(pub, priv)
            print("  Aspose: metered license aktif.")
        except Exception as e:
            print("  ! Gagal set metered license:", e)


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

def strip_marks(text):
    return text.replace("==", "")


# ======================================================================
# 4. GAMBAR PEXELS
# ======================================================================
def pexels_pick(query):
    def _search(q):
        r = requests.get("https://api.pexels.com/v1/search",
                         headers={"Authorization": PEXELS_API_KEY},
                         params={"query": q, "per_page": 10, "orientation": PEXELS_ORIENTATION}, timeout=60)
        r.raise_for_status()
        return r.json().get("photos", [])
    photos = _search(f"{query} {STYLE_HINT}".strip()) or _search(query) or _search("aesthetic minimal calm")
    if not photos:
        return None, None
    photo = random.choice(photos[:6])
    img_url = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    return requests.get(img_url, timeout=60).content, f'Foto: {photo.get("photographer", "-")} (Pexels)'

def prepare_image(img_bytes, out_path):
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    target = CANVAS_W / CANVAS_H
    w, h = im.size
    if w / h > target:
        nw = int(h * target); x = (w - nw) // 2; im = im.crop((x, 0, x + nw, h))
    else:
        nh = int(w / target); y = (h - nh) // 2; im = im.crop((0, y, w, y + nh))
    im = im.resize((CANVAS_W, CANVAS_H), Image.LANCZOS)

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
# 5b. RAKIT .PSD (untuk Photoshop, teks bisa diketik ulang) — EKSPERIMENTAL
# ======================================================================
def _psd_white():
    try:
        return AsColor.from_argb(255, 255, 255, 255)
    except Exception:
        try:
            return AsColor.white
        except Exception:
            return None

def _psd_add_text(psd, text, x, y, w, h, size, x_off=0):
    if not text:
        return
    rect = AsRect(int(x * CANVAS_W) + x_off, int(y * CANVAS_H), int(w * CANVAS_W), int(h * CANVAS_H))
    tl = psd.add_text_layer(strip_marks(text.replace("\n", " ")), rect)
    try:
        td = tl.text_data
        white = _psd_white()
        for portion in td.items:
            try:
                portion.style.font_size = float(size)
            except Exception:
                pass
            try:
                if white is not None:
                    portion.style.fill_color = white
            except Exception:
                pass
        td.update_layer_data()
    except Exception as te:
        print("    (gaya teks PSD dilewati:", te, ")")

def build_psd_single(slides_data, workdir, safe_name):
    """Semua slide dalam SATU file .psd, disusun berjejer kiri->kanan."""
    if not ASPOSE_OK:
        return None
    apply_aspose_license()
    total = len(slides_data)
    if total == 0:
        return None
    try:
        psd = PsdImage(CANVAS_W * total, CANVAS_H)   # kanvas lebar = semua slide berjejer
        for idx, s in enumerate(slides_data):
            x_off = idx * CANVAS_W
            # background slide ini (jadi 1 layer)
            try:
                layer = psd.add_regular_layer()
                g = AsGraphics(layer)
                with AsImage.load(s["image_path"]) as raster:
                    g.draw_image(raster, AsRect(x_off, 0, CANVAS_W, CANVAS_H))
            except Exception as be:
                print("    (background PSD dilewati:", be, ")")
            # teks slide ini (layer teks, bisa diketik ulang di Photoshop)
            _psd_add_text(psd, s["headline"], 0.07, 0.52, 0.86, 0.17, 52, x_off)
            _psd_add_text(psd, s["body"],     0.07, 0.70, 0.86, 0.19, 30, x_off)
        out = os.path.join(workdir, f"{safe_name}.psd")
        psd.save(out)
        return out
    except Exception as e:
        print("    ! PSD gagal:", e)
        return None


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
                      files={"document": f}, timeout=180)


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
        build_pptx(slides_data, pptx_path)
        tg_document(pptx_path, caption=f"{title} — .pptx: import ke Canva ✨")

    # OUTPUT 2: psd (Photoshop) — SATU file berisi semua slide berjejer
    if OUTPUT_PSD:
        psd_path = build_psd_single(slides_data, workdir, safe_name)
        if psd_path:
            tg_document(psd_path, caption=f"{title} — .psd 1 file ({total} slide berjejer, teks bisa diedit di Photoshop)")
        elif ASPOSE_OK:
            tg_message("ℹ️ PSD gagal dibuat untuk konten ini (cek log GitHub Actions).")

    return total


# ======================================================================
# 8. MAIN
# ======================================================================
def main():
    print(f"== Robot Desain Konten (v4, {CANVAS_W}x{CANVAS_H}) | pptx={OUTPUT_PPTX} psd={OUTPUT_PSD and ASPOSE_OK} ==")
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
