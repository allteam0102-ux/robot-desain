# -*- coding: utf-8 -*-
"""
ROBOT DESAIN NIKAH INSTITUTE  (gaya premium ungu)
=================================================
Notion (Status "Siap Desain") -> gambar (Unsplash utama, Pexels cadangan) ->
rakit desain full-bleed -> Telegram.
Output: .pptx (Canva) + .svg gabungan + .svg per slide.
Selesai: Status -> Terkirim & ikon halaman -> ✅.

BARU di versi ini:
- Teks otomatis ditaruh di area gambar yang paling KOSONG:
  atas / bawah / kiri / kanan / tengah (bukan cuma atas-bawah).
- Slide CTA khusus: ketik  [cta]  di awal slide terakhir di Notion ->
  robot bikin layout mockup HP + kotak ungu CTA (teksnya tetap bisa diedit).
- Gambar dibiaskan ke gaya aesthetic/editorial luar negeri + sentuhan
  film-tone halus biar seragam & premium (bukan norak).
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
from PIL import Image, ImageDraw, ImageStat, ImageEnhance
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
# >>> PENGATURAN BRAND <<<
# ======================================================================
BRAND_TAGLINE = os.environ.get("BRAND_TAGLINE") or "Balancing Your Love"
FOOTER_TEXT   = os.environ.get("FOOTER_TEXT")   or "nikahinstitute.com  |  Kelas & Konseling Pranikah"
CTA_TEXT      = os.environ.get("CTA_TEXT")      or "GESER →"
ACCENT_COLOR  = (os.environ.get("ACCENT_COLOR") or "7C3AED").lstrip("#")
LOGO_URL      = os.environ.get("LOGO_URL")      or ""
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Poppins"
BODY_FONT     = os.environ.get("BODY_FONT")     or "Poppins"
CTA_HANDLE    = os.environ.get("CTA_HANDLE")    or "@nikahinstitute"
# gaya gambar: dibiaskan ke aesthetic/editorial luar negeri, premium, bukan norak
STYLE_HINT    = os.environ.get("STYLE_HINT")    or "editorial lifestyle film"
FILM_GRADE    = (os.environ.get("FILM_GRADE") or "true").lower() == "true"

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG")  or "true").lower() == "true"
OUTPUT_SVG_PER_SLIDE = (os.environ.get("OUTPUT_SVG_PER_SLIDE") or "true").lower() == "true"


# ======================================================================
# 1. TEKNIS (Secrets)
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
UNSPLASH_ACCESS_KEY = os.environ.get("UNSPLASH_ACCESS_KEY", "")   # opsional (utama kalau diisi)

TITLE_PROPERTY   = env("TITLE_PROPERTY", "Judul")
STATUS_PROPERTY  = env("STATUS_PROPERTY", "Status")
FORMAT_PROPERTY  = env("FORMAT_PROPERTY", "Format")
CONTENT_PROPERTY = env("CONTENT_PROPERTY", "Konten")
KEYWORD_PROPERTY = env("KEYWORD_PROPERTY", "Kata Kunci Gambar")

STATUS_READY = env("STATUS_READY", "Siap Desain")
STATUS_DONE  = env("STATUS_DONE",  "Terkirim")
STATUS_ERROR = env("STATUS_ERROR", "Gagal")
STATUS_TYPE  = env("STATUS_TYPE", "select").strip().lower()
DONE_EMOJI   = env("DONE_EMOJI", "✅")

PEXELS_ORIENTATION = "portrait" if CANVAS_H > CANVAS_W else "square"

def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

def hex_tuple(h):
    h = h.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

ACCENT = hex_rgb(ACCENT_COLOR)
ACCENT_T = hex_tuple(ACCENT_COLOR)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

EMU_W = CANVAS_W * 9525
EMU_H = CANVAS_H * 9525
def fx(f): return Emu(int(f * EMU_W))
def fy(f): return Emu(int(f * EMU_H))

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/122.0 Safari/537.36"}


# ======================================================================
# 2. NOTION
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

def notion_set_status(page_id, status_value, note=None, icon_emoji=None):
    url = f"{NOTION_BASE}/pages/{page_id}"
    if STATUS_TYPE == "status":
        props = {STATUS_PROPERTY: {"status": {"name": status_value}}}
    else:
        props = {STATUS_PROPERTY: {"select": {"name": status_value}}}
    props_with_note = dict(props)
    if note:
        props_with_note["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
    body = {"properties": props_with_note}
    if icon_emoji:
        body["icon"] = {"type": "emoji", "emoji": icon_emoji}
    r = requests.patch(url, headers=NOTION_HEADERS, json=body, timeout=60)
    if r.status_code != 200:
        body2 = {"properties": props}
        if icon_emoji:
            body2["icon"] = {"type": "emoji", "emoji": icon_emoji}
        requests.patch(url, headers=NOTION_HEADERS, json=body2, timeout=60)

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

def get_title(props):
    t = read_property(props, TITLE_PROPERTY, "title")
    if t:
        return t
    for _n, p in props.items():
        if isinstance(p, dict) and p.get("type") == "title":
            return _plain_text(p.get("title", [])).strip()
    return ""


# ======================================================================
# 3. PARSING
# ======================================================================
def split_slides(content_text):
    raw = (content_text or "").replace("\r\n", "\n").replace("\r", "\n")
    blocks = re.split(r"(?m)^\s*(?:[-–—]{2,}|[–—])\s*$", raw)
    return [b.strip() for b in blocks if b.strip()]

def headline_and_body(block):
    lines = block.split("\n")
    idx = next((i for i, l in enumerate(lines) if l.strip()), None)
    if idx is None:
        return "", ""
    return lines[idx].strip(), "\n".join(lines[idx + 1:]).strip()

CTA_MARK = re.compile(r"^\s*\[\s*cta\s*\]\s*", re.IGNORECASE)

def is_cta_block(block):
    for l in (block or "").split("\n"):
        if l.strip():
            return bool(CTA_MARK.match(l))
    return False

def cta_text_of(block):
    """Buang penanda [cta], kembalikan teks ajakan (boleh beberapa baris)."""
    lines = [l for l in (block or "").split("\n")]
    cleaned = []
    for l in lines:
        if not l.strip():
            if cleaned:
                cleaned.append("")
            continue
        cleaned.append(CTA_MARK.sub("", l).rstrip())
    return "\n".join(cleaned).strip()

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
# 4. GAMBAR (Unsplash utama, Pexels cadangan)
# ======================================================================
FALLBACK_QUERIES = ["editorial lifestyle couple", "soft film tone portrait", "minimal european street"]

def _short_q(query, add_hint=True):
    words = (query or "").split()[:3]
    q = " ".join(words) if words else "aesthetic portrait"
    if add_hint and STYLE_HINT:
        q = (q + " " + " ".join(STYLE_HINT.split()[:2])).strip()
    return q

def pexels_pick(query, used_ids):
    def _search(q):
        r = requests.get("https://api.pexels.com/v1/search",
                         headers={"Authorization": PEXELS_API_KEY},
                         params={"query": q, "per_page": 15, "orientation": PEXELS_ORIENTATION}, timeout=60)
        r.raise_for_status()
        return r.json().get("photos", [])
    photos = _search(_short_q(query)) or _search(_short_q(query, add_hint=False)) or _search("aesthetic minimal portrait")
    if not photos:
        return None, None
    fresh = [p for p in photos if p.get("id") not in used_ids]
    pool = fresh if fresh else photos
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(photo["id"])
    img_url = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    return requests.get(img_url, timeout=60).content, f'Foto: {photo.get("photographer","-")} (Pexels)'

def unsplash_pick(query, used_ids):
    r = requests.get("https://api.unsplash.com/search/photos",
                     headers={"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"},
                     params={"query": _short_q(query), "per_page": 15,
                             "orientation": PEXELS_ORIENTATION}, timeout=60)
    r.raise_for_status()
    results = r.json().get("results", [])
    if not results:
        return None, None
    fresh = [p for p in results if p.get("id") not in used_ids]
    pool = fresh if fresh else results
    photo = random.choice(pool[:8])
    if photo.get("id"):
        used_ids.add(photo["id"])
    urls = photo.get("urls", {})
    img_url = urls.get("regular") or urls.get("full") or urls.get("raw")
    name = (photo.get("user") or {}).get("name", "-")
    return requests.get(img_url, timeout=60).content, f"Foto: {name} (Unsplash)"

def image_pick(query, used_ids):
    sources = (["unsplash", "pexels"] if UNSPLASH_ACCESS_KEY else ["pexels"])
    for q in [query] + FALLBACK_QUERIES:
        for src in sources:
            try:
                data, credit = (unsplash_pick if src == "unsplash" else pexels_pick)(q, used_ids)
                if data:
                    return data, credit
            except Exception as e:
                print(f"    ! {src} gagal ('{_short_q(q)}'): {e}")
    return None, None

def make_placeholder():
    im = Image.new("RGB", (CANVAS_W, CANVAS_H), (30, 20, 45))
    d = ImageDraw.Draw(im)
    for y in range(CANVAS_H):
        t = y / CANVAS_H
        d.line([(0, y), (CANVAS_W, y)], fill=(int(40*(1-t)+15*t), int(25*(1-t)+12*t), int(70*(1-t)+25*t)))
    return im

def crop_canvas(im):
    target = CANVAS_W / CANVAS_H
    w, h = im.size
    if w / h > target:
        nw = int(h * target); x = (w - nw) // 2; im = im.crop((x, 0, x + nw, h))
    else:
        nh = int(w / target); y = (h - nh) // 2; im = im.crop((0, y, w, y + nh))
    return im.resize((CANVAS_W, CANVAS_H), Image.LANCZOS)

def film_grade(im):
    """Sentuhan film halus: saturasi sedikit turun, kontras tipis, hangat tipis."""
    if not FILM_GRADE:
        return im
    try:
        im = ImageEnhance.Color(im).enhance(0.90)
        im = ImageEnhance.Contrast(im).enhance(1.04)
        im = ImageEnhance.Brightness(im).enhance(1.01)
        r, g, b = im.split()
        r = r.point(lambda v: min(255, int(v * 1.02)))
        b = b.point(lambda v: int(v * 0.985))
        im = Image.merge("RGB", (r, g, b))
    except Exception:
        pass
    return im


# ---------- deteksi area kosong: atas/bawah/kiri/kanan/tengah ----------
ZONES = ("top", "bottom", "left", "right", "center")

def _busy(gray, box):
    try:
        return ImageStat.Stat(gray.crop(box)).stddev[0]
    except Exception:
        return 999.0

def analyze_zone(im):
    """Pilih area teks di bagian paling 'kosong' (detail paling rendah)."""
    g = im.convert("L")
    W, H = CANVAS_W, CANVAS_H
    xs = [0, W // 3, 2 * W // 3, W]
    ys = [0, H // 3, 2 * H // 3, H]
    cell = [[_busy(g, (xs[c], ys[r], xs[c + 1], ys[r + 1])) for c in range(3)] for r in range(3)]
    def avg(vals): return sum(vals) / len(vals)
    score = {
        "top":    avg(cell[0]),
        "bottom": avg(cell[2]),
        "left":   avg([cell[0][0], cell[1][0], cell[2][0]]),
        "right":  avg([cell[0][2], cell[1][2], cell[2][2]]),
        "center": cell[1][1],
    }
    # sedikit condong ke atas/bawah (lebih rapi); kiri/kanan/tengah dipakai kalau jelas lebih kosong
    for z in ("left", "right", "center"):
        score[z] += 3.0
    return min(ZONES, key=lambda z: score[z])


# ---------- layout teks per zona (dipakai pptx & svg) ----------
ZONE_LAYOUT = {
    "top":    dict(hx=0.07, hy=0.10, hw=0.86, bx=0.07, by=0.28, bw=0.86, align="left"),
    "bottom": dict(hx=0.07, hy=0.50, hw=0.86, bx=0.07, by=0.68, bw=0.86, align="left"),
    "left":   dict(hx=0.07, hy=0.33, hw=0.52, bx=0.07, by=0.53, bw=0.52, align="left"),
    "right":  dict(hx=0.44, hy=0.33, hw=0.50, bx=0.44, by=0.53, bw=0.50, align="left"),
    "center": dict(hx=0.10, hy=0.37, hw=0.80, bx=0.10, by=0.57, bw=0.80, align="center"),
}
def head_pt(zone):  return 34 if zone in ("left", "right") else (36 if zone == "center" else 38)
def body_pt(zone):  return 19 if zone in ("left", "right") else 21
def head_px(zone):  return 54 if zone in ("left", "right") else (58 if zone == "center" else 60)
def body_px(zone):  return 31 if zone in ("left", "right") else 34


def bake_scrim(im, zone, out_path):
    """Gelapkan area teks biar putih terbaca, mengikuti posisi zona."""
    W, H = CANVAS_W, CANVAS_H
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    def vgrad(y0, y1, a0, a1):
        span = max(1, y1 - y0)
        for y in range(max(0, y0), min(H, y1)):
            t = (y - y0) / span
            draw.line([(0, y), (W, y)], fill=(0, 0, 0, int(a0 + (a1 - a0) * t)))

    def hgrad(x0, x1, a0, a1):
        span = max(1, x1 - x0)
        for x in range(max(0, x0), min(W, x1)):
            t = (x - x0) / span
            draw.line([(x, 0), (x, H)], fill=(0, 0, 0, int(a0 + (a1 - a0) * t)))

    if zone == "top":
        vgrad(0, int(0.55 * H), 205, 0)
    elif zone == "bottom":
        vgrad(int(0.40 * H), H, 0, 225)
    elif zone == "left":
        hgrad(0, int(0.62 * W), 205, 0)
    elif zone == "right":
        hgrad(int(0.38 * W), W, 0, 205)
    else:  # center
        draw.rectangle([0, 0, W, H], fill=(0, 0, 0, 70))
        vgrad(int(0.28 * H), int(0.80 * H), 110, 110)

    # strip tipis atas (logo & tagline) + bawah (footer & tombol) SELALU ada
    vgrad(0, int(0.16 * H), 120, 0)
    vgrad(int(0.86 * H), H, 0, 150)

    Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB").save(out_path, "PNG")
    return out_path


# ---------- SLIDE CTA: mockup HP + kotak ungu ----------
def _rrect(draw, box, radius, **kw):
    try:
        draw.rounded_rectangle(box, radius=radius, **kw)
    except Exception:
        draw.rectangle(box, **kw)

def make_cta_background():
    """Latar ungu gelap + mockup HP (profil IG abstrak). Teks CTA ditaruh belakangan."""
    W, H = CANVAS_W, CANVAS_H
    im = Image.new("RGB", (W, H), (20, 12, 34))
    d = ImageDraw.Draw(im)
    ar, ag, ab = ACCENT_T
    for y in range(H):
        t = y / H
        d.line([(0, y), (W, y)],
               fill=(int(ar * 0.30 * (1 - t) + 14 * t),
                     int(ag * 0.22 * (1 - t) + 9 * t),
                     int(ab * 0.40 * (1 - t) + 26 * t)))
    # --- mockup HP di tengah-atas ---
    pw, ph = int(W * 0.46), int(H * 0.46)
    px = (W - pw) // 2
    py = int(H * 0.11)
    # bayangan
    _rrect(d, [px + 10, py + 14, px + pw + 10, py + ph + 14], 54, fill=(0, 0, 0, 255))
    # badan HP
    _rrect(d, [px, py, px + pw, py + ph], 54, fill=(248, 246, 252))
    # layar
    sx, sy = px + 16, py + 16
    sw, sh = pw - 32, ph - 32
    _rrect(d, [sx, sy, sx + sw, sy + sh], 40, fill=(255, 255, 255))
    # header profil: foto bulat + bar nama + tombol follow ungu
    cx = sx + 46
    cy = sy + 52
    rr = 34
    d.ellipse([cx - rr, cy - rr, cx + rr, cy + rr], fill=ACCENT_T)
    d.ellipse([cx - rr + 4, cy - rr + 4, cx + rr - 4, cy + rr - 4], fill=(236, 230, 246))
    _rrect(d, [cx + rr + 18, cy - 18, sx + sw - 90, cy - 2], 8, fill=(60, 50, 80))
    _rrect(d, [cx + rr + 18, cy + 4, sx + sw - 150, cy + 16], 6, fill=(170, 160, 190))
    _rrect(d, [sx + sw - 78, cy - 16, sx + sw - 18, cy + 14], 15, fill=ACCENT_T)  # follow btn
    # baris highlight (3 lingkaran)
    hy = sy + 120
    for i in range(3):
        hx = sx + 40 + i * 70
        d.ellipse([hx, hy, hx + 48, hy + 48], outline=(190, 180, 205), width=3)
    # grid thumbnail 3x2
    gy = hy + 78
    gap = 10
    tw = (sw - 40 - gap * 2) // 3
    for rrow in range(2):
        for ccol in range(3):
            gx = sx + 20 + ccol * (tw + gap)
            gyy = gy + rrow * (tw + gap)
            if gyy + tw > sy + sh - 20:
                break
            shade = 232 - ((rrow * 3 + ccol) % 3) * 12
            _rrect(d, [gx, gyy, gx + tw, gyy + tw], 10, fill=(shade, shade - 6, shade + 4))
    return im

def render_cta_full(bg_img, cta_text, out_path):
    """Preview CTA lengkap (latar + kotak ungu + teks dibake) buat Telegram."""
    W, H = CANVAS_W, CANVAS_H
    im = bg_img.convert("RGB").copy()
    d = ImageDraw.Draw(im)
    bx0, by0 = int(0.10 * W), int(0.63 * H)
    bx1, by1 = int(0.90 * W), int(0.90 * H)
    _rrect(d, [bx0, by0, bx1, by1], 40, fill=ACCENT_T)
    # teks CTA — cari TTF yang umum ada; tampilan FINAL tetap dari pptx/svg (editable)
    from PIL import ImageFont
    font = None
    for fp in ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
               "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
               "DejaVuSans-Bold.ttf", "DejaVuSans.ttf", "Arial.ttf"):
        try:
            font = ImageFont.truetype(fp, 42); break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()
    cx = (bx0 + bx1) // 2
    maxw = (bx1 - bx0) - 80
    words = " ".join(cta_text.split()).split(" ")
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        wd = d.textlength(trial, font=font) if hasattr(d, "textlength") else len(trial) * 20
        if wd <= maxw or not cur:
            cur = trial
        else:
            lines.append(cur); cur = w
    if cur:
        lines.append(cur)
    lines = lines[:6]
    lh = 56
    ty = (by0 + by1) // 2 - (len(lines) - 1) * lh // 2
    for ln in lines:
        d.text((cx, ty), ln, fill=(255, 255, 255), font=font, anchor="mm")
        ty += lh
    im.save(out_path, "PNG")
    return out_path


def download_logo_path():
    if not LOGO_URL:
        return None
    try:
        data = requests.get(LOGO_URL, headers=UA, timeout=30).content
        path = os.path.join(tempfile.gettempdir(), f"logo_{int(time.time()*1000)}.png")
        with open(path, "wb") as fp:
            fp.write(data)
        Image.open(path).verify()
        return path
    except Exception as e:
        print("  ! logo gagal diambil:", e)
        return None


# ======================================================================
# 5. PPTX
# ======================================================================
def _add_text(slide, left, top, width, height, text, size_pt, bold, font, color_hex,
              align=PP_ALIGN.LEFT, highlight=False):
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
            f.size = Pt(size_pt); f.bold = bold or is_hl; f.name = font
            f.color.rgb = ACCENT if is_hl else hex_rgb(color_hex)
    return box

def add_cta_pptx(slide):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, fx(0.66), fy(0.915), fx(0.28), fy(0.05))
    shp.fill.solid(); shp.fill.fore_color.rgb = WHITE
    shp.line.fill.background(); shp.shadow.inherit = False
    tf = shp.text_frame; tf.word_wrap = False; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    run = p.add_run(); run.text = CTA_TEXT
    f = run.font; f.size = Pt(15); f.bold = True; f.name = BODY_FONT; f.color.rgb = ACCENT

def _pp_align(a):
    return PP_ALIGN.CENTER if a == "center" else PP_ALIGN.LEFT

def _slide_header(slide, logo_path):
    if logo_path:
        try:
            slide.shapes.add_picture(logo_path, fx(0.06), fy(0.045), height=fy(0.045))
        except Exception:
            pass
    if BRAND_TAGLINE:
        _add_text(slide, 0.52, 0.045, 0.42, 0.08, BRAND_TAGLINE, 14, True, BODY_FONT, "FFFFFF", align=PP_ALIGN.RIGHT)

def build_pptx(slides_data, out_path, logo_path):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W); prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]
    for s in slides_data:
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=prs.slide_height)
        _slide_header(slide, logo_path)

        if s.get("kind") == "cta":
            # kotak ungu + teks CTA (editable)
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, fx(0.10), fy(0.63), fx(0.80), fy(0.27))
            box.fill.solid(); box.fill.fore_color.rgb = ACCENT
            box.line.fill.background(); box.shadow.inherit = False
            tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            run = p.add_run(); run.text = s["cta_text"]
            f = run.font; f.size = Pt(22); f.bold = True; f.name = HEADLINE_FONT; f.color.rgb = WHITE
            if FOOTER_TEXT:
                _add_text(slide, 0.06, 0.93, 0.88, 0.055, FOOTER_TEXT, 11, False, BODY_FONT, "FFFFFF", align=PP_ALIGN.CENTER)
            continue

        L = ZONE_LAYOUT.get(s.get("zone"), ZONE_LAYOUT["bottom"])
        al = _pp_align(L["align"])
        if s["headline"]:
            _add_text(slide, L["hx"], L["hy"], L["hw"], 0.18, s["headline"],
                      head_pt(s["zone"]), True, HEADLINE_FONT, "FFFFFF", align=al, highlight=True)
        if s["body"]:
            _add_text(slide, L["bx"], L["by"], L["bw"], 0.30, s["body"],
                      body_pt(s["zone"]), False, BODY_FONT, "FFFFFF", align=al, highlight=True)
        if FOOTER_TEXT:
            _add_text(slide, 0.06, 0.93, 0.58, 0.055, FOOTER_TEXT, 11, False, BODY_FONT, "FFFFFF")
        if s["total"] > 1 and CTA_TEXT:
            add_cta_pptx(slide)
    prs.save(out_path)
    return out_path


# ======================================================================
# 6. SVG
# ======================================================================
def _img_data_uri(path):
    im = Image.open(path).convert("RGB")
    buf = io.BytesIO(); im.save(buf, "JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()

def _svg_escape(t):
    return html.escape(t or "", quote=True)

def _est_w(s, size):
    # perkiraan lebar teks (tanpa metrik font) — cukup buat center/right align
    return len(s) * size * 0.52

def _svg_text(text, x, y, width, size, bold, font, anchor="start", highlight=False, gap=1.25):
    """
    Semua <text> dipakai anchor="start" lalu x-nya digeser manual supaya
    center/right tetap benar. Ini penting: beberapa renderer (cairosvg/Figma)
    menaruh tiap <tspan> di titik anchor kalau anchor=middle/end -> teks numpuk.
    'x' = titik acuan: kiri (start), tengah (middle), atau kanan (end).
    """
    weight = "700" if bold else "400"
    max_chars = max(6, int(width / (size * 0.55)))
    def words_of(line):
        res = []
        for seg, is_hl in (parse_highlights(line) if highlight else [(line, False)]):
            for w in seg.split(" "):
                if w != "":
                    res.append((w, is_hl))
        return res
    out = []
    yy = y
    for src in text.split("\n"):
        ws = words_of(src)
        if not ws:
            yy += int(size * gap); continue
        vlines, cur, cur_len = [], [], 0
        for w, is_hl in ws:
            add = len(w) + (1 if cur else 0)
            if cur and cur_len + add > max_chars:
                vlines.append(cur); cur, cur_len = [], 0; add = len(w)
            cur.append((w, is_hl)); cur_len += add
        if cur:
            vlines.append(cur)
        for vl in vlines:
            base = yy + size
            line_str = " ".join(w for w, _ in vl)
            lw = _est_w(line_str, size)
            if anchor == "middle":
                lx = int(x - lw / 2)
            elif anchor == "end":
                lx = int(x - lw)
            else:
                lx = int(x)
            spans = ""
            for k, (w, is_hl) in enumerate(vl):
                prefix = " " if k > 0 else ""
                fill = ACCENT_COLOR if is_hl else "FFFFFF"
                spans += f'<tspan fill="#{fill}">{_svg_escape(prefix + w)}</tspan>'
            out.append(f'<text x="{lx}" y="{int(base)}" text-anchor="start" xml:space="preserve" '
                       f'font-family="{_svg_escape(font)}, Arial, sans-serif" '
                       f'font-size="{size}" font-weight="{weight}">{spans}</text>')
            yy += int(size * gap)
    return "\n".join(out)

def build_svg(slides_data, out_path, logo_path):
    total = len(slides_data)
    W = CANVAS_W * total; H = CANVAS_H
    logo_uri = _img_data_uri(logo_path) if logo_path else None
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">']
    for idx, s in enumerate(slides_data):
        xo = idx * CANVAS_W
        parts.append(f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{H}" '
                     f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(s["image_path"])}"/>')
        if logo_uri:
            parts.append(f'<image x="{xo + int(0.06*CANVAS_W)}" y="{int(0.05*H)}" '
                         f'height="{int(0.05*H)}" href="{logo_uri}"/>')
        if BRAND_TAGLINE:
            parts.append(_svg_text(BRAND_TAGLINE, xo + int(0.94*CANVAS_W), int(0.05*H),
                                   int(0.4*CANVAS_W), 26, True, BODY_FONT, anchor="end"))

        if s.get("kind") == "cta":
            bx = xo + int(0.10 * CANVAS_W); by = int(0.63 * H)
            bw = int(0.80 * CANVAS_W); bh = int(0.27 * H)
            parts.append(f'<rect x="{bx}" y="{by}" width="{bw}" height="{bh}" rx="40" fill="#{ACCENT_COLOR}"/>')
            parts.append(_svg_text(s["cta_text"], xo + int(0.50 * CANVAS_W), by + int(bh * 0.18),
                                   int(0.70 * CANVAS_W), 40, True, HEADLINE_FONT,
                                   anchor="middle", gap=1.3))
            if FOOTER_TEXT:
                parts.append(_svg_text(FOOTER_TEXT, xo + int(0.50 * CANVAS_W), int(0.925 * H),
                                       int(0.88 * CANVAS_W), 20, False, BODY_FONT, anchor="middle"))
            continue

        L = ZONE_LAYOUT.get(s.get("zone"), ZONE_LAYOUT["bottom"])
        if L["align"] == "center":
            anc = "middle"
            hx = xo + int((L["hx"] + L["hw"] / 2) * CANVAS_W)
            bx = xo + int((L["bx"] + L["bw"] / 2) * CANVAS_W)
        else:
            anc = "start"
            hx = xo + int(L["hx"] * CANVAS_W)
            bx = xo + int(L["bx"] * CANVAS_W)
        if s["headline"]:
            parts.append(_svg_text(s["headline"], hx, int(L["hy"] * H),
                                   int(L["hw"] * CANVAS_W), head_px(s["zone"]), True, HEADLINE_FONT,
                                   anchor=anc, highlight=True))
        if s["body"]:
            parts.append(_svg_text(s["body"], bx, int(L["by"] * H),
                                   int(L["bw"] * CANVAS_W), body_px(s["zone"]), False, BODY_FONT,
                                   anchor=anc, highlight=True))
        if FOOTER_TEXT:
            parts.append(_svg_text(FOOTER_TEXT, xo + int(0.06*CANVAS_W), int(0.925*H),
                                   int(0.6*CANVAS_W), 20, False, BODY_FONT))
        if total > 1 and CTA_TEXT:
            px = xo + int(0.66*CANVAS_W); py = int(0.90*H)
            pw = int(0.28*CANVAS_W); ph = int(0.06*H)
            parts.append(f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" rx="{ph//2}" fill="#FFFFFF"/>')
            parts.append(f'<text x="{px+pw//2}" y="{py+int(ph*0.66)}" text-anchor="middle" '
                         f'font-family="{_svg_escape(BODY_FONT)}, Arial, sans-serif" '
                         f'font-size="26" font-weight="700" fill="#{ACCENT_COLOR}">{_svg_escape(CTA_TEXT)}</text>')
    parts.append('</svg>')
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return out_path


# ======================================================================
# 7. TELEGRAM
# ======================================================================
TG_BASE = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

def _tg_call(method, data, files=None, label=""):
    try:
        r = requests.post(f"{TG_BASE}/{method}", data=data, files=files, timeout=300)
        try:
            j = r.json()
        except Exception:
            j = {}
        time.sleep(0.5)
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
        return _tg_call("sendPhoto", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]}, {"photo": f}, label="photo")

def tg_document(path, caption=""):
    with open(path, "rb") as f:
        return _tg_call("sendDocument", {"chat_id": TELEGRAM_CHAT_ID, "caption": caption[:1000]}, {"document": f}, label="doc")


# ======================================================================
# 8. PROSES SATU KONTEN
# ======================================================================
def process_page(page, workdir):
    props = page["properties"]
    title = get_title(props) or "Tanpa Judul"
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

    slides_data, preview_paths, credits = [], [], []
    used_ids = set()
    total = len(slide_blocks)
    for i, block in enumerate(slide_blocks, start=1):
        # --- SLIDE CTA khusus ---
        if is_cta_block(block):
            cta_text = cta_text_of(block) or f"Klik link di bio {CTA_HANDLE}"
            bg = make_cta_background()
            img_path = os.path.join(workdir, f"slide_{i}.png")
            bg.save(img_path, "PNG")                       # latar editable (tanpa kotak)
            prev_path = os.path.join(workdir, f"preview_{i}.png")
            render_cta_full(bg, cta_text, prev_path)       # preview penuh buat Telegram
            preview_paths.append(prev_path)
            slides_data.append({"kind": "cta", "cta_text": cta_text, "headline": "", "body": "",
                                "image_path": img_path, "index": i, "total": total, "zone": "bottom"})
            credits.append(f"Slide {i}: (slide CTA — mockup HP + kotak ungu)")
            continue

        headline, body = headline_and_body(block)
        q = keyword_for(i - 1, keyword_blocks, headline, body)
        data, credit = image_pick(q, used_ids)
        if data:
            im = Image.open(io.BytesIO(data)).convert("RGB")
            if credit:
                credits.append(f"Slide {i}: {credit}")
        else:
            im = make_placeholder()
            credits.append(f"Slide {i}: (background netral — gambar '{q}' tak ditemukan)")
        base = film_grade(crop_canvas(im))
        zone = "bottom" if i == 1 else analyze_zone(base)   # slide 1 = cover (bawah)
        img_path = os.path.join(workdir, f"slide_{i}.png")
        bake_scrim(base, zone, img_path)
        preview_paths.append(img_path)
        slides_data.append({"kind": "normal", "headline": headline, "body": body,
                            "image_path": img_path, "index": i, "total": total, "zone": zone})

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    logo_path = download_logo_path()

    lines = [f"🎨 {title}  ({fmt}, {total} slide)", ""]
    for i, s in enumerate(slides_data, start=1):
        if s.get("kind") == "cta":
            lines.append(f"— Slide {i}: [CTA] {s['cta_text'][:120]}")
            continue
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:150] + "…") if len(s["body"]) > 150 else s["body"]
        zmap = {"top": "atas", "bottom": "bawah", "left": "kiri", "right": "kanan", "center": "tengah"}
        lines.append(f"— Slide {i} (teks {zmap.get(s['zone'], s['zone'])}): {h}")
        if b:
            lines.append(f"  {b}")
    lines += ["", *credits]
    tg_message("\n".join(lines))

    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")

    tg_message(f"📎 ====================\nFILE DESAIN: {title}\n====================")

    if OUTPUT_PPTX:
        pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
        build_pptx(slides_data, pptx_path, logo_path)
        if not tg_document(pptx_path, caption=f"{title} — .pptx (semua slide): import ke Canva ✨"):
            tg_message("⚠️ File .pptx gagal dikirim (cek log).")

    if OUTPUT_SVG:
        svg_path = os.path.join(workdir, f"{safe_name}.svg")
        build_svg(slides_data, svg_path, logo_path)
        if not tg_document(svg_path, caption=f"{title} — .svg (semua slide): tarik ke Figma ✨"):
            tg_message("⚠️ File .svg gagal dikirim (cek log).")

    if OUTPUT_SVG_PER_SLIDE and total > 1:
        for s in slides_data:
            sp = os.path.join(workdir, f"{safe_name}_slide{s['index']}.svg")
            build_svg([s], sp, logo_path)
            if not tg_document(sp, caption=f"{title} — slide {s['index']}/{total} (.svg per slide)"):
                tg_message(f"⚠️ SVG slide {s['index']} gagal dikirim (cek log).")

    return total


# ======================================================================
# 9. MAIN
# ======================================================================
def main():
    print(f"== ROBOT DESAIN NIKAH INSTITUTE ({CANVAS_W}x{CANVAS_H}) | unsplash={'on' if UNSPLASH_ACCESS_KEY else 'off'} ==")
    if tg_message("✅ Robot Nikah Institute — mulai jalan."):
        print("  Telegram OK.")
    else:
        print("  !! Telegram BERMASALAH — cek TELEGRAM_BOT_TOKEN & TELEGRAM_CHAT_ID.")
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
            notion_set_status(page_id, STATUS_DONE, icon_emoji=DONE_EMOJI)
            print("  v Sukses & status ->", STATUS_DONE, "| ikon ->", DONE_EMOJI)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            print("  x GAGAL:", err)
            traceback.print_exc()
            try:
                tg_message(f"⚠️ Gagal memproses '{get_title(page['properties'])}'.\n{err}")
            except Exception:
                pass
            notion_set_status(page_id, STATUS_ERROR, note=err)
    print("== Selesai ==")


if __name__ == "__main__":
    main()
