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
import json
import html
import time
import base64
import random
import secrets
import tempfile
import traceback

import requests
from PIL import Image, ImageDraw, ImageStat, ImageEnhance, ImageChops, ImageFilter
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
HEADLINE_FONT = os.environ.get("HEADLINE_FONT") or "Montserrat"
BODY_FONT     = os.environ.get("BODY_FONT")     or "Montserrat"
CTA_HANDLE    = os.environ.get("CTA_HANDLE")    or "@nikahinstitute"
# gaya gambar: dibiaskan ke aesthetic/editorial luar negeri, premium, bukan norak
STYLE_HINT    = os.environ.get("STYLE_HINT")    or "editorial lifestyle film"
FILM_GRADE    = (os.environ.get("FILM_GRADE") or "true").lower() == "true"

CANVAS_W = int(os.environ.get("CANVAS_W") or 1080)
CANVAS_H = int(os.environ.get("CANVAS_H") or 1350)

OUTPUT_PPTX = (os.environ.get("OUTPUT_PPTX") or "true").lower() == "true"
OUTPUT_SVG  = (os.environ.get("OUTPUT_SVG")  or "true").lower() == "true"
OUTPUT_SVG_PER_SLIDE = (os.environ.get("OUTPUT_SVG_PER_SLIDE") or "true").lower() == "true"

# ----------------------------------------------------------------------
# SISTEM BELAJAR (memory + feedback). File memori disimpan di repo.
# ----------------------------------------------------------------------
MEMORY_PATH = os.environ.get("MEMORY_FILE") or "memory.json"
LEARN_ENABLED = (os.environ.get("LEARN_ENABLED") or "true").lower() == "true"

# "knob" yang bisa digeser oleh feedback (nilai awal = netral)
G_HEAD_SCALE = 1.0   # pengali ukuran judul
G_BODY_SCALE = 1.0   # pengali ukuran body
G_SCRIM      = 1.0   # pengali kegelapan overlay/masking
G_SOFT       = 0.0   # tambahan "film/soft" (0 = normal, makin besar makin lembut)

def _clamp(v, lo, hi):
    return max(lo, min(hi, v))


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

def rich_segments(text):
    """Pecah teks jadi segmen (teks, stabilo?, miring?). ==stabilo== dan _miring_."""
    out = []
    for part in re.split(r"(==.+?==|_[^_\n]+_)", text):
        if not part:
            continue
        if len(part) >= 4 and part.startswith("==") and part.endswith("=="):
            out.append((part[2:-2], True, False))
        elif len(part) >= 3 and part.startswith("_") and part.endswith("_"):
            out.append((part[1:-1], False, True))
        else:
            out.append((part, False, False))
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
        im = ImageEnhance.Color(im).enhance(_clamp(0.90 - G_SOFT, 0.45, 1.0))  # knob: makin soft makin turun saturasi
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
# Slide 1 (cover) judulnya besar; slide lain judul kecil (sesuai permintaan)
# Ukuran dikali "knob" dari sistem belajar (G_HEAD_SCALE / G_BODY_SCALE).
def head_pt(zone, cover):  return max(10, int(round((40 if cover else 23) * G_HEAD_SCALE)))
def body_pt(zone):         return max(9,  int(round((19 if zone in ("left", "right") else 21) * G_BODY_SCALE)))
def head_px(zone, cover):  return max(16, int(round((66 if cover else 37) * G_HEAD_SCALE)))
def body_px(zone):         return max(14, int(round((31 if zone in ("left", "right") else 34) * G_BODY_SCALE)))

def body_top(zone, cover):
    L = ZONE_LAYOUT[zone]
    return L["by"] if cover else (L["hy"] + 0.085)


def _interp(stops, t):
    """stops = [(frac, value), ...] menaik; interpolasi halus (smoothstep)."""
    if t <= stops[0][0]:
        return stops[0][1]
    if t >= stops[-1][0]:
        return stops[-1][1]
    for i in range(1, len(stops)):
        if t <= stops[i][0]:
            t0, v0 = stops[i - 1]; t1, v1 = stops[i]
            r = (t - t0) / (t1 - t0) if t1 > t0 else 0.0
            r = r * r * (3 - 2 * r)  # smoothstep -> tanpa garis keras
            return v0 + (v1 - v0) * r
    return stops[-1][1]

def _mask_v(stops):
    col = Image.new("L", (1, CANVAS_H)); px = col.load()
    for y in range(CANVAS_H):
        px[0, y] = int(max(0, min(255, _interp(stops, y / (CANVAS_H - 1)))))
    return col.resize((CANVAS_W, CANVAS_H))

def _mask_h(stops):
    row = Image.new("L", (CANVAS_W, 1)); px = row.load()
    for x in range(CANVAS_W):
        px[x, 0] = int(max(0, min(255, _interp(stops, x / (CANVAS_W - 1)))))
    return row.resize((CANVAS_W, CANVAS_H))

def bake_scrim(im, zone, out_path):
    """
    Gelapkan area teks dengan masking HALUS (smoothstep + blur), jadi nggak
    ada garis potong. Untuk zona samping (kiri/kanan), gelapnya memudar juga
    di atas & bawah supaya nggak kelihatan 'kepotong'.
    """
    W, H = CANVAS_W, CANVAS_H

    # feather vertikal buat zona samping: habis mulus di tepi atas & bawah
    vfeather = _mask_v([(0.0, 0), (0.17, 255), (0.83, 255), (1.0, 0)])

    if zone == "top":
        mask = _mask_v([(0.0, 210), (0.30, 170), (0.55, 0), (1.0, 0)])
    elif zone == "bottom":
        mask = _mask_v([(0.0, 0), (0.42, 0), (0.70, 170), (1.0, 225)])
    elif zone == "left":
        mask = ImageChops.multiply(_mask_h([(0.0, 210), (0.30, 150), (0.60, 0), (1.0, 0)]), vfeather)
    elif zone == "right":
        mask = ImageChops.multiply(_mask_h([(0.0, 0), (0.40, 0), (0.70, 150), (1.0, 210)]), vfeather)
    else:  # center
        mask = _mask_v([(0.0, 40), (0.30, 150), (0.70, 150), (1.0, 40)])

    # strip tipis atas (logo & tagline) + bawah (footer & tombol) SELALU ada,
    # tapi juga mulus (smoothstep) -> digabung pakai 'lighter' (ambil yang tergelap)
    top_strip = _mask_v([(0.0, 120), (0.16, 0), (1.0, 0)])
    bot_strip = _mask_v([(0.0, 0), (0.86, 0), (1.0, 150)])
    mask = ImageChops.lighter(ImageChops.lighter(mask, top_strip), bot_strip)

    # knob kegelapan overlay (dari sistem belajar)
    if G_SCRIM != 1.0:
        mask = mask.point(lambda v: int(_clamp(v * G_SCRIM, 0, 255)))

    # blur tipis -> hilangkan banding/garis
    mask = mask.filter(ImageFilter.GaussianBlur(radius=max(10, W // 54)))

    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 255))
    overlay.putalpha(mask)
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
        segs = rich_segments(line) if highlight else [(line, False, False)]
        for seg, is_hl, is_it in (segs or [("", False, False)]):
            run = p.add_run()
            run.text = seg
            f = run.font
            f.size = Pt(size_pt); f.bold = bold or is_hl; f.name = font
            f.italic = is_it
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

        zone = s.get("zone", "bottom")
        cover = (s.get("index") == 1)
        L = ZONE_LAYOUT.get(zone, ZONE_LAYOUT["bottom"])
        al = _pp_align(L["align"])
        if s["headline"]:
            _add_text(slide, L["hx"], L["hy"], L["hw"], 0.18, s["headline"],
                      head_pt(zone, cover), True, HEADLINE_FONT, "FFFFFF", align=al, highlight=True)
        if s["body"]:
            _add_text(slide, L["bx"], body_top(zone, cover), L["bw"], 0.34, s["body"],
                      body_pt(zone), False, BODY_FONT, "FFFFFF", align=al, highlight=True)
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
    SATU <text> per PARAGRAF (biar di Figma jadi 1 layer teks yang gampang
    diedit). Pemisahan baris (wrap) pakai <tspan dy=...> di dalam <text> yang
    sama, bukan <text> terpisah. Warna stabilo tetap jalan lewat tspan fill.
    'x' = titik acuan: kiri (start), tengah (middle), kanan (end).
    """
    weight = "700" if bold else "400"
    # SATU <text> untuk SELURUH blok teks ini (headline = 1 blok, body = 1 blok).
    # Baris (Enter di Notion) dipisah pakai NEWLINE ASLI di dalam satu elemen,
    # TANPA <tspan> ber-posisi. Alasan: Figma memecah tiap <tspan> ber-x/-dy jadi
    # LAYER terpisah; dgn newline asli, Figma menyatukannya jadi SATU text box
    # berisi baris-barisnya. Warna stabilo tetap inline (tidak bikin layer baru).
    raw_lines = text.split("\n")
    nonempty = [l for l in raw_lines if l.strip()]
    if not nonempty:
        return ""
    def est_line(line):
        plain = "".join(s for s, _, _ in (rich_segments(line) if highlight else [(line, False, False)]))
        return _est_w(plain, size)
    maxw = max(est_line(l) for l in nonempty)
    if anchor == "middle":
        lx = int(x - maxw / 2)
    elif anchor == "end":
        lx = int(x - maxw)
    else:
        lx = int(x)
    base = y + size
    parts = []
    for li, line in enumerate(raw_lines):
        if li > 0:
            parts.append("\n")   # newline asli -> jadi line break di dalam 1 text box Figma
        for s, is_hl, is_it in (rich_segments(line) if highlight else [(line, False, False)]):
            if s == "":
                continue
            tok = _svg_escape(s)
            if is_hl:
                parts.append(f'<tspan fill="#{ACCENT_COLOR}">{tok}</tspan>')
            elif is_it:
                parts.append(f'<tspan font-style="italic">{tok}</tspan>')
            else:
                parts.append(tok)
    content = "".join(parts)
    return (f'<text x="{lx}" y="{int(base)}" text-anchor="start" xml:space="preserve" '
            f'style="white-space:pre" fill="#FFFFFF" '
            f'font-family="{_svg_escape(font)}, Arial, sans-serif" '
            f'font-size="{size}" font-weight="{weight}">{content}</text>')


# ---------- masking versi VEKTOR (editable di Figma) ----------
def _grad_lin(gid, vertical, stops):
    coords = 'x1="0" y1="0" x2="0" y2="1"' if vertical else 'x1="0" y1="0" x2="1" y2="0"'
    s = "".join(f'<stop offset="{o}" stop-color="#000000" stop-opacity="{op}"/>' for o, op in stops)
    return f'<linearGradient id="{gid}" {coords}>{s}</linearGradient>'

def _grad_rad(gid, cx, stops):
    s = "".join(f'<stop offset="{o}" stop-color="#000000" stop-opacity="{op}"/>' for o, op in stops)
    return f'<radialGradient id="{gid}" cx="{cx}" cy="0.5" r="0.8">{s}</radialGradient>'

def _scrim_svg(zone, xo, idx):
    """Kembalikan (defs, rects) berisi gradasi gelap sebagai elemen vektor."""
    W, H = CANVAS_W, CANVAS_H
    gid = f"sc{idx}"
    defs, rects = [], []
    # gradasi utama sesuai zona
    if zone == "top":
        defs.append(_grad_lin(gid, True, [(0, 0.80), (0.55, 0), (1, 0)]))
    elif zone == "left":
        defs.append(_grad_lin(gid, False, [(0, 0.80), (0.60, 0), (1, 0)]))
    elif zone == "right":
        defs.append(_grad_lin(gid, False, [(0, 0), (0.40, 0), (1, 0.80)]))
    elif zone == "center":
        defs.append(_grad_rad(gid, 0.5, [(0, 0.60), (1, 0.06)]))
    else:  # bottom
        defs.append(_grad_lin(gid, True, [(0, 0), (0.42, 0), (1, 0.88)]))
    rects.append(f'<rect x="{xo}" y="0" width="{W}" height="{H}" fill="url(#{gid})"/>')
    # strip tipis atas (logo/tagline) & bawah (footer/tombol) — tetap editable
    defs.append(_grad_lin(gid + "t", True, [(0, 0.45), (1, 0)]))
    defs.append(_grad_lin(gid + "b", True, [(0, 0), (1, 0.55)]))
    rects.append(f'<rect x="{xo}" y="0" width="{W}" height="{int(0.16*H)}" fill="url(#{gid}t)"/>')
    rects.append(f'<rect x="{xo}" y="{int(0.86*H)}" width="{W}" height="{int(0.14*H)}" fill="url(#{gid}b)"/>')
    return "\n".join(defs), "\n".join(rects)

def build_svg(slides_data, out_path, logo_path):
    total = len(slides_data)
    W = CANVAS_W * total; H = CANVAS_H
    logo_uri = _img_data_uri(logo_path) if logo_path else None
    defs_all, parts = [], []
    for idx, s in enumerate(slides_data):
        xo = idx * CANVAS_W
        # foto BERSIH (tanpa masking dibakar) -> masking ditaruh sbg vektor
        photo = s.get("clean_path") or s["image_path"]
        parts.append(f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{H}" '
                     f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(photo)}"/>')
        # masking vektor (cuma utk slide foto biasa; CTA latar sudah gelap)
        if s.get("kind") != "cta":
            d, r = _scrim_svg(s.get("zone", "bottom"), xo, idx)
            defs_all.append(d); parts.append(r)
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

        zone = s.get("zone", "bottom")
        cover = (s.get("index") == 1)
        L = ZONE_LAYOUT.get(zone, ZONE_LAYOUT["bottom"])
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
                                   int(L["hw"] * CANVAS_W), head_px(zone, cover), True, HEADLINE_FONT,
                                   anchor=anc, highlight=True))
        if s["body"]:
            parts.append(_svg_text(s["body"], bx, int(body_top(zone, cover) * H),
                                   int(L["bw"] * CANVAS_W), body_px(zone), False, BODY_FONT,
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
    header = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
              f'viewBox="0 0 {W} {H}">')
    defs = "<defs>\n" + "\n".join(d for d in defs_all if d) + "\n</defs>" if defs_all else ""
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(header + "\n" + defs + "\n" + "\n".join(parts) + "\n</svg>")
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

def tg_buttons(text, keyboard):
    """Kirim pesan + tombol inline. keyboard = list of rows, tiap tombol (label, callback_data)."""
    kb = {"inline_keyboard": [[{"text": t, "callback_data": d} for (t, d) in row] for row in keyboard]}
    return _tg_call("sendMessage", {"chat_id": TELEGRAM_CHAT_ID, "text": text[:3500],
                                    "reply_markup": json.dumps(kb)}, label="btn")

def tg_get_updates(offset):
    try:
        r = requests.get(f"{TG_BASE}/getUpdates",
                         params={"offset": offset, "timeout": 0, "allowed_updates": json.dumps(["callback_query", "message"])},
                         timeout=40)
        j = r.json()
        return j.get("result", []) if j.get("ok") else []
    except Exception as e:
        print("    ! getUpdates error:", e)
        return []

def tg_answer_callback(cb_id, text=""):
    _tg_call("answerCallbackQuery", {"callback_query_id": cb_id, "text": text[:180]}, label="ans")


# ======================================================================
# 7b. SISTEM BELAJAR — memori + feedback (approve/reject + alasan)
# ======================================================================
DEFAULT_MEMORY = {
    "preferences": {"head_scale": 1.0, "body_scale": 1.0, "scrim": 1.0, "soft": 0.0},
    "pending": {},        # token -> {title, ts}
    "awaiting_text": None,  # token yang menunggu alasan ketik
    "log": [],            # riwayat {ts, title, status, reason}
    "stats": {"approved": 0, "rejected": 0, "weekly": []},
    "tg_offset": 0,
}

# tombol alasan saat Reject: (label, kode)
REASON_BUTTONS = [
    ("Teks kegedean", "tbig"),
    ("Teks kekecilan", "tsmall"),
    ("Teks susah kebaca", "hard"),
    ("Overlay kegelapan", "dark"),
    ("Gambar kurang aesthetic", "img"),
    ("Alasan lain (ketik)", "other"),
]
REASON_LABEL = {k: v for v, k in REASON_BUTTONS}

def load_memory():
    try:
        with open(MEMORY_PATH, "r", encoding="utf-8") as f:
            mem = json.load(f)
        for k, v in DEFAULT_MEMORY.items():
            mem.setdefault(k, v if not isinstance(v, (dict, list)) else (dict(v) if isinstance(v, dict) else list(v)))
        for k, v in DEFAULT_MEMORY["preferences"].items():
            mem["preferences"].setdefault(k, v)
        return mem
    except Exception:
        import copy
        return copy.deepcopy(DEFAULT_MEMORY)

def save_memory(mem):
    try:
        with open(MEMORY_PATH, "w", encoding="utf-8") as f:
            json.dump(mem, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("    ! gagal simpan memori:", e)

def apply_reason(prefs, code):
    """Geser knob sesuai alasan. Mengembalikan deskripsi singkat perubahan."""
    if code == "tbig":
        prefs["head_scale"] = _clamp(prefs["head_scale"] - 0.06, 0.7, 1.3)
        prefs["body_scale"] = _clamp(prefs["body_scale"] - 0.05, 0.7, 1.3)
        return "teks dikecilin"
    if code == "tsmall":
        prefs["head_scale"] = _clamp(prefs["head_scale"] + 0.06, 0.7, 1.3)
        prefs["body_scale"] = _clamp(prefs["body_scale"] + 0.05, 0.7, 1.3)
        return "teks digedein"
    if code == "hard":
        prefs["scrim"] = _clamp(prefs["scrim"] + 0.12, 0.6, 1.6)
        return "overlay dipergelap (biar teks kebaca)"
    if code == "dark":
        prefs["scrim"] = _clamp(prefs["scrim"] - 0.12, 0.6, 1.6)
        return "overlay dipertipis"
    if code == "img":
        prefs["soft"] = _clamp(prefs["soft"] + 0.08, 0.0, 0.4)
        return "gambar dibikin lebih soft/film"
    return "dicatat (tanpa ubah setelan)"

def prefs_summary(prefs):
    return (f"Setelan belajar sekarang: judul {int(prefs['head_scale']*100)}%, "
            f"body {int(prefs['body_scale']*100)}%, overlay {int(prefs['scrim']*100)}%, "
            f"soft +{int(prefs['soft']*100)}%.")

def process_feedback(mem):
    """Baca pencetan tombol sejak run terakhir, update stats + knob."""
    prefs = mem["preferences"]
    updates = tg_get_updates(mem.get("tg_offset", 0) + 1)
    changed = []
    for up in updates:
        mem["tg_offset"] = max(mem.get("tg_offset", 0), up.get("update_id", 0))

        # pesan teks: alasan ketik / perintah stats / reset
        msg = up.get("message")
        if msg:
            text = (msg.get("text") or "").strip()
            low = text.lower()
            if mem.get("awaiting_text"):
                tok = mem["awaiting_text"]
                if text:
                    rec = mem["pending"].pop(tok, {"title": "?"})
                    mem["log"].append({"ts": int(time.time()), "title": rec.get("title", "?"),
                                       "status": "rejected", "reason": "ketik: " + text[:120]})
                    mem["awaiting_text"] = None
                    tg_message(f"📝 Alasan dicatat: “{text[:120]}”. Makasih, bro.")
                continue
            if low in ("stats", "/stats", "statistik"):
                a = mem["stats"]["approved"]; r = mem["stats"]["rejected"]; tot = a + r
                rate = int(100 * a / tot) if tot else 0
                tg_message(f"📊 Statistik belajar Nikah:\n"
                           f"Approve {a} / Reject {r}  (approval rate {rate}%).\n"
                           f"{prefs_summary(mem['preferences'])}")
                continue
            if low in ("/reset", "reset belajar"):
                mem["preferences"] = dict(DEFAULT_MEMORY["preferences"])
                tg_message("🔄 Setelan belajar direset ke awal (netral). Statistik tetap tersimpan.")
                continue
            continue

        cb = up.get("callback_query")
        if not cb:
            continue
        data = cb.get("data", "")
        cb_id = cb.get("id", "")
        parts = data.split(":")
        kind = parts[0]
        tok = parts[1] if len(parts) > 1 else ""

        if kind == "a":   # approve
            rec = mem["pending"].pop(tok, None)
            mem["stats"]["approved"] += 1
            mem["log"].append({"ts": int(time.time()), "title": (rec or {}).get("title", "?"),
                               "status": "approved", "reason": ""})
            tg_answer_callback(cb_id, "✅ Disimpan sebagai contoh bagus!")
        elif kind == "r":  # reject -> minta alasan
            mem["stats"]["rejected"] += 1
            tg_answer_callback(cb_id, "❌ Oke, pilih alasannya ya.")
            rows = []
            row = []
            for label, code in REASON_BUTTONS:
                row.append((label, f"rs:{tok}:{code}"))
                if len(row) == 2:
                    rows.append(row); row = []
            if row:
                rows.append(row)
            title = mem["pending"].get(tok, {}).get("title", "desain ini")
            tg_buttons(f"Kenapa '{title}' ditolak?", rows)
        elif kind == "rs":  # reason chosen
            code = parts[2] if len(parts) > 2 else ""
            if code == "other":
                mem["awaiting_text"] = tok
                tg_answer_callback(cb_id, "Ketik alasannya di chat ya.")
                tg_message("✍️ Tulis alasan singkatnya di sini (1 pesan).")
            else:
                desc = apply_reason(prefs, code)
                rec = mem["pending"].pop(tok, {"title": "?"})
                mem["log"].append({"ts": int(time.time()), "title": rec.get("title", "?"),
                                   "status": "rejected", "reason": REASON_LABEL.get(code, code)})
                changed.append(desc)
                tg_answer_callback(cb_id, f"Paham. {desc}.")
    if changed:
        tg_message("🧠 Bot menyesuaikan diri: " + "; ".join(changed) + ".\n" + prefs_summary(prefs))
    return mem

def apply_prefs_to_globals(mem):
    global G_HEAD_SCALE, G_BODY_SCALE, G_SCRIM, G_SOFT
    p = mem.get("preferences", {})
    G_HEAD_SCALE = float(p.get("head_scale", 1.0))
    G_BODY_SCALE = float(p.get("body_scale", 1.0))
    G_SCRIM      = float(p.get("scrim", 1.0))
    G_SOFT       = float(p.get("soft", 0.0))


# ======================================================================
# 8. PROSES SATU KONTEN
# ======================================================================
def process_page(page, workdir, mem=None):
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
                                "image_path": img_path, "clean_path": img_path,
                                "index": i, "total": total, "zone": "bottom"})
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
        clean_path = os.path.join(workdir, f"clean_{i}.png")
        base.save(clean_path, "PNG")                        # foto bersih utk .svg (Figma)
        img_path = os.path.join(workdir, f"slide_{i}.png")
        bake_scrim(base, zone, img_path)                    # versi gelap utk preview & .pptx
        preview_paths.append(img_path)
        slides_data.append({"kind": "normal", "headline": headline, "body": body,
                            "image_path": img_path, "clean_path": clean_path,
                            "index": i, "total": total, "zone": zone})

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

    # ---- tombol penilaian (approve / reject) utk sistem belajar ----
    if LEARN_ENABLED and mem is not None:
        token = secrets.token_hex(4)
        mem["pending"][token] = {"title": title, "ts": int(time.time())}
        tg_buttons(f"Gimana desain '{title}'? Nilai ya biar bot makin ngerti seleramu 👇",
                   [[("✅ Approve", f"a:{token}"), ("❌ Reject", f"r:{token}")]])

    return total


# ======================================================================
# 9. MAIN
# ======================================================================
def main():
    print(f"== ROBOT DESAIN NIKAH INSTITUTE ({CANVAS_W}x{CANVAS_H}) | unsplash={'on' if UNSPLASH_ACCESS_KEY else 'off'} | belajar={'on' if LEARN_ENABLED else 'off'} ==")

    # 1) muat memori + 2) proses feedback (pencetan tombol) sejak run lalu
    mem = load_memory()
    if LEARN_ENABLED:
        try:
            mem = process_feedback(mem)
        except Exception as e:
            print("  ! proses feedback gagal:", e)
        apply_prefs_to_globals(mem)
        print("  " + prefs_summary(mem["preferences"]))

    if tg_message("✅ Robot Nikah Institute — mulai jalan."):
        print("  Telegram OK.")
    else:
        print("  !! Telegram BERMASALAH — cek TELEGRAM_BOT_TOKEN & TELEGRAM_CHAT_ID.")

    pages = notion_find_ready()
    print(f"Ditemukan {len(pages)} konten berstatus '{STATUS_READY}'.")
    if not pages:
        print("Tidak ada konten baru. (Feedback tetap diproses.)")
        save_memory(mem)
        return
    for page in pages:
        page_id = page["id"]
        workdir = tempfile.mkdtemp()
        try:
            process_page(page, workdir, mem)
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

    save_memory(mem)
    print("== Selesai ==")


if __name__ == "__main__":
    main()
