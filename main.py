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
ACCENT_COLOR  = (os.environ.get("ACCENT_COLOR") or "7C3AED").lstrip("#")   # ungu (warna utama)
ACCENT2_COLOR = (os.environ.get("ACCENT2_COLOR") or "159A9A").lstrip("#")  # tosca/teal (warna kedua)
LOGO_URL      = os.environ.get("LOGO_URL")      or ""
LOGO_DARK_URL = os.environ.get("LOGO_DARK_URL") or ""   # logo versi gelap (dipakai di slide background TERANG)
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
FORMAT_PROPERTY  = env("FORMAT_PROPERTY", "Bentuk Konten")
CONTENT_PROPERTY = env("CONTENT_PROPERTY", "Isi Konten")
KEYWORD_PROPERTY = env("KEYWORD_PROPERTY", "Kata Kunci Gambar")
GAYA_PROPERTY    = env("GAYA_PROPERTY", "Gaya Desain")   # kolom dropdown pilih gaya cover

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
ACCENT2 = hex_rgb(ACCENT2_COLOR)
ACCENT2_T = hex_tuple(ACCENT2_COLOR)
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
    order = ["status", "select"] if STATUS_TYPE == "status" else ["select", "status"]
    last = None
    for stype in order:   # coba kedua tipe (Status vs Select) biar nggak perlu setel manual
        flt = {"property": STATUS_PROPERTY, stype: {"equals": STATUS_READY}}
        resp = requests.post(url, headers=NOTION_HEADERS, json={"filter": flt}, timeout=60)
        if resp.status_code == 200:
            return resp.json().get("results", [])
        last = resp
    raise SystemExit(f"[NOTION ERROR] {getattr(last,'status_code','?')}: {getattr(last,'text','')}")

def notion_get_page(page_id):
    """Ambil ulang 1 halaman Notion by ID (buat auto-revisi)."""
    try:
        r = requests.get(f"{NOTION_BASE}/pages/{page_id}", headers=NOTION_HEADERS, timeout=60)
        if r.status_code == 200:
            return r.json()
    except Exception as e:
        print("  ! notion_get_page gagal:", e)
    return None

def notion_page_images(page_id):
    """Ambil URL semua gambar yang DITEMPEL di body halaman Notion, urut dari atas.
    Dipakai fitur 'bawa gambar sendiri' saat slide ditandai [img] tanpa URL."""
    urls, cursor = [], None
    try:
        for _ in range(10):   # maks 10 halaman blok (aman)
            params = {"page_size": 100}
            if cursor:
                params["start_cursor"] = cursor
            r = requests.get(f"{NOTION_BASE}/blocks/{page_id}/children",
                             headers=NOTION_HEADERS, params=params, timeout=30)
            if r.status_code != 200:
                break
            j = r.json()
            for b in j.get("results", []):
                if b.get("type") == "image":
                    img = b.get("image", {})
                    u = (img.get("file") or {}).get("url") or (img.get("external") or {}).get("url")
                    if u:
                        urls.append(u)
            if j.get("has_more"):
                cursor = j.get("next_cursor")
            else:
                break
    except Exception as e:
        print("    ! gagal baca attachment Notion:", e)
    return urls

def notion_set_status(page_id, status_value, note=None, icon_emoji=None):
    url = f"{NOTION_BASE}/pages/{page_id}"
    order = ["status", "select"] if STATUS_TYPE == "status" else ["select", "status"]
    for stype in order:   # coba kedua tipe; berhenti kalau salah satu berhasil
        props = {STATUS_PROPERTY: {stype: {"name": status_value}}}
        body = {"properties": dict(props)}
        if note:
            body["properties"]["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
        if icon_emoji:
            body["icon"] = {"type": "emoji", "emoji": icon_emoji}
        r = requests.patch(url, headers=NOTION_HEADERS, json=body, timeout=60)
        if r.status_code == 200:
            return
        # kalau gagal (mis. kolom 'Catatan' nggak ada), ulangi tanpa catatan
        body2 = {"properties": props}
        if icon_emoji:
            body2["icon"] = {"type": "emoji", "emoji": icon_emoji}
        r2 = requests.patch(url, headers=NOTION_HEADERS, json=body2, timeout=60)
        if r2.status_code == 200:
            return

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
    # Pemisah slide didukung: baris 'Slide 1/2/3...' ATAU baris '---' / dash.
    blocks = re.split(r"(?mi)^\s*(?:slide\s*\d+\s*[:.)\-]?|[-–—]{2,}|[–—])\s*$", raw)
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
    q = re.sub(r"==|~~", "", q)
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
    """Pecah teks jadi segmen (teks, ungu?, miring?, tosca?).
    ==ungu==  ~~tosca~~  _miring_"""
    out = []
    for part in re.split(r"(==.+?==|~~.+?~~|_[^_\n]+_)", text):
        if not part:
            continue
        if len(part) >= 4 and part.startswith("==") and part.endswith("=="):
            out.append((part[2:-2], True, False, False))     # ungu
        elif len(part) >= 4 and part.startswith("~~") and part.endswith("~~"):
            out.append((part[2:-2], False, False, True))      # tosca
        elif len(part) >= 3 and part.startswith("_") and part.endswith("_"):
            out.append((part[1:-1], False, True, False))      # miring
        else:
            out.append((part, False, False, False))
    return out


# ---------- PEMILIH GAYA DESAIN (6 gaya cover Nikah Institute) ----------
# key internal: dark / card / duotone / doodle / doodle_top
STYLE_MAP = {
    "1": "dark", "foto gelap": "dark", "gelap": "dark",
    "2": "card", "kartu putih": "card", "kartu": "card", "putih": "card",
    "3": "duotone", "tosca duotone": "duotone", "duotone": "duotone", "tosca": "duotone",
    "4": "dark", "foto stabilo": "dark", "stabilo": "dark",
    "5": "doodle", "doodle terang": "doodle", "doodle": "doodle",
    "6": "doodle_top", "doodle teks atas": "doodle_top", "doodle atas": "doodle_top",
}
def parse_style(s):
    """Terima teks/angka dari kolom 'Gaya Desain' -> key gaya internal. Default 'dark'."""
    k = re.sub(r"[^\w\s]", " ", (s or "").strip().lower())
    k = re.sub(r"\s+", " ", k).strip()
    if not k:
        return "dark"
    if k in STYLE_MAP:
        return STYLE_MAP[k]
    m = re.search(r"\b([1-6])\b", k)           # "3", "3 tosca duotone", "gaya 3", "style 2"
    if m and m.group(1) in STYLE_MAP:
        return STYLE_MAP[m.group(1)]
    for key, val in STYLE_MAP.items():
        if not key.isdigit() and key in k:
            return val
    return "dark"

def extract_gaya(block):
    """Deteksi penanda [gaya: X] di teks slide (override per-slide). Kembalikan (teks_bersih, gaya|None)."""
    gaya = None
    m = re.search(r"\[\s*gaya\s*:\s*([^\]]+)\]", block, re.I)
    if m:
        gaya = parse_style(m.group(1))
        block = block[:m.start()] + block[m.end():]
    return block.strip(), gaya


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

# ---------- FITUR: BAWA GAMBAR SENDIRI (mis. hasil doodle Gemini) ----------
def extract_custom_image(block):
    """Deteksi penanda gambar custom di teks slide.
    Penanda yg didukung (huruf besar/kecil bebas):
      [img]https://.../foto.png[/img]  -> pakai gambar dari URL itu
      [img]                            -> pakai gambar tempelan Notion berikutnya
      [img:light] / [gambar:terang]    -> paksa teks GELAP (buat background terang)
      [img:dark]  / [gambar:gelap]     -> paksa teks PUTIH (buat background gelap)
    Kembalikan (teks_bersih, url_atau_None, minta_custom, tema_paksa_atau_None).
    """
    wants, url, forced = False, None, None
    mt = re.search(r"\[\s*(?:img|gambar)\s*:\s*(light|dark|terang|gelap)\s*\]", block, re.I)
    if mt:
        wants = True
        v = mt.group(1).lower()
        forced = "light" if v in ("light", "terang") else "dark"
        block = block[:mt.start()] + block[mt.end():]
    m = re.search(r"\[\s*(?:img|gambar)\s*\]", block, re.I)
    if m:
        wants = True
        after = block[m.end():]
        um = re.match(r"\s*(https?://[^\s\[\]]+)", after)   # URL berhenti sebelum '[' (penutup [/img])
        if um:
            url = um.group(1).rstrip(').,')
            after = after[um.end():]
        after = re.sub(r"^\s*\[\s*/\s*(?:img|gambar)\s*\]", "", after, flags=re.I)
        block = block[:m.start()] + after
    return block.strip(), url, wants, forced

def download_image_bytes(url):
    """Ambil gambar dari URL (buat fitur bawa-gambar-sendiri). None kalau gagal/bukan gambar."""
    try:
        data = requests.get(url, headers=UA, timeout=60).content
        Image.open(io.BytesIO(data)).verify()
        return data
    except Exception as e:
        print("    ! gagal ambil gambar custom:", e)
        return None

def _mean_lum(im):
    """Rata-rata terang (0 gelap .. 255 putih)."""
    try:
        return ImageStat.Stat(im.convert("L")).mean[0]
    except Exception:
        return 0.0

def pick_theme(base_img, forced=None):
    """Tentukan tema teks: 'light' (background terang -> teks gelap) / 'dark' (teks putih)."""
    if forced:
        return forced
    return "light" if _mean_lum(base_img) >= 140 else "dark"

# ---------- filter per-GAYA ----------
def duotone_teal(im):
    """GAYA 3: ubah foto jadi duotone tosca gelap (grayscale -> gradasi tosca)."""
    g = im.convert("L")
    dark, light = (6, 28, 30), (150, 212, 208)   # tosca gelap -> tosca terang
    def lut(a, b):
        return [int(a + (b - a) * i / 255) for i in range(256)]
    r = g.point(lut(dark[0], light[0]))
    gg = g.point(lut(dark[1], light[1]))
    b = g.point(lut(dark[2], light[2]))
    return Image.merge("RGB", (r, gg, b))

def make_card_bg(photo):
    """GAYA 2: foto di atas (~58%) + kartu krem/putih di bawah (~42%).
    Teks nanti ditaruh di kartu (warna gelap)."""
    W, H = CANVAS_W, CANVAS_H
    split = int(H * 0.58)
    canvas = Image.new("RGB", (W, H), (248, 247, 250))     # kartu krem/putih
    ph = photo if photo.size == (W, H) else photo.resize((W, H), Image.LANCZOS)
    canvas.paste(ph.crop((0, 0, W, split)), (0, 0))        # foto area atas
    # gradasi tipis di pucuk foto biar logo/tagline (putih) tetap kebaca
    strip_h = int(0.16 * H)
    overlay = Image.new("RGBA", (W, strip_h), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    for y in range(strip_h):
        a = int(130 * (1 - y / strip_h))
        od.line([(0, y), (W, y)], fill=(0, 0, 0, a))
    canvas.paste(Image.alpha_composite(
        canvas.crop((0, 0, W, strip_h)).convert("RGBA"), overlay).convert("RGB"), (0, 0))
    return canvas

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
    "card":   dict(hx=0.07, hy=0.615, hw=0.86, bx=0.07, by=0.74, bw=0.86, align="left"),
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


def download_logo_path(url=None):
    url = url if url is not None else LOGO_URL
    if not url:
        return None
    try:
        data = requests.get(url, headers=UA, timeout=30).content
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
        segs = rich_segments(line) if highlight else [(line, False, False, False)]
        for seg, is_hl, is_it, is_tl in (segs or [("", False, False, False)]):
            run = p.add_run()
            run.text = seg
            f = run.font
            f.size = Pt(size_pt); f.bold = bold or is_hl or is_tl; f.name = font
            f.italic = is_it
            if is_hl:
                f.color.rgb = ACCENT          # ungu
            elif is_tl:
                f.color.rgb = ACCENT2         # tosca
            else:
                f.color.rgb = hex_rgb(color_hex)
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

def add_dots_pptx(slide, index, total, color_hex):
    """Titik-titik halaman di bawah-tengah (editable). index = slide aktif (0-based)."""
    if total <= 1:
        return
    d = 0.013           # diameter (fraksi lebar) -> pakai fx utk w & h biar bulat
    gap = d * 1.9
    start = 0.5 - (total - 1) * gap / 2
    y = 0.90
    for k in range(total):
        x = start + k * gap
        shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, fx(x), fy(y), fx(d), fx(d))
        shp.fill.solid()
        shp.fill.fore_color.rgb = ACCENT if k == index else hex_rgb(color_hex)
        shp.line.fill.background(); shp.shadow.inherit = False

def _pp_align(a):
    return PP_ALIGN.CENTER if a == "center" else PP_ALIGN.LEFT

def _slide_header(slide, logo_path, txt="FFFFFF"):
    if logo_path:
        try:
            slide.shapes.add_picture(logo_path, fx(0.06), fy(0.045), height=fy(0.045))
        except Exception:
            pass
    if BRAND_TAGLINE:
        _add_text(slide, 0.52, 0.045, 0.42, 0.08, BRAND_TAGLINE, 14, True, BODY_FONT, txt, align=PP_ALIGN.RIGHT)

def build_pptx(slides_data, out_path, logo_path, logo_dark_path=None):
    prs = Presentation()
    prs.slide_width = Emu(EMU_W); prs.slide_height = Emu(EMU_H)
    blank = prs.slide_layouts[6]
    for s in slides_data:
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=prs.slide_height)
        theme = s.get("theme", "dark")
        if theme == "card":
            tag_txt, body_txt, dark_logo = "FFFFFF", "222222", False
        elif theme == "light":
            tag_txt, body_txt, dark_logo = "222222", "222222", True
        else:
            tag_txt, body_txt, dark_logo = "FFFFFF", "FFFFFF", False
        hdr_logo = logo_dark_path if (dark_logo and logo_dark_path) else logo_path
        _slide_header(slide, hdr_logo, txt=tag_txt)

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
                      head_pt(zone, cover), True, HEADLINE_FONT, body_txt, align=al, highlight=True)
        if s["body"]:
            _add_text(slide, L["bx"], body_top(zone, cover), L["bw"], 0.34, s["body"],
                      body_pt(zone), False, BODY_FONT, body_txt, align=al, highlight=True)
        if FOOTER_TEXT:
            _add_text(slide, 0.06, 0.93, 0.58, 0.055, FOOTER_TEXT, 11, False, BODY_FONT, body_txt)
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

def _svg_text(text, x, y, width, size, bold, font, anchor="start", highlight=False, gap=1.25,
              color="FFFFFF"):
    """
    Teks OTOMATIS TURUN BARIS (wrap) biar PAS di dalam frame (nggak meleber).
    Tiap baris hasil wrap = satu <text> (anchor start, x digeser manual utk
    center/right). Stabilo (==) & miring (_) tetap jalan inline.
    'color' = warna teks biasa (FFFFFF utk bg gelap, 222222 utk bg terang).
    """
    weight = "700" if bold else "400"
    # huruf tebal lebih lebar -> pakai faktor lebih besar biar wrap lebih awal (nggak meleber)
    char_factor = 0.63 if bold else 0.56
    max_chars = max(6, int(width / (size * char_factor)))
    lh = int(size * gap)

    def words_of(line):
        res = []
        for s, hl, it, tl in (rich_segments(line) if highlight else [(line, False, False, False)]):
            for w in s.split(" "):
                if w != "":
                    res.append((w, hl, it, tl))
        return res

    out = []
    yy = y
    for para in text.split("\n"):
        ws = words_of(para)
        if not ws:
            yy += lh; continue
        # bagi jadi beberapa baris sesuai lebar frame
        vlines, cur, cur_len = [], [], 0
        for w, hl, it, tl in ws:
            add = len(w) + (1 if cur else 0)
            if cur and cur_len + add > max_chars:
                vlines.append(cur); cur, cur_len = [], 0; add = len(w)
            cur.append((w, hl, it, tl)); cur_len += add
        if cur:
            vlines.append(cur)
        for vl in vlines:
            line_str = " ".join(w for w, _, _, _ in vl)
            lw = _est_w(line_str, size)
            if anchor == "middle":
                lx = int(x - lw / 2)
            elif anchor == "end":
                lx = int(x - lw)
            else:
                lx = int(x)
            frag = ""
            for k, (w, hl, it, tl) in enumerate(vl):
                prefix = " " if k > 0 else ""
                tok = _svg_escape(prefix + w)
                if hl:
                    frag += f'<tspan fill="#{ACCENT_COLOR}" font-weight="700">{tok}</tspan>'
                elif tl:
                    frag += f'<tspan fill="#{ACCENT2_COLOR}" font-weight="700">{tok}</tspan>'
                elif it:
                    frag += f'<tspan font-style="italic">{tok}</tspan>'
                else:
                    frag += tok
            base = yy + size
            out.append(f'<text x="{lx}" y="{int(base)}" text-anchor="start" xml:space="preserve" '
                       f'fill="#{color}" font-family="{_svg_escape(font)}, Arial, sans-serif" '
                       f'font-size="{size}" font-weight="{weight}">{frag}</text>')
            yy += lh
    return "\n".join(out)


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

def _dots_svg(xo, index, total, color):
    """Titik-titik indikator halaman di bawah-tengah. index = slide aktif (0-based)."""
    if total <= 1:
        return ""
    cy = int(0.905 * CANVAS_H)
    r = max(4, int(0.007 * CANVAS_W))
    gap = r * 3
    span = (total - 1) * gap
    start = xo + CANVAS_W // 2 - span // 2
    out = []
    for k in range(total):
        cx = start + k * gap
        if k == index:
            out.append(f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="#{ACCENT_COLOR}"/>')
        else:
            out.append(f'<circle cx="{cx}" cy="{cy}" r="{int(r*0.78)}" fill="#{color}" fill-opacity="0.35"/>')
    return "\n".join(out)

def build_svg(slides_data, out_path, logo_path, logo_dark_path=None):
    total = len(slides_data)
    W = CANVAS_W * total; H = CANVAS_H
    logo_uri = _img_data_uri(logo_path) if logo_path else None
    logo_dark_uri = _img_data_uri(logo_dark_path) if logo_dark_path else None
    defs_all, parts = [], []
    for idx, s in enumerate(slides_data):
        xo = idx * CANVAS_W
        theme = s.get("theme", "dark")
        # warna teks per-tema: tag_txt = logo/tagline (atas), body_txt = judul/isi/footer
        if theme == "card":          # foto atas (putih) + kartu bawah (gelap)
            tag_txt, body_txt, scrim_on, dark_logo = "FFFFFF", "222222", False, False
        elif theme == "light":       # bg terang -> semua teks gelap
            tag_txt, body_txt, scrim_on, dark_logo = "222222", "222222", False, True
        else:                        # dark / duotone
            tag_txt, body_txt, scrim_on, dark_logo = "FFFFFF", "FFFFFF", True, False
        # foto BERSIH (tanpa masking dibakar) -> masking ditaruh sbg vektor
        photo = s.get("clean_path") or s["image_path"]
        parts.append(f'<image x="{xo}" y="0" width="{CANVAS_W}" height="{H}" '
                     f'preserveAspectRatio="xMidYMid slice" href="{_img_data_uri(photo)}"/>')
        # masking vektor cuma utk slide gelap/duotone
        if s.get("kind") != "cta" and scrim_on:
            d, r = _scrim_svg(s.get("zone", "bottom"), xo, idx)
            defs_all.append(d); parts.append(r)
        # logo: versi gelap kalau tema terang (kalau ada)
        use_logo = logo_dark_uri if (dark_logo and logo_dark_uri) else logo_uri
        if use_logo:
            parts.append(f'<image x="{xo + int(0.06*CANVAS_W)}" y="{int(0.05*H)}" '
                         f'height="{int(0.05*H)}" href="{use_logo}"/>')
        if BRAND_TAGLINE:
            parts.append(_svg_text(BRAND_TAGLINE, xo + int(0.94*CANVAS_W), int(0.05*H),
                                   int(0.4*CANVAS_W), 26, True, BODY_FONT, anchor="end", color=tag_txt))

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
                                   anchor=anc, highlight=True, color=body_txt))
        if s["body"]:
            parts.append(_svg_text(s["body"], bx, int(body_top(zone, cover) * H),
                                   int(L["bw"] * CANVAS_W), body_px(zone), False, BODY_FONT,
                                   anchor=anc, highlight=True, color=body_txt))
        if FOOTER_TEXT:
            parts.append(_svg_text(FOOTER_TEXT, xo + int(0.06*CANVAS_W), int(0.925*H),
                                   int(0.6*CANVAS_W), 20, False, BODY_FONT, color=body_txt))
        if s.get("total", total) > 1 and CTA_TEXT:
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

def tg_delete_webhook():
    """Pastikan nggak ada webhook nyangkut. Webhook bikin getUpdates MANDEK (balik kosong),
    jadi tombol/ketikan user nggak pernah kebaca walau bot tetap bisa KIRIM pesan.
    drop_pending_updates=false -> pesan yg ketahan TIDAK dibuang (langsung kebaca sesudahnya)."""
    try:
        r = requests.post(f"{TG_BASE}/deleteWebhook",
                          data={"drop_pending_updates": "false"}, timeout=20)
        j = r.json()
        print(f"    deleteWebhook: ok={j.get('ok')} {j.get('description','')}")
    except Exception as e:
        print("    ! deleteWebhook error:", e)

def tg_get_updates(offset):
    try:
        r = requests.get(f"{TG_BASE}/getUpdates",
                         params={"offset": offset, "timeout": 0, "allowed_updates": json.dumps(["callback_query", "message"])},
                         timeout=40)
        j = r.json()
        if not j.get("ok"):
            print(f"    ! getUpdates NOT ok: {str(j)[:200]}")
        return j.get("result", []) if j.get("ok") else []
    except Exception as e:
        print("    ! getUpdates error:", e)
        return []

def tg_answer_callback(cb_id, text=""):
    # notif kecil di tombol; kalau gagal (query kedaluwarsa krn batch) abaikan diam-diam
    try:
        requests.post(f"{TG_BASE}/answerCallbackQuery",
                      data={"callback_query_id": cb_id, "text": text[:180]}, timeout=20)
    except Exception:
        pass


# ======================================================================
# 7b. SISTEM BELAJAR — memori + feedback (approve/reject + alasan)
# ======================================================================
DEFAULT_MEMORY = {
    "preferences": {"head_scale": 1.0, "body_scale": 1.0, "scrim": 1.0, "soft": 0.0},
    "pending": {},        # token -> {title, page_id, ts, revision}
    "resolved": [],       # token yang sudah dinilai (anti dobel)
    "awaiting_text": None,  # token yang menunggu alasan ketik
    "revise_queue": {},   # page_id -> {title, revision, reasons[]} konten yg minta direvisi
    "log": [],            # riwayat {ts, title, status, reason}
    "stats": {"approved": 0, "rejected": 0, "weekly": []},
    "tg_offset": 0,
}

MAX_REVISI = int(os.environ.get("MAX_REVISI") or 10)   # batas aman auto-revisi per konten

# tombol alasan saat Reject: (label, kode)
REASON_BUTTONS = [
    ("Teks kegedean", "tbig"),
    ("Teks kekecilan", "tsmall"),
    ("Teks susah kebaca", "hard"),
    ("Overlay kegelapan", "dark"),
    ("Gambar kurang aesthetic", "img"),
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


# ---------- OTAK AI: Gemini (utama) + Groq (cadangan). Dua-duanya gratis ----------
GEMINI_API_KEY  = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL    = os.environ.get("GEMINI_MODEL") or "gemini-3.8-flash"
GROQ_API_KEY    = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL      = os.environ.get("GROQ_MODEL") or "openai/gpt-oss-20b"
AI_CHAT_ENABLED = ((os.environ.get("AI_CHAT_ENABLED") or "true").lower() == "true"
                   and (bool(GEMINI_API_KEY) or bool(GROQ_API_KEY)))
# cek mandiri pakai 'mata' AI (Gemini vision). gratis (gambar sbg INPUT), cuma butuh Gemini key.
VISION_CHECK_ENABLED = ((os.environ.get("VISION_CHECK_ENABLED") or "true").lower() == "true"
                        and bool(GEMINI_API_KEY))

def _gemini_json(prompt):
    if not GEMINI_API_KEY:
        return None
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    body = {"contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.3, "responseMimeType": "application/json"}}
    try:
        r = requests.post(url, json=body, timeout=45)
        if r.status_code != 200:
            print(f"    ! Gemini {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])
    except Exception as e:
        print("    ! Gemini gagal:", e); return None

def _groq_json(prompt):
    if not GROQ_API_KEY:
        return None
    try:
        r = requests.post("https://api.groq.com/openai/v1/chat/completions",
                          headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
                          json={"model": GROQ_MODEL, "temperature": 0.3,
                                "response_format": {"type": "json_object"},
                                "messages": [{"role": "user", "content": prompt}]},
                          timeout=45)
        if r.status_code != 200:
            print(f"    ! Groq {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["choices"][0]["message"]["content"])
    except Exception as e:
        print("    ! Groq gagal:", e); return None

AI_PRIMARY = (os.environ.get("AI_PRIMARY") or "groq").lower()   # 'groq' (stabil) atau 'gemini'

def ai_json(prompt):
    """Minta jawaban JSON ke AI. Default: Groq dulu (lebih stabil), Gemini cadangan.
    Bisa dibalik lewat secret AI_PRIMARY=gemini."""
    if not AI_CHAT_ENABLED:
        return None
    if AI_PRIMARY == "gemini":
        first, second, second_name = _gemini_json, _groq_json, "Groq"
    else:
        first, second, second_name = _groq_json, _gemini_json, "Gemini"
    data = first(prompt)
    if data is None:
        data = second(prompt)   # cadangan kalau yang utama ngadat
        if data is not None:
            print(f"    (pakai {second_name} sebagai cadangan)")
    return data

AI_PROMPT = """Kamu asisten desain untuk bot konten Instagram (brand Nikah Institute).
User memberi perintah/feedback (bahasa Indonesia atau Inggris) untuk merevisi sebuah desain.
Terjemahkan jadi penyesuaian angka. Setelan sekarang (1.0 = normal):
- head_scale: ukuran JUDUL (batas 0.7-1.3)
- body_scale: ukuran teks BODY (batas 0.7-1.3)
- scrim: kegelapan overlay di belakang teks (batas 0.6-1.6; naik = lebih gelap = teks lebih terbaca)
- soft: kesan film/lembut pada gambar (batas 0.0-0.4; naik = lebih lembut/pudar; turun = lebih cerah/tajam)
PENTING: hampir SEMUA masukan user = permintaan revisi -> set action "revise".
Keluhan/komentar sekecil apapun soal desain (foto/teks/warna/overlay) = "revise".
Contoh yang HARUS "revise": "gambar kurang bagus", "kurang aesthetic", "ganti foto",
"fotonya jelek", "judul kegedean", "kurang jelas", "overlay kurang gelap".
Kalau user komentarin/minta ganti foto, set action "revise" (foto otomatis diganti baru saat revisi).
Pakai action "none" HANYA kalau pesan jelas cuma sapaan/basa-basi/terima kasih/pertanyaan umum
(mis. "halo", "makasih", "lagi apa") — BUKAN komentar soal desain.
Balas HANYA JSON valid, tanpa teks lain:
{"head_scale_delta": <angka -0.12..0.12>, "body_scale_delta": <angka>, "scrim_delta": <angka>, "soft_delta": <angka>, "note": "<ringkasan singkat dalam bahasa Indonesia, maks 12 kata>", "action": "revise" atau "none"}
Pesan user: "%s" """

def ai_parse_command(text):
    """Ubah kalimat bebas jadi penyesuaian angka (Gemini/Groq). None kalau gagal/mati."""
    return ai_json(AI_PROMPT % text[:400])

def apply_ai(prefs, data):
    """Terapkan delta dari AI (dibatasi aman). Kembalikan daftar yg berubah."""
    changed = []
    plan = [("head_scale_delta", "head_scale", 0.7, 1.3),
            ("body_scale_delta", "body_scale", 0.7, 1.3),
            ("scrim_delta", "scrim", 0.6, 1.6),
            ("soft_delta", "soft", 0.0, 0.4)]
    for key, knob, lo, hi in plan:
        try:
            d = float(data.get(key, 0) or 0)
        except Exception:
            d = 0.0
        d = _clamp(d, -0.15, 0.15)   # batasi per-perintah biar nggak loncat jauh
        if abs(d) >= 0.01:
            prefs[knob] = _clamp(prefs[knob] + d, lo, hi)
            changed.append(knob)
    return changed

def _latest_pending(mem):
    best = None
    for tok, rec in mem.get("pending", {}).items():
        if best is None or rec.get("ts", 0) > best[1].get("ts", 0):
            best = (tok, rec)
    return best

def ai_image_keywords(slide_texts):
    """1 panggilan AI (Gemini/Groq) -> kata kunci FOTO (Inggris, aesthetic) per slide. None kalau gagal."""
    if not AI_CHAT_ENABLED or not slide_texts:
        return None
    joined = "\n".join(f"{i+1}. {(t or '')[:160]}" for i, t in enumerate(slide_texts))
    prompt = (
        "Kamu art director untuk brand konseling pernikahan (gaya editorial/aesthetic luar negeri, nuansa sinematik & soft).\n"
        "Untuk TIAP slide di bawah, buat 1 kata kunci pencarian FOTO STOK dalam BAHASA INGGRIS (3-5 kata), "
        "fokus ke suasana/visual (bukan terjemahan harfiah teks), hindari tulisan/logo, utamakan natural light & tone lembut.\n"
        'Balas HANYA JSON object: {"keywords": ["...", "..."]} dengan panjang array PERSIS sama dengan jumlah slide.\n\n' + joined)
    data = ai_json(prompt)
    try:
        arr = (data or {}).get("keywords")
        if isinstance(arr, list) and arr:
            return [str(x) for x in arr]
    except Exception:
        pass
    return None

def _gemini_vision_json(prompt, image_path):
    """Kirim gambar + pertanyaan ke Gemini (mode 'mata'/vision), minta jawaban JSON.
    Gambar dikecilin dulu (max 640px, JPEG) biar payload kecil & cepat."""
    if not GEMINI_API_KEY:
        return None
    try:
        im = Image.open(image_path).convert("RGB")
        im.thumbnail((640, 640))
        buf = io.BytesIO(); im.save(buf, "JPEG", quality=80)
        b64 = base64.b64encode(buf.getvalue()).decode()
    except Exception as e:
        print("    ! vision: gagal siapin gambar:", e); return None
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    body = {"contents": [{"parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": b64}}]}],
            "generationConfig": {"temperature": 0.2, "responseMimeType": "application/json"}}
    try:
        r = requests.post(url, json=body, timeout=60)
        if r.status_code != 200:
            print(f"    ! Gemini vision {r.status_code}: {r.text[:120]}")
            return None
        return json.loads(r.json()["candidates"][0]["content"]["parts"][0]["text"])
    except Exception as e:
        print("    ! Gemini vision gagal:", e); return None

def ai_vision_check(image_path, zone, theme):
    """Cek mandiri: kirim BACKGROUND slide ke Gemini, tanya apakah teks bakal mudah dibaca.
    Kembalikan {ok: bool, issue: str} atau None kalau mati/gagal."""
    if not VISION_CHECK_ENABLED:
        return None
    zmap = {"top": "atas", "bottom": "bawah", "left": "kiri", "right": "kanan", "center": "tengah"}
    warna = "putih" if theme == "dark" else "gelap (hampir hitam)"
    prompt = (
        "Kamu reviewer desain konten Instagram. Ini BACKGROUND sebuah slide (teks belum ditaruh).\n"
        f"Rencana: teks judul + isi diletakkan di area {zmap.get(zone, zone)}, warna teks {warna}.\n"
        "Nilai apakah di area itu teks akan MUDAH DIBACA: kontras cukup, area tidak terlalu ramai/berpola, "
        "dan tidak menutupi wajah/objek penting.\n"
        'Balas HANYA JSON: {"ok": true atau false, "issue": "<masalah singkat bahasa Indonesia maks 12 kata; '
        'kosongkan kalau sudah ok>"}'
    )
    data = _gemini_vision_json(prompt, image_path)   # 1x aja (retry malah boros kuota)
    if not isinstance(data, dict):
        return None
    return {"ok": bool(data.get("ok", True)), "issue": str(data.get("issue") or "").strip()}

def _queue_revision(mem, rec, reason_label):
    """Masukkan konten ke antrian revisi (biar run berikutnya di-generate ulang)."""
    page_id = ((rec or {}).get("page_id") or "").replace("-", "")   # normalisasi kunci
    if not page_id:
        return False
    rev = (rec or {}).get("revision", 0)
    q = mem.setdefault("revise_queue", {})
    entry = q.get(page_id) or {"title": (rec or {}).get("title", "desain"), "revision": rev, "reasons": []}
    entry["title"] = (rec or {}).get("title", entry.get("title", "desain"))
    entry["revision"] = rev + 1
    entry["reasons"].append(reason_label)
    entry["reasons"] = entry["reasons"][-8:]
    q[page_id] = entry
    return True

def _mark_resolved(mem, tok):
    res = mem.setdefault("resolved", [])
    if tok not in res:
        res.append(tok)
    if len(res) > 500:
        del res[:len(res) - 500]

def process_feedback(mem):
    """Baca pencetan tombol sejak run terakhir, update stats + knob.
    Knob diterapkan BERDASARKAN alasan yg dipencet (tidak tergantung 'pending'),
    biar tetap jalan walau catatan pending hilang. Anti-dobel pakai daftar 'resolved'."""
    prefs = mem["preferences"]
    tg_delete_webhook()   # jaga-jaga: webhook nyangkut bikin tombol/ketikan nggak kebaca
    updates = tg_get_updates(mem.get("tg_offset", 0) + 1)
    changed = []
    processed = 0
    for up in updates:
        mem["tg_offset"] = max(mem.get("tg_offset", 0), up.get("update_id", 0))

        # ---- pesan teks: perintah stats / reset / revisi bebas (AI) ----
        msg = up.get("message")
        if msg:
            text = (msg.get("text") or "").strip()
            low = text.lower()
            mem["awaiting_text"] = None   # fitur lama nggak dipakai lagi
            if low in ("stats", "/stats", "statistik"):
                a = mem["stats"]["approved"]; r = mem["stats"]["rejected"]; tot = a + r
                rate = int(100 * a / tot) if tot else 0
                tg_message(f"📊 Statistik belajar Nikah:\n"
                           f"Approve {a} / Reject {r}  (approval rate {rate}%).\n"
                           f"{prefs_summary(mem['preferences'])}")
                continue
            if low in ("/reset", "reset belajar"):
                mem["preferences"] = dict(DEFAULT_MEMORY["preferences"])
                prefs = mem["preferences"]
                tg_message("🔄 Setelan belajar direset ke awal (netral). Statistik tetap tersimpan.")
                continue
            # ---- CHAT AI: perintah revisi bebas (ketik kalimat biasa) ----
            if text:
                data = ai_parse_command(text)
                # PENGAMAN: AI bilang "none" tapi teks jelas keluhan + ada desain nunggu -> tetap REVISI
                _rev_kata = ("ganti", "ubah", "revisi", "perbaiki", "kurang", "jelek", "buram",
                             "norak", "kegedean", "kekecilan", "lebih ", "jangan", "terlalu")
                if (data and data.get("action") == "none"
                        and any(k in low for k in _rev_kata) and _latest_pending(mem)):
                    data["action"] = "revise"
                if data and data.get("action") == "none":
                    tg_message("Oke, dicatat 🙂 Kalau mau gw revisi, kasih tau yang perlu diubah ya "
                               "(misal: “judul kekecilan”, “overlay kurang gelap”, “ganti foto lebih cerah”).")
                    continue
                note = (data or {}).get("note") or text[:80]
                ch = apply_ai(prefs, data) if data else []
                latest = _latest_pending(mem)
                if latest:
                    tok, rec = latest
                    _queue_revision(mem, rec, "chat: " + note)
                    mem["pending"].pop(tok, None); _mark_resolved(mem, tok)
                    mem["stats"]["rejected"] += 1
                    extra = ("\n" + prefs_summary(prefs)) if ch else ""
                    tail = "" if data else " (AI lagi nggak aktif, jadi gw revisi dgn foto baru aja)"
                    tg_message(f"🧠 Paham: {note}. Gw revisi '{rec.get('title','desain')}' ya{tail}.{extra}")
                else:
                    if ch:
                        tg_message(f"🧠 Paham: {note}. Setelan disesuaikan, kepakai di desain berikutnya.\n{prefs_summary(prefs)}")
                    else:
                        tg_message("Belum ada desain yang bisa direvisi. Kirim konten dulu ya 🙂")
                processed += 1
                continue
            continue

        # ---- pencetan tombol (callback) ----
        cb = up.get("callback_query")
        if not cb:
            continue
        data = cb.get("data", "")
        cb_id = cb.get("id", "")
        parts = data.split(":")
        kind = parts[0]
        tok = parts[1] if len(parts) > 1 else ""
        if tok and tok in mem.get("resolved", []):
            tg_answer_callback(cb_id, "Sudah dinilai sebelumnya 👍"); continue

        # page_id & code diambil dari DATA TOMBOL (nggak gantung ke pending yg bisa hilang)
        if kind == "rs":
            code = parts[2] if len(parts) > 2 else ""
            pid_cb = parts[3] if len(parts) > 3 else ""
        else:
            code = ""
            pid_cb = parts[2] if len(parts) > 2 else ""
        pend = mem["pending"].get(tok, {})
        rec = {"title": pend.get("title", "desain"),
               "page_id": pid_cb or pend.get("page_id"),
               "revision": pend.get("revision", 0)}

        if kind == "a":                      # approve (1x pencet) -> FINAL
            mem["pending"].pop(tok, None)
            mem["stats"]["approved"] += 1
            _mark_resolved(mem, tok); processed += 1
            pid_norm = (rec.get("page_id") or "").replace("-", "")
            if pid_norm:
                mem.get("revise_queue", {}).pop(pid_norm, None)   # batal revisi, udah oke
            mem["log"].append({"ts": int(time.time()), "title": rec["title"],
                               "status": "approved", "reason": ""})
            tg_answer_callback(cb_id, "✅ Disimpan sebagai contoh bagus!")
        elif kind == "rs":                   # reject + alasan (1x pencet) -> knob + antri revisi
            if code == "other":   # tombol lama: sekarang cukup KETIK langsung
                tg_answer_callback(cb_id, "Ketik aja perintah revisimu langsung ya 🙂")
                tg_message("💬 Ketik aja apa yang mau diubah (misal: “judul kekecilan, foto lebih cerah”) — "
                           "langsung gw proses, nggak usah pencet tombol ini.")
            else:
                mem["stats"]["rejected"] += 1
                desc = apply_reason(prefs, code)
                mem["pending"].pop(tok, None)
                _mark_resolved(mem, tok); processed += 1
                ok = _queue_revision(mem, rec, REASON_LABEL.get(code, code))
                print(f"  reject({code}) pid={rec.get('page_id')} queued={ok}")
                mem["log"].append({"ts": int(time.time()), "title": rec["title"],
                                   "status": "rejected", "reason": REASON_LABEL.get(code, code)})
                changed.append(desc)
                tg_answer_callback(cb_id, f"Paham. {desc}. Gw revisi ya.")
        elif kind == "r":                    # kompat tombol lama (Reject 2-langkah) -> abaikan halus
            tg_answer_callback(cb_id, "Pakai tombol alasan di desain baru ya 🙏")

    # log diagnostik di console (biar kelihatan apa yg kebaca)
    print(f"  feedback: {len(updates)} update, {processed} diproses, "
          f"revise_queue={len(mem.get('revise_queue', {}))}, offset={mem.get('tg_offset')}")

    # laporan ke Telegram biar KELIHATAN bot baca feedback
    if processed > 0:
        head = f"🔎 Feedback terbaca: {processed} pencetan diproses."
        if changed:
            head += "\n🧠 Bot menyesuaikan diri: " + "; ".join(changed) + "."
        head += "\n" + prefs_summary(prefs)
        tg_message(head)
    mem["_last_processed"] = processed
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
def process_page(page, workdir, mem=None, revision=None):
    props = page["properties"]
    page_id = page.get("id", "")
    title = get_title(props) or "Tanpa Judul"
    # 'Bentuk Konten' bisa tipe Select/Status/Teks -> baca fleksibel; default Carousel (aman, nggak motong slide)
    fmt = (read_property(props, FORMAT_PROPERTY, "select")
           or read_property(props, FORMAT_PROPERTY, "status")
           or read_property(props, FORMAT_PROPERTY, "rich_text")
           or "Carousel")
    content = read_property(props, CONTENT_PROPERTY, "rich_text")
    keywords = read_property(props, KEYWORD_PROPERTY, "rich_text")
    # GAYA DESAIN (buat cover) — baca dari kolom dropdown; default 'dark' (Foto Gelap)
    gaya_raw = (read_property(props, GAYA_PROPERTY, "select")
                or read_property(props, GAYA_PROPERTY, "status")
                or read_property(props, GAYA_PROPERTY, "rich_text") or "")
    content_style = parse_style(gaya_raw)
    print(f"  [gaya cover: {content_style}  (dari '{gaya_raw or '-'}')]")

    if not content.strip():
        raise ValueError("Kolom 'Konten' kosong.")

    slide_blocks = split_slides(content)
    if "single" in fmt.lower():
        slide_blocks = slide_blocks[:1]
    keyword_blocks = split_slides(keywords)
    # kata kunci gambar pintar (AI) — 1 panggilan utk semua slide; None kalau AI mati/gagal
    ai_kws = ai_image_keywords(slide_blocks)
    if ai_kws:
        print(f"  [AI keywords: {ai_kws}]")

    print(f"  -> '{title}' | {fmt} | {len(slide_blocks)} slide")

    slides_data, preview_paths, credits = [], [], []
    vision_notes = []           # catatan cek mandiri (AI vision)
    used_ids = set()
    total = len(slide_blocks)

    # fitur "bawa gambar sendiri": daftar gambar tempelan Notion diambil sekali (lazy),
    # lalu dipakai urut utk slide yang ditandai [img] tanpa URL.
    _att = {"list": None, "idx": 0}
    def next_notion_image():
        if _att["list"] is None:
            _att["list"] = notion_page_images(page_id)
            if _att["list"]:
                print(f"  [attachment Notion: {len(_att['list'])} gambar ditemukan]")
        lst = _att["list"] or []
        if _att["idx"] < len(lst):
            u = lst[_att["idx"]]; _att["idx"] += 1
            return u
        return None

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

        # --- FITUR: bawa gambar sendiri + penanda gaya ---
        clean_block, cimg_url, wants_custom, forced_theme = extract_custom_image(block)
        clean_block, slide_gaya = extract_gaya(clean_block)
        headline, body = headline_and_body(clean_block)
        # gaya efektif: penanda per-slide > gaya cover (slide 1) > default 'dark' (slide lain)
        eff_style = slide_gaya or (content_style if i == 1 else "dark")

        base = None
        theme = "dark"
        is_custom = False
        if wants_custom:
            cdata = download_image_bytes(cimg_url) if cimg_url else None
            if cdata is None:                       # nggak ada URL / URL gagal -> coba tempelan Notion
                u = next_notion_image()
                if u:
                    cdata = download_image_bytes(u)
            if cdata is not None:
                try:
                    cim = Image.open(io.BytesIO(cdata)).convert("RGB")
                    base = crop_canvas(cim)         # gambar sendiri: JANGAN di-film-grade (biar doodle tetap bersih)
                    theme = pick_theme(base, forced_theme)
                    is_custom = True
                    credits.append(f"Slide {i}: (gambar sendiri / custom, tema {theme})")
                except Exception as e:
                    print("    ! gambar custom rusak:", e); base = None
            if base is None:
                print(f"    ! gambar custom slide {i} gagal -> pakai foto stok")

        if base is None:
            # jalur normal: cari foto stok
            explicit = keyword_blocks[i - 1].strip() if (i - 1) < len(keyword_blocks) else ""
            if explicit:
                q = keyword_for(i - 1, keyword_blocks, headline, body)
            elif ai_kws and (i - 1) < len(ai_kws) and ai_kws[i - 1].strip():
                q = ai_kws[i - 1].strip()
            else:
                q = keyword_for(i - 1, keyword_blocks, headline, body)
            data, credit = image_pick(q, used_ids)
            if data:
                im = Image.open(io.BytesIO(data)).convert("RGB")
                if credit:
                    credits.append(f"Slide {i}: {credit}")
            else:
                im = make_placeholder()
                credits.append(f"Slide {i}: (background netral — gambar '{q}' tak ditemukan)")
            graded = film_grade(crop_canvas(im))
            # terapkan GAYA ke foto stok
            if eff_style == "duotone":
                base = duotone_teal(graded); theme = "dark"
            elif eff_style == "card":
                base = make_card_bg(graded); theme = "card"      # foto atas + kartu putih bawah
            elif eff_style in ("doodle", "doodle_top"):
                base = graded; theme = pick_theme(graded, forced_theme)  # doodle idealnya pakai [img]
            else:
                base = graded; theme = forced_theme or "dark"

        # posisi teks (zone) sesuai gaya
        if eff_style == "card" and not is_custom:
            zone = "card"
        elif eff_style == "doodle_top":
            zone = "top"
        elif i == 1:
            zone = "bottom"                 # cover default: teks bawah
        else:
            zone = analyze_zone(base)

        clean_path = os.path.join(workdir, f"clean_{i}.png")
        base.save(clean_path, "PNG")                        # foto bersih utk .svg (Figma)
        img_path = os.path.join(workdir, f"slide_{i}.png")
        if theme == "dark":
            bake_scrim(base, zone, img_path)                # overlay gelap (gaya gelap/duotone)
        else:                                               # card / light -> tanpa overlay gelap
            base.save(img_path, "PNG")
        preview_paths.append(img_path)
        slides_data.append({"kind": "normal", "headline": headline, "body": body,
                            "image_path": img_path, "clean_path": clean_path,
                            "index": i, "total": total, "zone": zone,
                            "theme": theme, "style": eff_style})

        # --- #2 CEK MANDIRI (AI vision) — CUKUP slide COVER biar hemat kuota gratis Gemini ---
        if i == 1:
            chk = ai_vision_check(img_path, zone, theme)
            if chk is None:
                print("    [vision cover] dilewati/gagal (kuota/AI nggak jawab)")
            elif not chk["ok"]:
                vision_notes.append(f"Cover: ⚠️ {chk['issue'] or 'teks mungkin kurang terbaca'}")
                print(f"    [vision cover] ⚠️ {chk['issue']}")
            else:
                vision_notes.append("Cover: ✅ aman")
                print("    [vision cover] aman")

    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    logo_path = download_logo_path()
    # logo versi gelap dipakai di slide background terang (kalau disetel & memang ada slide terang)
    logo_dark_path = None
    if LOGO_DARK_URL and any(s.get("theme") == "light" for s in slides_data):
        logo_dark_path = download_logo_path(LOGO_DARK_URL)

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
    if VISION_CHECK_ENABLED:
        warn = [n for n in vision_notes if "⚠️" in n]
        if warn:
            lines += ["", "🔎 Cek mandiri (AI vision):", *warn,
                      "   (kalau mau diperbaiki, pencet Reject atau ketik revisimu)"]
        elif vision_notes:
            lines += ["", "🔎 Cek mandiri (cover): aman ✅"]
        else:
            lines += ["", "🔎 Cek mandiri: dilewati — kuota AI lagi penuh, nanti dicoba lagi"]
    tg_message("\n".join(lines))

    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")

    tg_message(f"📎 ====================\nFILE DESAIN: {title}\n====================")

    if OUTPUT_PPTX:
        pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
        build_pptx(slides_data, pptx_path, logo_path, logo_dark_path)
        if not tg_document(pptx_path, caption=f"{title} — .pptx (semua slide): import ke Canva ✨"):
            tg_message("⚠️ File .pptx gagal dikirim (cek log).")

    if OUTPUT_SVG:
        svg_path = os.path.join(workdir, f"{safe_name}.svg")
        build_svg(slides_data, svg_path, logo_path, logo_dark_path)
        if not tg_document(svg_path, caption=f"{title} — .svg (semua slide): tarik ke Figma ✨"):
            tg_message("⚠️ File .svg gagal dikirim (cek log).")

    if OUTPUT_SVG_PER_SLIDE and total > 1:
        for s in slides_data:
            sp = os.path.join(workdir, f"{safe_name}_slide{s['index']}.svg")
            build_svg([s], sp, logo_path, logo_dark_path)
            if not tg_document(sp, caption=f"{title} — slide {s['index']}/{total} (.svg per slide)"):
                tg_message(f"⚠️ SVG slide {s['index']} gagal dikirim (cek log).")

    # ---- tombol penilaian (1x pencet langsung selesai) utk sistem belajar ----
    if LEARN_ENABLED and mem is not None:
        rev_n = (revision or {}).get("revision", 0)
        token = secrets.token_hex(4)
        pid_cb = (page_id or "").replace("-", "")     # page_id ditaruh di tombol (anti-hilang)
        mem["pending"][token] = {"title": title, "page_id": page_id,
                                 "ts": int(time.time()), "revision": rev_n}
        rows = [[("✅ Approve (oke!)", f"a:{token}:{pid_cb}")]]
        rr = []
        for label, code in REASON_BUTTONS:
            rr.append(("❌ " + label, f"rs:{token}:{code}:{pid_cb}"))
            if len(rr) == 2:
                rows.append(rr); rr = []
        if rr:
            rows.append(rr)
        head = f"Nilai desain '{title}' 👇"
        if rev_n:
            prev = ", ".join((revision or {}).get("reasons", [])[-4:]) or "-"
            head = (f"🔄 REVISI ke-{rev_n} dari '{title}' 👇\n"
                    f"⚠️ Catatan dari sebelumnya: {prev}")
        tip = ("\n💬 atau ketik aja perintahmu (misal: “judul kekecilan, foto lebih cerah”)"
               if AI_CHAT_ENABLED else "")
        tg_buttons(head + "\n✅ kalau udah oke — atau pencet alasannya kalau masih kurang pas:" + tip, rows)

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

    # 3) AUTO-REVISI: generate ulang konten yg tadi di-reject (sampai di-approve)
    if LEARN_ENABLED and mem.get("revise_queue"):
        for pid in list(mem["revise_queue"].keys()):
            info = mem["revise_queue"].pop(pid)   # pindah dari antrian -> jadi pending lagi
            if info.get("revision", 1) > MAX_REVISI:
                tg_message(f"🛑 '{info.get('title','desain')}' udah direvisi {MAX_REVISI}x tapi belum pas. "
                           f"Mungkin lebih enak kamu kasih contoh manual biar gw niru. Gw stop auto-revisi yg ini dulu.")
                continue
            page = notion_get_page(pid)
            if not page:
                tg_message(f"⚠️ Gagal ambil ulang '{info.get('title','desain')}' dari Notion buat revisi.")
                continue
            wd = tempfile.mkdtemp()
            try:
                tg_message(f"🔄 Merevisi '{info.get('title','desain')}' (revisi ke-{info.get('revision',1)})…")
                process_page(page, wd, mem, revision=info)
            except Exception as e:
                print("  x revisi gagal:", e); traceback.print_exc()
                tg_message(f"⚠️ Revisi '{info.get('title','desain')}' gagal: {type(e).__name__}. Gw coba lagi next.")
                mem["revise_queue"][pid] = info   # balikin ke antrian biar dicoba lagi

    # 4) konten BARU berstatus "Siap Desain"
    pages = notion_find_ready()
    print(f"Ditemukan {len(pages)} konten berstatus '{STATUS_READY}'.")
    if not pages:
        # tidak ada kerjaan -> jangan spam Telegram tiap 10 menit, cukup simpan memori
        print("Tidak ada konten baru. (Feedback tetap diproses.)")
        save_memory(mem)
        return

    tg_message(f"✅ Robot Nikah Institute jalan — ada {len(pages)} konten baru.")
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
