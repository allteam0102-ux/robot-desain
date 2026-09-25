# -*- coding: utf-8 -*-
"""
ROBOT DESAIN KONTEN OTOMATIS
============================
Alur kerja:
1. Cek Notion: cari baris konten yang Status-nya = "Siap Desain".
2. Untuk tiap konten: baca isinya per slide, ambil gambar dari Pexels sesuai konteks.
3. Rakit jadi file PowerPoint (.pptx) -> nanti tinggal di-import ke Canva (gratis, bisa diedit).
4. Kirim hasilnya ke Telegram (preview gambar + file .pptx + teks caption).
5. Ubah Status di Notion jadi "Terkirim" biar nggak diproses dua kali.

Kamu TIDAK perlu mengedit file ini. Semua pengaturan lewat "environment variables"
(dijelaskan di README). Kalau nama kolom Notion kamu beda, cukup ubah lewat env, bukan di sini.
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
from pptx.enum.text import PP_ALIGN

# Buat testing di laptop: kalau ada file .env, muat otomatis.
# Di server (GitHub Actions/Railway) bagian ini diabaikan dengan aman.
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


# ======================================================================
# 1. PENGATURAN (diambil dari environment variables)
# ======================================================================
def env(name, default=None, required=False):
    val = os.environ.get(name, default)
    if required and (val is None or str(val).strip() == ""):
        raise SystemExit(f"[SETUP ERROR] Environment variable '{name}' belum diisi. Cek README.")
    return val

# Kunci-kunci (wajib)
NOTION_TOKEN        = env("NOTION_TOKEN", required=True)
NOTION_DATABASE_ID  = env("NOTION_DATABASE_ID", required=True)
PEXELS_API_KEY      = env("PEXELS_API_KEY", required=True)
TELEGRAM_BOT_TOKEN  = env("TELEGRAM_BOT_TOKEN", required=True)
TELEGRAM_CHAT_ID    = env("TELEGRAM_CHAT_ID", required=True)

# Nama kolom di Notion (kalau punyamu beda, ubah lewat env — default ini sesuai README)
TITLE_PROPERTY   = env("TITLE_PROPERTY", "Judul")
STATUS_PROPERTY  = env("STATUS_PROPERTY", "Status")
FORMAT_PROPERTY  = env("FORMAT_PROPERTY", "Format")
CONTENT_PROPERTY = env("CONTENT_PROPERTY", "Konten")
KEYWORD_PROPERTY = env("KEYWORD_PROPERTY", "Kata Kunci Gambar")

# Nilai status
STATUS_READY = env("STATUS_READY", "Siap Desain")   # pemicu / trigger
STATUS_DONE  = env("STATUS_DONE",  "Terkirim")       # setelah sukses
STATUS_ERROR = env("STATUS_ERROR", "Gagal")          # kalau ada error

# Tipe properti Status: "select" (default) atau "status" (kalau kamu pakai tipe Status bawaan Notion)
STATUS_TYPE = env("STATUS_TYPE", "select").strip().lower()

SLIDE_PX = 1080  # ukuran kotak Instagram (1080x1080)


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
    """Cari semua baris yang Status = STATUS_READY."""
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
    """Ubah Status sebuah baris (dan tulis catatan kalau ada kolom 'Catatan')."""
    url = f"{NOTION_BASE}/pages/{page_id}"
    if STATUS_TYPE == "status":
        props = {STATUS_PROPERTY: {"status": {"name": status_value}}}
    else:
        props = {STATUS_PROPERTY: {"select": {"name": status_value}}}
    props_with_note = dict(props)
    if note:
        props_with_note["Catatan"] = {"rich_text": [{"text": {"content": note[:1900]}}]}
    # Coba update lengkap dengan catatan. Kalau kolom 'Catatan' tidak ada (patch gagal),
    # ulangi TANPA catatan supaya status tetap berubah (biar tidak diproses berulang).
    r = requests.patch(url, headers=NOTION_HEADERS, json={"properties": props_with_note}, timeout=60)
    if r.status_code != 200 and note:
        requests.patch(url, headers=NOTION_HEADERS, json={"properties": props}, timeout=60)

def _plain_text(rich_list):
    return "".join(part.get("plain_text", "") for part in (rich_list or []))

def read_property(props, name, kind):
    """Ambil isi sebuah kolom Notion dengan aman (kalau tidak ada -> string kosong)."""
    p = props.get(name)
    if not p:
        return ""
    if kind == "title":
        return _plain_text(p.get("title", [])).strip()
    if kind == "rich_text":
        return _plain_text(p.get("rich_text", [])).strip()
    if kind == "select":
        sel = p.get("select")
        return (sel or {}).get("name", "").strip()
    if kind == "status":
        sel = p.get("status")
        return (sel or {}).get("name", "").strip()
    return ""


# ======================================================================
# 3. PARSING ISI KONTEN JADI SLIDE
# ======================================================================
def split_slides(content_text):
    """Pisah konten jadi beberapa slide. Pemisah antar slide = baris berisi '---'."""
    raw = (content_text or "").replace("\r\n", "\n").replace("\r", "\n")
    blocks = re.split(r"(?m)^\s*---\s*$", raw)
    return [b.strip() for b in blocks if b.strip()]

def headline_and_body(block):
    """Baris pertama = judul besar (headline), sisanya = isi (body)."""
    lines = block.split("\n")
    idx = next((i for i, l in enumerate(lines) if l.strip()), None)
    if idx is None:
        return "", ""
    headline = lines[idx].strip()
    body = "\n".join(lines[idx + 1:]).strip()
    return headline, body

def keyword_for(index, keyword_blocks, headline, body):
    """Tentukan kata kunci gambar untuk sebuah slide."""
    if index < len(keyword_blocks) and keyword_blocks[index].strip():
        q = keyword_blocks[index].strip()
    elif headline:
        q = headline
    elif body:
        q = body
    else:
        q = "aesthetic minimal background"
    # ambil maksimal 5 kata pertama biar hasil Pexels lebih relevan
    words = re.sub(r"[^\w\s]", " ", q).split()
    return " ".join(words[:5]) if words else "aesthetic minimal background"


# ======================================================================
# 4. AMBIL & OLAH GAMBAR PEXELS
# ======================================================================
def pexels_pick(query):
    """Cari gambar di Pexels, kembalikan (bytes_gambar, kredit_teks)."""
    def _search(q):
        r = requests.get(
            "https://api.pexels.com/v1/search",
            headers={"Authorization": PEXELS_API_KEY},
            params={"query": q, "per_page": 8, "orientation": "square"},
            timeout=60,
        )
        r.raise_for_status()
        return r.json().get("photos", [])

    photos = _search(query)
    if not photos:  # cadangan kalau kata kuncinya terlalu spesifik
        photos = _search("calm minimal aesthetic")
    if not photos:
        return None, None

    photo = random.choice(photos[:5])  # variasi biar tidak monoton
    img_url = photo["src"].get("large2x") or photo["src"].get("large") or photo["src"]["original"]
    img_bytes = requests.get(img_url, timeout=60).content
    credit = f'Foto: {photo.get("photographer", "-")} (Pexels)'
    return img_bytes, credit

def prepare_image(img_bytes, out_path):
    """Potong jadi kotak 1080x1080 + kasih gradasi gelap di bawah biar teks terbaca."""
    im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    w, h = im.size
    m = min(w, h)
    left, top = (w - m) // 2, (h - m) // 2
    im = im.crop((left, top, left + m, top + m)).resize((SLIDE_PX, SLIDE_PX), Image.LANCZOS)

    overlay = Image.new("RGBA", (SLIDE_PX, SLIDE_PX), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    start = int(SLIDE_PX * 0.32)
    for y in range(start, SLIDE_PX):
        a = int(215 * (y - start) / (SLIDE_PX - start))
        draw.line([(0, y), (SLIDE_PX, y)], fill=(0, 0, 0, a))
    im = Image.alpha_composite(im.convert("RGBA"), overlay).convert("RGB")
    im.save(out_path, "PNG")
    return out_path


# ======================================================================
# 5. RAKIT FILE .PPTX
# ======================================================================
def frac(f):
    """Konversi pecahan (0-1) ke satuan EMU untuk kanvas 1080px."""
    return Emu(int(f * 10287000))  # 1080px ~ 11.25 inci ~ 10.287.000 EMU

def add_textbox(slide, text, top_frac, height_frac, size_pt, bold, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(frac(0.07), top_frac, frac(0.86), height_frac)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in text.split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        run = p.add_run()
        run.text = line
        f = run.font
        f.size = Pt(size_pt)
        f.bold = bold
        f.name = "Poppins"
        f.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return box

def build_pptx(slides_data, out_path):
    """slides_data = list of dict {headline, body, image_path, index, total}."""
    prs = Presentation()
    prs.slide_width = frac(1.0)
    prs.slide_height = frac(1.0)
    blank = prs.slide_layouts[6]

    for s in slides_data:
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(s["image_path"], 0, 0, width=prs.slide_width, height=prs.slide_height)
        if s["total"] > 1:  # nomor slide untuk carousel
            add_textbox(slide, f'{s["index"]}/{s["total"]}', frac(0.05), frac(0.08), 18, True, PP_ALIGN.RIGHT)
        if s["headline"]:
            add_textbox(slide, s["headline"], frac(0.50), frac(0.20), 38, True)
        if s["body"]:
            add_textbox(slide, s["body"], frac(0.70), frac(0.25), 22, False)

    prs.save(out_path)
    return out_path


# ======================================================================
# 6. KIRIM KE TELEGRAM
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
        slide_blocks = slide_blocks[:1]  # single post = 1 slide saja

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
        slides_data.append({
            "headline": headline, "body": body,
            "image_path": img_path, "index": i, "total": total,
        })

    # rakit pptx
    safe_name = re.sub(r"[^\w\- ]", "", title).strip().replace(" ", "_")[:40] or "desain"
    pptx_path = os.path.join(workdir, f"{safe_name}.pptx")
    build_pptx(slides_data, pptx_path)

    # ringkasan caption untuk kamu copy
    caption_lines = [f"🎨 {title}  ({fmt}, {total} slide)", ""]
    for i, s in enumerate(slides_data, start=1):
        h = s["headline"] or "(tanpa judul)"
        b = (s["body"][:160] + "…") if len(s["body"]) > 160 else s["body"]
        caption_lines.append(f"— Slide {i}: {h}")
        if b:
            caption_lines.append(f"  {b}")
    caption_lines += ["", *credits, "", "➡️ Cara pakai: buka Canva → Upload → pilih file .pptx ini → tiap slide otomatis jadi halaman yang bisa diedit."]
    tg_message("\n".join(caption_lines))

    # kirim preview gambar (biar bisa dilihat langsung di HP)
    for i, p in enumerate(preview_paths, start=1):
        tg_photo(p, caption=f"Preview slide {i}/{total}")

    # kirim file .pptx (yang di-import ke Canva)
    tg_document(pptx_path, caption=f"{title} — import file ini ke Canva untuk mengedit ✨")

    return total


# ======================================================================
# 8. MAIN (jalan sekali, lalu selesai)
# ======================================================================
def main():
    print("== Robot Desain Konten: mulai cek Notion ==")
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
            print("  ✔ Sukses & status di Notion diubah jadi:", STATUS_DONE)
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
