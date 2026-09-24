"""
Slide 10 — BSC Business Related Training: tabel training per region, kotak
highlight merah baris `region`, dan insight otomatis (realisasi vs
nasional).
"""
import re

from ..config import REGION_TRAINING
from ..xml_updater import update_table_rows
from ..chart_data import get_slide10_insight

def _fmt_training_num(v):
    """Format angka tabel training slide 10: bulatkan ke integer terdekat sebagai string; None → "0"."""
    if v is None:
        return "0"
    try:
        return str(int(round(float(v))))
    except (TypeError, ValueError):
        return str(v)


def _fmt_training_pct(v):
    """
    Format persentase YTD tabel training slide 10 (1 desimal, mis. "83.3%").
    Menerima nilai fraksi (0.833) ATAU sudah dalam skala persen (83.3) —
    dibedakan lewat ambang `abs(v) <= 1.5`; None → "0.0%".
    """
    if v is None:
        return "0.0%"
    try:
        pct = float(v) * 100 if abs(float(v)) <= 1.5 else float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{pct:.1f}%"


_SLIDE10_SPECIAL_LABELS = ("HO/Sentralisasi", "National", "Regional")
_SLIDE10_LABEL_LOOKUP = {abbrev: canon for canon, abbrev in REGION_TRAINING.items()}
_SLIDE10_LABEL_LOOKUP.update({canon: canon for canon in REGION_TRAINING})
_SLIDE10_LABEL_LOOKUP.update({lbl: lbl for lbl in _SLIDE10_SPECIAL_LABELS})


def _slide10_row_key(cur_texts):
    """
    Mengenali baris tabel BSC training slide 10 itu milik region/label apa,
    berdasarkan daftar teks `<a:t>` yang sudah diekstrak dari barisnya
    (`cur_texts`, urutan run apa adanya).

    Cara kerja: label biasanya cuma 1 run `<a:t>`, tapi beberapa (mis.
    "HO/" + "Sentralisasi") terpecah jadi 2 run — dan karena label itu
    tampil DUA KALI di baris yang sama (kolom label kiri dan kanan pada
    layout tabel training), jumlah run label per baris disimpulkan dari
    total jumlah `<a:t>` baris: total = 2*n (label muncul 2x) + 6 (kolom
    actual) + 6 (kolom plan) + 1 (YTD). Fungsi mencoba n=1 lalu n=2,
    menggabungkan n run pertama jadi satu string label, dan mencocokkannya
    ke `_SLIDE10_LABEL_LOOKUP` (peta singkatan/nama kanonis region + label
    khusus "HO/Sentralisasi"/"National"/"Regional").

    Parameter: `cur_texts` — list string, diteruskan oleh `update_table_rows`
    (di `xml_updater.py`) sebagai `row_key_fn(cur_texts)`, dan juga dipanggil
    langsung dari `_apply_slide10_rect_highlight` untuk baris yang sama.

    Return: key region kanonis (str) atau `None` jika baris ini bukan baris
    data region (mis. header). Dipakai oleh `update_table_rows` untuk
    memutuskan apakah baris perlu diisi, dan oleh `_apply_slide10_rect_highlight`
    untuk mencocokkan baris dengan posisi highlight.
    """
    for n in (1, 2):
        if len(cur_texts) != 2 * n + 13:
            continue
        label = ''.join(cur_texts[:n])
        if label in _SLIDE10_LABEL_LOOKUP:
            return _SLIDE10_LABEL_LOOKUP[label]
    return None



def _apply_slide10_table(slide_xml, data):
    """
    Mengisi tabel training BSC di slide 10 — tabel ini SAMA untuk output
    semua region (data trainingnya bersifat nasional per baris/label, bukan
    per region yang sedang di-generate), sehingga hanya perlu mengisi
    ulang nilai-nilainya dari `data["s10"]` tanpa bergantung pada parameter
    `region`.

    Cara kerja: mendelegasikan pengisian baris ke `update_table_rows()`
    (`xml_updater.py`), dengan `_slide10_row_key` sebagai `row_key_fn`
    (mengenali baris mana milik label/region apa) dan `row_values_fn`
    lokal yang, untuk baris yang cocok, menyusun list nilai kolom baru:
    `n` sel kosong (placeholder label kolom kiri) + kolom actual (6 nilai,
    diformat `_fmt_training_num`) + `n` sel kosong lagi (placeholder label
    kolom kanan) + kolom plan (6 nilai) + 1 nilai YTD% (`_fmt_training_pct`).
    `n` dihitung ulang dari `len(cur_texts)` per baris (bisa 1 atau 2,
    tergantung label pecah jadi berapa run — lihat `_slide10_row_key`).

    Parameter: `slide_xml` (bytes XML slide 10), `data` (dict `load_all()`
    — hanya key `"s10"` yang dipakai, berisi dict `{label: {"actual": [...],
    "plan": [...], "ytd_ach": ...}}`).

    Return: bytes XML slide 10 yang sudah terisi, dipanggil dari
    `generate_pptx_for_region` (langkah 9), lalu diteruskan ke
    `_apply_slide10_rect_highlight` sebelum `editor.update(...)`.
    """
    s10 = data.get("s10", {})

    def row_values_fn(key, cur_texts):
        row = s10.get(key)
        if row is None:
            return None
        n = (len(cur_texts) - 13) // 2
        actual = [_fmt_training_num(v) for v in row["actual"]]
        plan   = [_fmt_training_num(v) for v in row["plan"]]
        ytd    = _fmt_training_pct(row["ytd_ach"])
        return [None] * n + actual + [None] * n + plan + [ytd]

    return update_table_rows(slide_xml, _slide10_row_key, row_values_fn)


_TR_ROW_RE = re.compile(r'<a:tr h="(\d+)"[^>]*>.*?</a:tr>', re.DOTALL)


def _apply_slide10_rect_highlight(slide_xml, region):
    """
    Memindahkan kotak highlight merah slide 10 ke baris tabel yang cocok
    dengan `region`, dengan teknik yang sama seperti pemindahan kotak-kotak
    di slide 3/14: posisi baris ditentukan MURNI dari struktur tabelnya
    sendiri (tinggi tiap baris + label tiap baris), TANPA ada offset yang
    di-hardcode per region.

    Cara kerja: menyisir semua `<a:tr h="...">` di slide untuk mendapatkan
    tinggi (`heights`) dan label/key (`keys`, lewat `_slide10_row_key`)
    tiap baris tabel, lalu membangun `cum` (tinggi kumulatif tiap batas
    baris — posisi y relatif tiap baris terhadap baris pertama). Posisi y
    target dihitung sebagai selisih posisi kumulatif `region` dikurangi
    posisi kumulatif "Jawa Tengah" (baris acuan template, tempat kotak
    highlight berada di posisi originalnya) — `delta`. Kotak highlight
    (dikenali dari warna "FF0000" di dalam `<p:sp>`) lalu digeser sejauh
    `delta` dari posisi y aslinya.

    Parameter: `slide_xml` (bytes XML slide 10, sudah memuat data tabel
    yang baru dari `_apply_slide10_table`), `region` (region yang sedang
    diproses).

    Return: bytes XML slide 10 dengan posisi kotak highlight yang sudah
    diperbarui (atau `slide_xml` apa adanya jika `region` atau "Jawa
    Tengah" tidak ditemukan di tabel, atau delta-nya 0). Dipanggil dari
    `generate_pptx_for_region` (langkah 9) sebagai langkah terakhir sebelum
    `editor.update(slide10_path, slide10_xml)`.
    """
    text = slide_xml.decode('utf-8')

    heights, keys = [], []
    for m in _TR_ROW_RE.finditer(text):
        heights.append(int(m.group(1)))
        row_texts = re.findall(r'<a:t>([^<]*)</a:t>', m.group(0))
        keys.append(_slide10_row_key(row_texts))

    cum = [0]
    for h in heights:
        cum.append(cum[-1] + h)

    try:
        target_idx = keys.index(region)
        anchor_idx = keys.index("Jawa Tengah")
    except ValueError:
        return slide_xml
    delta = cum[target_idx] - cum[anchor_idx]
    if delta == 0:
        return slide_xml

    sp_blocks = re.findall(r'<p:sp\b.*?</p:sp>', text, re.DOTALL)
    for sp in sp_blocks:
        if 'FF0000' not in sp:
            continue
        m = re.search(r'(<a:off x="\d+" y=")(\d+)(")', sp)
        if not m:
            continue
        old_y = int(m.group(2))
        new_y = old_y + delta
        new_sp = sp[:m.start(2)] + str(new_y) + sp[m.end(2):]
        text = text.replace(sp, new_sp, 1)
        break

    return text.encode('utf-8')


def _apply_slide10_insight(slide_xml, data, region):
    """
    Mengisi 1 placeholder insight slide 10 ({SLIDE10_INSIGHT}) dengan teks
    draf hasil `get_slide10_insight()` di chart_data.py (perbandingan %YTD
    ACH training `region` vs nasional).

    Cara kerja: sama seperti `_apply_slide3_highlights` — exact-match
    `<a:t>{token}</a:t>`, aman karena token ini unik dan cuma muncul sekali
    di slide 10.

    Parameter: `slide_xml` (bytes XML slide 10), `data` (dict load_all()),
    `region`.

    Return: bytes XML slide 10 dengan token sudah diganti (atau bytes asli
    tanpa perubahan kalau `get_slide10_insight` mengembalikan string
    kosong). Dipanggil dari `generate_pptx_for_region` (langkah 9), setelah
    `_apply_slide10_rect_highlight`.
    """
    insight = get_slide10_insight(data, region)
    if not insight:
        return slide_xml
    text = slide_xml.decode('utf-8')
    text = text.replace('<a:t>{SLIDE10_INSIGHT}</a:t>', f'<a:t>{insight}</a:t>', 1)
    return text.encode('utf-8')


# ---------------------------------------------------------------------------
# Slide 14 – Tabel YoY attrition NR / RG / Total + kotak highlight
# ---------------------------------------------------------------------------
