"""
Utilitas bersama yang dipakai lintas modul slide (src/slides/*.py) dan oleh
orkestrator utama (src/pptx_updater.py) — penemuan & pembaruan chart,
penggantian teks generik, pembangunan ulang sel tabel secara posisional,
serta beberapa helper format/style yang kebetulan dipakai lebih dari satu
slide (mis. indikator naik/turun/tetap yang sama bentuknya di slide 15
maupun 18, dan fill highlight yang sama pola regex-nya di slide 14 maupun
16).
"""
import re
import os
import math
from lxml import etree

from .xml_updater import (
    update_chart_xml, update_chart_xml_scatter, update_embedded_workbook,
    replace_text_in_slide, PptxEditor, set_value_axis_autoscale,
)
from .chart_data import SCATTER_CHARTS, get_chart_data

_TC_RE = re.compile(r'<a:tc\b.*?</a:tc>', re.DOTALL)


def _rebuild_row_cells(row_xml, transform):
    """
    Membangun ulang satu baris tabel (`<a:tr>...</a:tr>`) dengan
    mentransformasi tiap sel `<a:tc>` di dalamnya lewat
    `transform(index, cell_xml) -> new_cell_xml`.

    KENAPA berbasis POSISI (match span), BUKAN search-and-replace
    berdasarkan ISI (`row_xml.replace(old_cell, new_cell, 1)`):
    sel-sel bertetangga dalam satu baris sangat sering byte-identical satu
    sama lain — baik karena masih placeholder template yang belum disentuh,
    maupun karena nilai aslinya kebetulan sama dengan nilai default
    placeholder (mis. sama-sama "0"). Kalau pakai `str.replace` berbasis
    isi, kasus itu diam-diam rusak: bila `old_cell == new_cell` (nilai sel
    itu memang tidak berubah), "penggantian" itu jadi no-op yang TIDAK
    mengonsumsi kemunculan (occurrence) yang ditemukan — akibatnya panggilan
    replace untuk sel BERIKUTNYA akan kembali mencocokkan kemunculan
    leftmost yang sama itu juga, sehingga isi sel-sel setelah kecocokan
    kebetulan pertama itu jadi berantakan/tertukar. Dengan merekonstruksi
    lewat span posisi tiap match `_TC_RE` (index karakter awal/akhir tiap
    `<a:tc>` di `row_xml`), setiap sel diproses tepat satu kali di posisinya
    sendiri, terlepas dari isi sel-sel lain.

    Parameter:
        row_xml: string XML satu baris `<a:tr>...</a:tr>`, dioper oleh
                 pemanggil (mis. `_apply_slide14_reason_table`,
                 `_apply_slide15_tables`, `_apply_slide16_branch_table`,
                 `_apply_slide16_reason_table`, `_apply_slide18_worst5_table`)
                 setelah mereka mengekstrak baris itu dari `<a:tbl>` lewat
                 regex `<a:tr\\b.*?</a:tr>`.
        transform: fungsi `(index, cell_xml) -> new_cell_xml` yang disuplai
                   si pemanggil (biasanya closure lokal bernama `transform`)
                   untuk mengisi nilai baru ke sel ke-`index` (0-based sesuai
                   urutan kemunculan `<a:tc>` dalam baris).

    Return: tuple `(new_row_xml, num_cells_found)`. Semua pemanggil memakai
    `num_cells_found` sebagai validasi struktural sebelum menerima hasilnya
    (mis. dicek `!= 4` atau `!= 8` untuk memastikan strukturnya belum berubah
    dari yang diharapkan) — kalau tidak cocok, baris lama dipertahankan dan
    `new_row_xml` dibuang.
    """
    matches = list(_TC_RE.finditer(row_xml))
    pieces = []
    prev_end = 0
    for i, m in enumerate(matches):
        pieces.append(row_xml[prev_end:m.start()])
        pieces.append(transform(i, m.group(0)))
        prev_end = m.end()
    pieces.append(row_xml[prev_end:])
    return ''.join(pieces), len(matches)


# ---------------------------------------------------------------------------
# Penemuan file chart dari dalam zip PPTX
# ---------------------------------------------------------------------------
def _discover_charts(editor):
    """
    Menemukan semua file chart di dalam paket PPTX secara dinamis (bukan
    hardcode daftar nomor chart), dengan menyisir isi zip lewat
    `editor.names()` mencari path yang cocok pola
    `ppt/charts/chartN.xml`. Untuk tiap chart yang ditemukan, file relasinya
    (`ppt/charts/_rels/chartN.xml.rels`) dibaca dan di-parse dengan lxml
    untuk mencari relasi yang targetnya mengandung
    "Microsoft_Excel_Worksheet" — itulah workbook Excel ter-embed yang jadi
    sumber data chart tersebut saat dibuka lewat "Edit Data" di PowerPoint.
    Chart yang tidak punya file rels (mis. sudah dihapus/tidak lengkap)
    dilewati (`except KeyError: continue`).

    Parameter:
        editor: instance `PptxEditor` (dari `xml_updater.py`) yang sudah
                dibuka di `generate_pptx_for_region`; dipakai untuk
                `editor.names()` (daftar semua path di zip) dan
                `editor.read(rels_path)`.

    Return: dict `{chart_num (int): {"chart_path": str, "embed_path": str}}`.
    Dipakai oleh `generate_pptx_for_region` (langkah 1) untuk tahu chart apa
    saja yang ada, lalu diteruskan satu per satu ke `_update_chart()`
    (langkah 2) sebagai `chart_info`.
    """
    result = {}
    for name in editor.names():
        m = re.match(r"ppt/charts/chart(\d+)\.xml$", name)
        if not m:
            continue
        chart_num = int(m.group(1))
        rels_path = f"ppt/charts/_rels/chart{chart_num}.xml.rels"
        try:
            rels_data = editor.read(rels_path)
        except KeyError:
            continue
        root = etree.fromstring(rels_data)
        for rel in root:
            target = rel.get("Target", "")
            if "Microsoft_Excel_Worksheet" in target:
                embed_name = os.path.basename(target)
                embed_path = f"ppt/embeddings/{embed_name}"
                result[chart_num] = {
                    "chart_path": name,
                    "embed_path": embed_path,
                }
    return result


# ---------------------------------------------------------------------------
# Pembaruan chart
# ---------------------------------------------------------------------------
# Nomor chart slide 4 (7=Function, 10-12=histogram LOS/EDU/AGE, 13-14=Span
# of Control) yang di template-nya punya batas atas sumbu nilai (`<c:max>`)
# di-hardcode ke nilai region demo — dilepas ke auto-scale
# (`set_value_axis_autoscale`) di sini supaya bar tidak terpotong/keluar
# bingkai untuk region dengan angka lebih besar dari template. Chart 8, 9,
# 15, 16, 17 di slide yang sama sudah auto-scale dari sononya (tidak ada
# `<c:max>` di template), jadi tidak perlu masuk daftar ini.
_AUTOSCALE_CHART_NUMS = {7, 10, 11, 12, 13, 14}


def _update_chart(editor, chart_num, chart_info, data, region):
    """
    Memperbarui satu chart (nomor `chart_num`) untuk `region`: mengambil
    data barunya lewat `get_chart_data()` (dispatcher di `chart_data.py`
    yang meneruskan ke fungsi `chartN_xxx(data, region)` yang sesuai), lalu
    menuliskannya ke XML chart itu sendiri DAN ke workbook Excel ter-embed
    di baliknya (lihat catatan di bawah kenapa keduanya harus disinkronkan).

    Ada dua jalur tergantung jenis chart:
      - Chart scatter (nomor ada di `SCATTER_CHARTS`): `chart_result` berupa
        baris-baris titik scatter, diteruskan ke `update_chart_xml_scatter`
        (di `xml_updater.py`). Workbook ter-embed TIDAK disentuh untuk chart
        scatter (lihat pengecekan `chart_num not in SCATTER_CHARTS` di
        bawah).
      - Chart kategori/garis biasa: `chart_result` berupa
        `(categories, series_list)`, diteruskan ke `update_chart_xml`
        (di `xml_updater.py`). Kalau `categories` kosong, chart itu
        dilewati (`return`) — biasanya berarti region ini tidak punya baris
        data untuk chart tersebut.

    Parameter:
        editor: instance `PptxEditor` yang dipakai untuk `read`/`update`
                path chart dan path workbook ter-embed.
        chart_num: nomor chart (int), key dari dict yang dikembalikan
                   `_discover_charts()`.
        chart_info: dict `{"chart_path", "embed_path"}` untuk chart ini,
                    hasil dari `_discover_charts()`.
        data: dict besar hasil `load_all()` (lihat `data_loader.py`),
              diteruskan apa adanya ke `get_chart_data`.
        region: nama region yang sedang diproses, diteruskan ke
                `get_chart_data`.

    Fungsi ini tidak mengembalikan nilai; efeknya berupa pemanggilan
    `editor.update(chart_path, ...)` dan `editor.update(embed_path, ...)`,
    yang akan ikut tersimpan saat `editor.save(out_path)` dipanggil di akhir
    `generate_pptx_for_region`. Dipanggil dari `generate_pptx_for_region`
    (langkah 2) untuk setiap entri hasil `_discover_charts`.
    """
    chart_path = chart_info["chart_path"]

    chart_result = get_chart_data(chart_num, data, region)
    if chart_result is None:
        return

    chart_xml = editor.read(chart_path)

    if chart_num in SCATTER_CHARTS:
        new_chart_xml = update_chart_xml_scatter(chart_xml, chart_result)
    else:
        categories, series_list = chart_result
        if not categories:
            return
        new_chart_xml = update_chart_xml(chart_xml, categories, series_list)

    if chart_num in _AUTOSCALE_CHART_NUMS:
        new_chart_xml = set_value_axis_autoscale(new_chart_xml)

    editor.update(chart_path, new_chart_xml)

    # Juga sinkronkan workbook Excel yang ter-embed. PowerPoint akan
    # menyinkronkan ulang chart dari workbook ini setiap kali pengguna
    # mengklik "Edit Data" (atau pada beberapa rekalkulasi), sehingga
    # membiarkannya berisi nilai placeholder template akan membuat chart
    # kembali ke data placeholder setelah refresh semacam itu.
    embed_path = chart_info.get("embed_path")
    if embed_path and chart_num not in SCATTER_CHARTS:
        try:
            embed_bytes = editor.read(embed_path)
            new_embed_bytes = update_embedded_workbook(
                embed_bytes, new_chart_xml, categories, series_list)
            editor.update(embed_path, new_embed_bytes)
        except Exception as e:
            print(f"    [WARN] chart{chart_num}: embedded workbook not updated: {e}")


# ---------------------------------------------------------------------------
# Helper pembaruan teks slide
# ---------------------------------------------------------------------------
def _replace_text_in_slide(slide_xml_bytes, replacements):
    """
    Wrapper tipis di atas `replace_text_in_slide()` (`xml_updater.py`),
    yang melakukan penggantian string `(old, new)` polos pada XML slide
    tanpa parsing. Dipanggil dari `generate_pptx_for_region` (langkah 7)
    untuk setiap slide, dengan `replacements` dari
    `_build_slide_replacements(slide_idx, data, region)`. Mengembalikan
    bytes XML slide yang sudah diperbarui, langsung dioper ke
    `editor.update(slide_path, ...)`.
    """
    return replace_text_in_slide(slide_xml_bytes, replacements)


def _get_slide_paths(editor):
    """
    Mengembalikan daftar path XML slide (`ppt/slides/slideN.xml`) yang ada
    di paket PPTX, terurut numerik berdasarkan N (bukan urutan alfabetis
    string, yang akan salah untuk N >= 10, mis. "slide10" < "slide2" secara
    string). Dipanggil dua kali oleh `generate_pptx_for_region`: sekali di
    awal (`slide_paths_early`, sebelum langkah 3-6) dan sekali lagi setelah
    langkah 6 (`slide_paths`, dipakai langkah 7-9) — keduanya menghasilkan
    daftar yang identik, dipanggil ulang sekadar demi keterbacaan alur kode
    per tahap, bukan karena daftar itu berubah di antaranya.
    """
    paths = [n for n in editor.names() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    return sorted(paths, key=lambda x: int(re.search(r"\d+", x.split("/")[-1]).group()))


# ---------------------------------------------------------------------------
# Posisi highlight kotak merah di slide 3
# ---------------------------------------------------------------------------
#
# Slide 3 punya 6 chart bar horizontal (HC, SSD, COLL, CREDIT, LAR, BISNIS).
# Tiap chart di template punya satu kotak highlight merah yang menandai bar
# region tertentu (di template asli: Jawa Tengah), ditambah satu segitiga
# indikator (▲ hijau naik / ▼ merah turun) dan satu text box YoY%. Karena
# output per-region harus menyorot BAR MILIK REGION ITU (posisinya berbeda-
# beda tergantung urutan bar chart-nya), ketiga shape ini harus dipindah
# secara vertikal (offset y) ke posisi bar region yang benar — bukan
# di-hardcode per region, melainkan dihitung dari peringkat (rank) region
# tersebut di data chart yang sama.

# Nilai rect_y template (region Jawa Tengah) = chOff_y (child-offset y) dari
# tiap grpSp kotak highlight. Dipakai baik sebagai kunci pencarian unik di
# XML slide (untuk menemukan shape mana yang harus digeser) MAUPUN sebagai
# titik jangkar (anchor) rumus perhitungan posisi baru.
def _apply_highlight_tokens(slide_xml, prefix, highlights):
    """
    Helper generik: mengganti token `{prefix_KEY}` (mis.
    "SLIDE6_HIGHLIGHT_A") dengan isi `highlights[KEY]` di `slide_xml`, lewat
    exact-match `<a:t>{token}</a:t>` seperti `_apply_slide3_highlights` —
    dipakai bersama oleh `_apply_slide6_highlights` dan
    `_apply_slide7_highlights` supaya logika penggantiannya tidak
    diduplikasi.

    Parameter:
      slide_xml: bytes XML slide.
      prefix: str, mis. "SLIDE6_HIGHLIGHT" — token penuhnya jadi
        "{prefix_KEY}" untuk tiap KEY di `highlights`.
      highlights: dict {KEY (str): value (str)} hasil get_slideN_highlights().
        Slot dengan value "" dilewati (dibiarkan placeholder aslinya, siap
        diisi manual — sama seperti slide 3). Isi value di-escape dulu
        (&, <, >) sebelum disisipkan — PENTING karena beberapa label data
        mentah (mis. kategori LOS "1<x<5 thn") mengandung karakter "<" yang
        akan merusak struktur XML kalau ditulis apa adanya.

    Return: bytes XML yang sudah diganti.
    """
    text = slide_xml.decode('utf-8')
    for key, value in highlights.items():
        if not value:
            continue
        escaped = value.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        token = f"{{{prefix}_{key}}}"
        text = text.replace(f'<a:t>{token}</a:t>', f'<a:t>{escaped}</a:t>', 1)
    return text.encode('utf-8')


def _fmt_slide15_num(v):
    """
    Format angka bulat untuk tabel slide 15 (avg_hc/out_nr/out_rg/out_total).
    Pembulatan half-up (konvensi Excel: 0,5 selalu naik) lewat
    `floor(v + 0.5)`, BUKAN pembulatan half-to-even bawaan Python
    (`round()`), supaya rata-rata yang persis .5 (mis. dari (19+18)/2 =
    18,5) ditampilkan sama seperti cara sheet sumber menampilkannya
    (18,5 → 19, bukan 18 seperti hasil `round()` Python untuk kasus genap).
    None/nilai tak valid → "0".
    """
    try:
        return str(math.floor(float(v) + 0.5))
    except (TypeError, ValueError):
        return "0"


def _fmt_slide15_pct(v):
    """Format persentase 2 desimal gaya Indonesia dari nilai fraksi; nilai tak valid → "0,00%"."""
    try:
        return f"{float(v) * 100:.2f}%".replace(".", ",")
    except (TypeError, ValueError):
        return "0,00%"


_XFRM_OPEN_RE = re.compile(r'<a:xfrm([^>]*)>')
_PRSTGEOM_RE = re.compile(r'<a:prstGeom prst="[^"]*"')
_DIRECT_FILL_RE = re.compile(r'<a:solidFill><a:srgbClr val="[0-9A-Fa-f]{6}"/></a:solidFill>|<a:grpFill/>')


def _set_slide15_indicator_style(sp_xml, direction, group_flip=False):
    """
    Mengubah TAMPILAN satu shape indikator (bentuk + orientasi + warna)
    sesuai `direction`, TANPA mengubah posisi maupun ukurannya (atribut lain
    di `<a:xfrm>` dan `<a:ext>` dibiarkan) — jadi shape yang di template-nya
    kebetulan sebuah segitiga bisa "berubah" jadi lingkaran (dan sebaliknya)
    hanya dengan mengganti elemen `<a:prstGeom>`-nya, tetap di kotak
    pembatas (bounding box) yang sama persis.

    Parameter `group_flip` (bool, default False): True jika grup pembungkus
    shape ini sendiri punya flipV="1" (lihat
    `_compute_slide15_indicator_map`, kasus nyata: "Group 18" / baris Sales
    Support di slide 15). Kalau True, flip yang di-set ke shape ini
    DIBALIK (XOR) supaya hasil akhirnya tetap terlihat benar meski ada
    flip tambahan dari grup — tanpa ini, panah baris tersebut akan
    tampil terbalik (naik kelihatan turun, dan sebaliknya) walau
    warnanya sendiri sudah benar.

    Makna `direction`:
      - "down"  → attrition TURUN (bagus) → segitiga hijau menghadap bawah
                  (flipV="1" ditambahkan)
      - "up"    → attrition NAIK (buruk) → segitiga merah menghadap atas
                  (flipV dihapus jika ada)
      - lainnya ("flat") → tidak berubah → lingkaran oranye (FFC000)

    Cara kerja: `_XFRM_OPEN_RE` menyisipkan/menghapus atribut `flipV="1"`
    pada tag pembuka `<a:xfrm>` (hanya untuk down/up — flat tidak menyentuh
    flip karena bentuknya lingkaran, tidak punya orientasi atas/bawah).
    `_PRSTGEOM_RE` mengganti nilai `prst` pada `<a:prstGeom>` (bentuk
    geometri shape: "triangle" atau "ellipse"). `_DIRECT_FILL_RE` mengganti
    warna fill shape.

    Sejumlah shape template memakai `<a:grpFill/>` (mewarisi warna fill
    grup induknya) alih-alih `<a:solidFill>` miliknya sendiri; kedua bentuk
    itu SAMA-SAMA dicocokkan oleh `_DIRECT_FILL_RE` (lewat alternasi regex)
    dan SAMA-SAMA diganti dengan `<a:solidFill>` eksplisit, supaya setiap
    indikator akhirnya punya warna sendiri yang independen dari grupnya
    (tidak lagi ikut berubah kalau warna grup berubah).

    Parameter: `sp_xml` (string XML satu `<p:sp>`), `direction`
    ("down"/"up"/lainnya). Dipanggil dari `_apply_slide15_indicators` (per
    indikator baris-fungsi slide 15) dan dari `_apply_slide18_summary` (per
    indikator ringkasan fraud slide 18 — fungsi ini dipakai ulang lintas
    slide karena logika bentuk/warna indikatornya identik).

    Return: string XML shape yang sudah dimodifikasi tampilannya.
    """
    def set_flip(m, want_flip):
        attrs = re.sub(r'\s*flipV="1"', '', m.group(1))
        if want_flip:
            attrs += ' flipV="1"'
        return f'<a:xfrm{attrs}>'

    def fill(color):
        return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'

    if direction == "down":
        want_flip = True
    elif direction == "up":
        want_flip = False
    else:
        want_flip = None

    if group_flip and want_flip is not None:
        # Kompensasi flipV="1" di level grup (lihat
        # `_compute_slide15_indicator_map`) — tanpa ini, baris yang
        # grupnya punya flip sendiri akan tampil terbalik meski warnanya
        # tetap benar.
        want_flip = not want_flip

    if direction == "down":
        sp_xml = _XFRM_OPEN_RE.sub(lambda m: set_flip(m, want_flip), sp_xml, count=1)
        sp_xml = _PRSTGEOM_RE.sub('<a:prstGeom prst="triangle"', sp_xml, count=1)
        sp_xml = _DIRECT_FILL_RE.sub(fill("92D050"), sp_xml, count=1)
    elif direction == "up":
        sp_xml = _XFRM_OPEN_RE.sub(lambda m: set_flip(m, want_flip), sp_xml, count=1)
        sp_xml = _PRSTGEOM_RE.sub('<a:prstGeom prst="triangle"', sp_xml, count=1)
        sp_xml = _DIRECT_FILL_RE.sub(fill("FF0000"), sp_xml, count=1)
    else:
        sp_xml = _PRSTGEOM_RE.sub('<a:prstGeom prst="ellipse"', sp_xml, count=1)
        sp_xml = _DIRECT_FILL_RE.sub(fill("FFC000"), sp_xml, count=1)
    return sp_xml


_TCPR_TAIL_OPTIONAL_RE = re.compile(r'(</a:lnB>)(?:<a:noFill/>|<a:solidFill>.*?</a:solidFill>)?(</a:tcPr>)', re.DOTALL)


