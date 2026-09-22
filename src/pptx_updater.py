"""
Mengorkestrasi seluruh pembaruan PPTX per region:
 - Workbook Excel yang ter-embed (data chart)
 - Cache XML chart
 - Teks slide (judul, kotak anotasi)
 - Tabel native (`<a:tbl>`) dan shape highlight (rect/segitiga/indikator)

Modul ini adalah "otak" orkestrasi: setiap fungsi `_apply_slideNN_*` /
`_compute_slideNN_*` di sini melakukan satu potongan pekerjaan (mengisi satu
tabel, memindahkan satu shape, dsb) dengan cara memanipulasi string XML
slide secara langsung (regex/string replace), BUKAN dengan parse-modify-
serialize penuh lewat lxml (kecuali di beberapa tempat yang memang butuh
navigasi struktural, mis. `_compute_slide15_indicator_map`). Pendekatan ini
dipilih karena python-pptx/lxml serialization ulang akan mengubah urutan
atribut/whitespace di seluruh file dan berisiko merusak referensi relasi
OOXML lain; manipulasi string yang presisi (dengan regex non-greedy `.*?`
dan `re.DOTALL`) menjaga byte-byte lain di file tetap identik dengan
template asli.

Titik masuk (entry point) modul ini adalah `generate_pptx_for_region()` di
bagian paling bawah file, dipanggil sekali per region dari `main.py`.
"""
import re
import math
from lxml import etree

from .config import (
    TEMPLATE_PATH, OUTPUT_DIR, CHART_EMBED_MAP,
    TEMPLATE_REGION, TEMPLATE_ABBREV, TEMPLATE_TOTAL_HC, REGION_ABBREV,
    REGION_TRAINING, REGIONS,
)
from .xml_updater import (
    update_chart_xml, update_chart_xml_scatter, update_embedded_workbook,
    replace_text_in_slide, update_table_rows, PptxEditor,
)
from .chart_data import (
    CHART_DATA_FN, SCATTER_CHARTS, get_chart_data,
    get_slide3_annotations, get_slide3_highlights, get_slide4_total_hc, get_slide4_wc_pct,
    chart1_hc, chart2_ssd, chart3_coll, chart4_credit, chart5_lar, chart6_bisnis,
)

import os

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

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
_SLIDE3_TEMPLATE_RECT_Y = {
    1: 3015959,   # HC
    2: 1742492,   # SSD
    3: 3727074,   # COLL
    4: 6035200,   # CREDIT
    5: 1467363,   # LAR
    6: 3310999,   # BISNIS
}

# Shape segitiga indikator (di dalam grpSp yang sama dengan kotak merah,
# jadi memakai koordinat anak/child-coordinate grpSp itu). Ini adalah shape
# ▲ hijau / ▼ merah yang diposisikan sedikit di bawah bagian atas kotak.
_SLIDE3_TEMPLATE_TRI_Y = {
    1: 3079791,   # HC
    2: 1753500,   # SSD
    3: 3739759,   # COLL
    4: 6056307,   # CREDIT
    5: 1483631,   # LAR
    6: 3327267,   # BISNIS
}

# Shape text box YoY% (berdiri sendiri, pakai koordinat slide — BUKAN di
# dalam grpSp). Isinya label persentase seperti "6,2%" atau "-2,9%".
_SLIDE3_TEMPLATE_TEXT_Y = {
    1: 2229970,   # HC
    2: 1579909,   # SSD
    3: 3435407,   # COLL
    4: 5571577,   # CREDIT
    5: 1964642,   # LAR
    6: 3546078,   # BISNIS
}

# v_k (slot bar dihitung dari atas, 0=teratas) milik Jawa Tengah pada urutan
# ascending tiap chart. Diturunkan dari peringkat Jawa Tengah pada tiap
# metrik: HC idx=9/12, SSD idx=8/12, COLL idx=9/12, CREDIT idx=7/11,
# LAR idx=5/12, BISNIS idx=8/12.
_SLIDE3_TEMPLATE_V_K_JT = {1: 2, 2: 3, 3: 2, 4: 3, 5: 6, 6: 3}

# Tinggi satu slot bar (empiris) dalam satuan koordinat-anak grup, diturunkan
# dari (off_y_rect_group - plot_top_slide) / v_k_jt memakai geometri slide
# yang sebenarnya. Kotak highlight merah berada di dalam grup grpSp;
# atribut y-nya memakai sistem koordinat-anak grup tersebut, bukan koordinat
# slide. Dengan memakai anchor template (v_k_jt, old_y) sebagai basis
# perhitungan, konversi koordinat slide↔anak-grup bisa dihindari sepenuhnya
# — cukup mengalikan selisih slot (v_k - v_k_jt) dengan tinggi-per-slot ini.
_SLIDE3_EMPIRICAL_BAR_STEP = {
    1: 414047,   # HC  (standalone chart, slide_cy=5428331)
    2: 131363,   # SSD
    3: 132816,   # COLL
    4: 143480,   # CREDIT (N=11, Jabotabek 2 excluded)
    5: 130918,   # LAR
    6: 130687,   # BISNIS
}

_SLIDE3_CHART_FNS = {
    1: chart1_hc, 2: chart2_ssd, 3: chart3_coll,
    4: chart4_credit, 5: chart5_lar, 6: chart6_bisnis,
}


def _compute_slide3_rect_positions(data, region):
    """
    Menghitung, untuk masing-masing dari 6 kotak highlight merah di slide 3,
    koordinat y baru (dalam koordinat-anak grup) yang menempatkan kotak itu
    tepat di atas bar milik `region` pada chart yang bersangkutan.

    Cara kerja: untuk tiap chart (lihat `_SLIDE3_CHART_FNS`), kategori dan
    posisi `region` di dalamnya diambil lewat fungsi chart_data yang sama
    yang dipakai untuk mengisi chart itu sendiri (mis. `chart1_hc(data)`),
    sehingga urutan bar yang dipakai untuk menentukan posisi PASTI konsisten
    dengan urutan bar yang benar-benar dirender di chart. Peringkat/slot bar
    `region` dari atas (`v_k`, 0=teratas) dibandingkan dengan slot Jawa
    Tengah di template (`v_k_jt`), lalu:

        new_y = old_y + (v_k_region - v_k_jt) * bar_step

    dengan `old_y` = `_SLIDE3_TEMPLATE_RECT_Y[chart_num]` (chOff_y anchor
    template) dan `bar_step` = `_SLIDE3_EMPIRICAL_BAR_STEP[chart_num]`.

    Parameter:
        data: dict hasil `load_all()`; diteruskan ke tiap fungsi chart di
              `_SLIDE3_CHART_FNS` (mis. `chart1_hc(data)`) untuk
              mendapatkan `(cats, _)` — daftar kategori/region terurut.
        region: nama region yang sedang diproses; posisinya di `cats`
                dicari lewat `cats.index(region)`.

    Return: list `(old_y, new_y)` — satu entri per chart di mana `region`
    ditemukan (chart yang tidak memuat `region` dilewati). Dipanggil dari
    `generate_pptx_for_region` (langkah 8) dan hasilnya diteruskan ke
    `_apply_slide3_rect_updates()` sebagai `rect_updates`.
    """
    results = []
    for chart_num, fn in _SLIDE3_CHART_FNS.items():
        cats, _ = fn(data)
        N = len(cats)
        if N == 0:
            continue
        try:
            data_idx = cats.index(region)   # 0 = nilai terkecil = bar paling bawah
        except ValueError:
            continue

        v_k = N - 1 - data_idx                              # 0 = slot teratas
        v_k_jt = _SLIDE3_TEMPLATE_V_K_JT[chart_num]
        bar_step = _SLIDE3_EMPIRICAL_BAR_STEP[chart_num]
        old_y = _SLIDE3_TEMPLATE_RECT_Y[chart_num]
        new_y = int(round(old_y + (v_k - v_k_jt) * bar_step))
        results.append((old_y, new_y))

    return results


def _apply_slide3_rect_updates(slide_xml_bytes, rect_updates):
    """
    Menerapkan pemindahan posisi y yang sudah dihitung `_compute_slide3_rect_positions`
    ke XML slide 3 yang sesungguhnya.

    Cara kerja: memecah XML slide jadi blok-blok `<p:sp>...</p:sp>` lewat
    regex, lalu hanya menyentuh blok yang mengandung warna "FF0000" (merah —
    ciri khas shape kotak highlight, bukan shape lain). Untuk tiap kandidat,
    dicocokkan `y="{old_y}"` terhadap tiap pasangan di `rect_updates`; kalau
    cocok, atribut y itu diganti dengan nilai barunya dan pencarian berhenti
    (`break`) supaya satu shape tidak tertimpa dua kali oleh pasangan lain
    yang kebetulan juga cocok.

    Parameter:
        slide_xml_bytes: bytes XML slide 3, dibaca lewat
                         `editor.read(slide3_path)` di
                         `generate_pptx_for_region`.
        rect_updates: list `(old_y, new_y)` dari
                      `_compute_slide3_rect_positions(data, region)`.

    Return: bytes XML slide 3 yang sudah diperbarui, diteruskan lagi (masih
    di dalam `generate_pptx_for_region`, langkah 8) ke
    `_apply_slide3_indicator_updates` sebelum akhirnya disimpan lewat
    `editor.update(slide3_path, slide3_xml)`.
    """
    text = slide_xml_bytes.decode('utf-8')
    update_map = {old_y: new_y for old_y, new_y in rect_updates}

    SP_RE = re.compile(r'(<p:sp\b.*?</p:sp>)', re.DOTALL)

    def repl(m):
        sp = m.group(1)
        if 'FF0000' not in sp:
            return sp
        for old_y, new_y in update_map.items():
            y_attr = f'y="{old_y}"'
            if y_attr in sp:
                sp = sp.replace(y_attr, f'y="{new_y}"', 1)
                break
        return sp

    return SP_RE.sub(repl, text).encode('utf-8')


def _compute_slide3_indicator_updates(data, region):
    """
    Untuk tiap chart di slide 3, menghitung pembaruan segitiga indikator +
    text box YoY% milik `region`.

    Cara kerja: mengambil ulang `(cats, series_list)` dari fungsi chart yang
    sama (`_SLIDE3_CHART_FNS`), mencari posisi `region`, lalu menghitung
    `yoy_pct = (v26 - v25) / v25 * 100` dari dua series pertama (asumsi
    series_list[0] = tahun berjalan/v26, series_list[1] = tahun lalu/v25 —
    urutan ini ditentukan oleh fungsi chart_data yang bersangkutan). Offset
    posisi `delta` dihitung dengan RUMUS YANG SAMA PERSIS seperti
    `_compute_slide3_rect_positions` (selisih slot bar × bar_step), karena
    segitiga dan text box itu bergerak mengikuti kotak highlight yang sama.

    Parameter:
        data: dict hasil `load_all()`, diteruskan ke fungsi chart di
              `_SLIDE3_CHART_FNS`.
        region: nama region yang sedang diproses.

    Return: list tuple `(chart_num, delta, pct_str, is_increase)`:
      delta       – offset y yang akan diterapkan (satuan koordinat-anak
                    grup yang sama dengan bar_step)
      pct_str     – persentase yang sudah diformat gaya Indonesia, mis.
                    "6,2%" atau "-2,9%"
      is_increase – True → segitiga ▲ hijau (naik), False → ▼ merah (turun)
    Dipanggil dari `generate_pptx_for_region` (langkah 8); hasilnya
    diteruskan ke `_apply_slide3_indicator_updates`.
    """
    results = []
    for chart_num, fn in _SLIDE3_CHART_FNS.items():
        cats, series_list = fn(data)
        N = len(cats)
        if N == 0:
            continue
        try:
            data_idx = cats.index(region)
        except ValueError:
            continue

        v26 = series_list[0][1][data_idx]
        v25 = series_list[1][1][data_idx]
        if not v25:
            continue
        yoy_pct = (v26 - v25) / v25 * 100

        v_k = N - 1 - data_idx
        v_k_jt = _SLIDE3_TEMPLATE_V_K_JT[chart_num]
        bar_step = _SLIDE3_EMPIRICAL_BAR_STEP[chart_num]
        delta = int(round((v_k - v_k_jt) * bar_step))

        pct_str = f"{yoy_pct:.1f}".replace(".", ",") + "%"
        results.append((chart_num, delta, pct_str, yoy_pct >= 0))

    return results


def _apply_slide3_indicator_updates(slide_xml_bytes, indicator_updates):
    """
    Menerapkan hasil `_compute_slide3_indicator_updates` ke XML slide 3:
    untuk tiap chart, memindahkan DAN mengubah tampilan segitiga indikator,
    serta memindahkan DAN mengganti teks text box YoY%.
      - Segitiga (di dalam grpSp, koordinat-anak grup): pindahkan y, atur
        warna + arah:
          ▲ hijau (naik): rot="10800000" flipV="1"  warna=92D050
          ▼ merah (turun): rot="10800000" saja (tanpa flipV) warna=FF0000
        (rotasi 180° dasarnya membuat segitiga menghadap bawah; flipV="1"
        membalikkannya lagi sehingga net effect-nya menghadap atas — kedua
        atribut itu "saling meniadakan" untuk kasus naik, sesuai orientasi
        default template.)
      - Text box YoY% (shape berdiri sendiri, koordinat slide): pindahkan y,
        ganti isi teksnya dengan `pct_str`.

    Semua grup grpSp kotak highlight punya scale_y=1.0, sehingga `delta`
    yang sama bisa langsung dipakai baik untuk y segitiga (koordinat-anak
    grup) maupun y text box (koordinat slide) tanpa perlu konversi skala.

    Parameter:
        slide_xml_bytes: bytes XML slide 3, hasil dari
                         `_apply_slide3_rect_updates` (sudah memuat posisi
                         kotak highlight yang baru).
        indicator_updates: list dari `_compute_slide3_indicator_updates(data, region)`.

    Return: bytes XML slide 3 final untuk bagian slide-3, disimpan lewat
    `editor.update(slide3_path, slide3_xml)` di `generate_pptx_for_region`
    (langkah 8).
    """
    text = slide_xml_bytes.decode('utf-8')

    # Membangun tabel lookup berdasarkan nilai y template lama
    tri_map  = {}   # old_tri_y  → (new_tri_y, is_increase)
    text_map = {}   # old_text_y → (new_text_y, pct_str)
    for chart_num, delta, pct_str, is_increase in indicator_updates:
        tri_map[_SLIDE3_TEMPLATE_TRI_Y[chart_num]]  = (
            _SLIDE3_TEMPLATE_TRI_Y[chart_num]  + delta, is_increase)
        text_map[_SLIDE3_TEMPLATE_TEXT_Y[chart_num]] = (
            _SLIDE3_TEMPLATE_TEXT_Y[chart_num] + delta, pct_str)

    SP_RE = re.compile(r'(<p:sp\b.*?</p:sp>)', re.DOTALL)

    def repl(m):
        sp = m.group(1)

        if 'prst="triangle"' in sp:
            for old_y, (new_y, is_increase) in tri_map.items():
                if f'y="{old_y}"' not in sp:
                    continue
                # Bangun ulang seluruh elemen xfrm: perbarui y dan atur arah secara eksplisit.
                # Naik ▲: rot=10800000 flipV=1  (saling meniadakan → menghadap atas, sesuai template)
                # Turun ▼: rot=10800000 saja     (rotasi 180° → menghadap bawah)
                def make_xfrm_repl(oy, ny, inc):
                    def xfrm_repl(xm):
                        inner = xm.group(1).replace(f'y="{oy}"', f'y="{ny}"', 1)
                        if inc:
                            return f'<a:xfrm rot="10800000" flipV="1">{inner}</a:xfrm>'
                        else:
                            return f'<a:xfrm rot="10800000">{inner}</a:xfrm>'
                    return xfrm_repl
                sp = re.sub(r'<a:xfrm\b[^>]*>(.*?)</a:xfrm>',
                            make_xfrm_repl(old_y, new_y, is_increase),
                            sp, flags=re.DOTALL)
                # Warna
                if is_increase:
                    sp = sp.replace('<a:srgbClr val="FF0000"/>', '<a:srgbClr val="92D050"/>')
                else:
                    sp = sp.replace('<a:srgbClr val="92D050"/>', '<a:srgbClr val="FF0000"/>')
                break

        elif 'txBox="1"' in sp:
            for old_y, (new_y, pct_str) in text_map.items():
                if f'y="{old_y}"' not in sp:
                    continue
                sp = sp.replace(f'y="{old_y}"', f'y="{new_y}"', 1)
                sp = re.sub(r'(<a:t>)[^<]*(</a:t>)', rf'\g<1>{pct_str}\g<2>', sp, count=1)
                break

        return sp

    return SP_RE.sub(repl, text).encode('utf-8')


def _apply_slide3_highlights(slide_xml, data, region):
    """
    Mengisi 4 placeholder highlight/insight slide 3 ({SLIDE3_HIGHLIGHT_A}
    s.d. {SLIDE3_HIGHLIGHT_D}, semuanya berada dalam satu shape "Rectangle:
    Rounded Corners 57") dengan teks draf hasil `get_slide3_highlights()`
    di chart_data.py.

    Cara kerja: tiap token diganti lewat exact-match `<a:t>{token}</a:t>`
    (bukan pencarian isi teks bebas), karena keempat token ini unik dan
    masing-masing hanya muncul sekali di seluruh slide — jadi replace
    literal count=1 sudah cukup dan aman, tidak perlu mekanisme
    `_rebuild_row_cells` seperti pada sel tabel (yang isinya sering
    duplikat/tidak unik).

    Slot yang isinya "" (string kosong) dari `get_slide3_highlights` —
    saat ini slot B, sengaja dikosongkan karena butuh narasi bisnis yang
    tidak ada di data — akan membuat baris itu tampil kosong di kotak
    highlight, siap diisi manual oleh pengguna di PowerPoint.

    Parameter: `slide_xml` (bytes XML slide 3 — hasil langkah rect/
    indicator sebelumnya dalam `generate_pptx_for_region`, meski urutan
    tidak kritikal di sini karena token highlight independen dari posisi
    kotak/segitiga), `data` (dict `load_all()`), `region`.

    Return: bytes XML slide 3 dengan keempat token highlight sudah diganti,
    disimpan lewat `editor.update(slide3_path, slide3_xml)` di
    `generate_pptx_for_region` (langkah 8).
    """
    highlights = get_slide3_highlights(data, region)
    text = slide_xml.decode('utf-8')
    for key, value in highlights.items():
        token = f"{{SLIDE3_HIGHLIGHT_{key}}}"
        text = text.replace(f'<a:t>{token}</a:t>', f'<a:t>{value}</a:t>', 1)
    return text.encode('utf-8')


# ---------------------------------------------------------------------------
# Anotasi teks per-slide
# ---------------------------------------------------------------------------
def _slide3_text_replacements(data, region, template_region):
    """
    CATATAN: tidak ada pemanggil fungsi ini di manapun dalam codebase (sudah
    dicek dengan grep di seluruh `main.py` dan `src/*.py`) — sepertinya
    sudah digantikan oleh `_build_slide_replacements()` yang dipakai
    langsung di `generate_pptx_for_region` (langkah 7). Tampaknya sisa kode
    dari iterasi sebelumnya yang belum dibersihkan.

    Semula dimaksudkan untuk menghasilkan daftar `(old, new)` replacement
    teks khusus slide 3, memakai `get_slide3_annotations(data, region)`
    (dari `chart_data.py`) untuk mendapatkan teks anotasi YoY% per metrik —
    namun `ann` yang dihasilkan tidak pernah benar-benar dipakai untuk
    membangun `replacements` (lihat komentar asli: target per text box
    persentase, mis. TextBox30=COLL, 56=BISNIS, 62=SSD, 78=HC, 79=CREDIT,
    80=LAR, tidak bisa dibedakan lewat pass string-replace polos ini, jadi
    sengaja dilewati dan mengandalkan chart XML sebagai sumber nilai
    persentase yang tampil).

    Parameter: `data` (dict `load_all()`), `region` (region yang diproses),
    `template_region` (nama region template, dari config `TEMPLATE_REGION`).
    Return: list `(old, new)` — hanya berisi penggantian nama region dasar.
    """
    ann = get_slide3_annotations(data, region)
    replacements = [
        (template_region, region),
        # Text box % spesifik (TextBox30=COLL, 56=BISNIS, 62=SSD, 78=HC, 79=CREDIT, 80=LAR)
        # Tidak bisa ditarget lewat nama shape pada pass ini, jadi nilai % mentah
        # dilewati di sini dan mengandalkan data yang sudah diperbarui di chart XML
    ]
    return replacements


def _slide4_text_replacements(data, region, template_region):
    """
    CATATAN: tidak ditemukan pemanggil fungsi ini di manapun (dicek via
    grep) — sepertinya sudah digantikan oleh cabang `slide_idx == 3` di
    `_build_slide_replacements()`, yang melakukan hal yang sama persis.

    Menghasilkan daftar `(old, new)` untuk replacement teks slide 4: nama
    region, dan angka total HC (dari `get_slide4_total_hc`, di
    `chart_data.py`) menggantikan nilai placeholder template
    `TEMPLATE_TOTAL_HC` (dari `config.py`). `perm_pct`/`non_perm_pct` dari
    `get_slide4_wc_pct` dihitung tapi tidak dipakai — kemungkinan sisa dari
    versi sebelum text box persentase kontrak kerja ditangani lewat jalur
    lain.

    Parameter: `data`, `region`, `template_region` — sama seperti fungsi
    sejenis lainnya di bagian ini.
    """
    total_hc = get_slide4_total_hc(data, region)
    perm_pct, non_perm_pct = get_slide4_wc_pct(data, region)
    return [
        (template_region, region),
        (TEMPLATE_TOTAL_HC, total_hc),   # anotasi total HC (nilai placeholder template)
    ]


def _generic_region_replacements(region, template_region, template_abbrev):
    """
    CATATAN: tidak ditemukan pemanggil fungsi ini di manapun (dicek via
    grep) — `_build_slide_replacements()` hanya memakai `(template_region,
    region)` secara langsung dan tidak lagi mengganti token abbrev terpisah
    (lihat komentar di `_build_slide_replacements`: template saat ini cuma
    punya satu token "REGION", bukan token singkatan terpisah).

    Semula dimaksudkan sebagai penggantian generik untuk slide-slide yang
    tidak punya penanganan khusus: mengganti nama region penuh, singkatan
    region (`template_abbrev` → `REGION_ABBREV.get(region, region)`), dan
    kemunculan literal "Jateng" tambahan.

    Parameter: `region`, `template_region`, `template_abbrev` — nilai
    template dari `config.py` (`TEMPLATE_REGION`, `TEMPLATE_ABBREV`).
    Return: list `(old, new)`.
    """
    abbrev = REGION_ABBREV.get(region, region)
    return [
        (template_region, region),
        (template_abbrev,  abbrev),
        ("Jateng", abbrev),   # kemunculan tambahan
    ]


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


# ---------------------------------------------------------------------------
# Slide 14 – Tabel YoY attrition NR / RG / Total + kotak highlight
# ---------------------------------------------------------------------------
_SLIDE14_METRIC_KEYS = [
    ("yoy_non_regret_25", "yoy_non_regret_26"),  # tabel 0: NR
    ("yoy_regret_25",     "yoy_regret_26"),      # tabel 1: RG
    ("yoy_total_25",      "yoy_total_26"),       # tabel 2: Total
]


def _fmt_pct_id(v):
    """
    Format persentase gaya Indonesia 2 desimal (koma sebagai pemisah
    desimal), mis. "12,34%". Menerima nilai fraksi (0.1234) ATAU sudah
    dalam skala persen, dibedakan lewat ambang `abs(v) <= 1.5`;
    None → "0,00%".
    """
    if v is None:
        return "0,00%"
    try:
        pct = float(v) * 100 if abs(float(v)) <= 1.5 else float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{pct:.2f}%".replace(".", ",")


def _slide14_label_lookup():
    """
    Membangun peta label→region-kanonis untuk slide 14: tiap region
    (`REGIONS`, dari `config.py`) memetakan ke dirinya sendiri, ditambah
    setiap singkatan region (`REGION_ABBREV`) memetakan balik ke nama
    kanonisnya, plus dua label khusus non-region ("Head Office", "Nasional")
    yang memetakan ke dirinya sendiri. Dipanggil sekali saat modul di-import
    untuk mengisi konstanta `_SLIDE14_LABEL_LOOKUP` di bawah, dipakai oleh
    `_resolve_slide14_table_rows` untuk mengenali label baris tabel apa pun
    bentuknya (nama penuh atau singkatan).
    """
    lookup = {r: r for r in REGIONS}
    lookup.update({abbrev: canon for canon, abbrev in REGION_ABBREV.items()})
    lookup["Head Office"] = "Head Office"
    lookup["Nasional"] = "Nasional"
    return lookup


_SLIDE14_LABEL_LOOKUP = _slide14_label_lookup()


def _resolve_slide14_table_rows(table_xml):
    """
    Mengenali region mana yang menjadi pemilik tiap baris pada satu tabel
    ranking di slide 14 (satu dari 3 tabel: NR/RG/Total).

    MASALAH yang dipecahkan: ketiga tabel di template masing-masing punya
    urutan region sendiri-sendiri yang independen (sudah terurut berdasarkan
    nilainya masing-masing — lihat `load_attrition` di `data_loader.py`),
    dan satu baris labelnya bisa jadi MASIH berupa placeholder literal
    template "{REGION}"/"REGION" (bukan nama/singkatan region yang
    sebenarnya). Karena setiap tabel punya urutan berbeda, tidak bisa
    ditebak begitu saja placeholder itu "pasti region X" hanya dari posisi
    barisnya.

    CARA MENYELESAIKAN — resolusi lewat ELIMINASI: setiap baris yang
    labelnya SUDAH berupa nama/singkatan region yang dikenali
    (`_SLIDE14_LABEL_LOOKUP`) dicatat sebagai `seen`. Setelah semua baris
    diperiksa, region mana pun di `REGIONS` yang TIDAK PERNAH muncul di
    `seen` untuk tabel ini pastilah region yang bersembunyi di balik
    placeholder "REGION" tersebut. Ini hanya bisa disimpulkan dengan pasti
    kalau tepat ada SATU baris placeholder dan SATU region yang hilang
    (`len(placeholder_idx) == 1 and len(missing) == 1`); kalau situasinya
    ambigu (lebih dari satu placeholder atau lebih dari satu kandidat),
    baris itu sengaja dibiarkan tidak teridentifikasi (`None`) daripada
    menebak secara serampangan (menebak salah akan menimpa/menduplikasi
    baris region yang salah).

    Parameter:
        table_xml: string XML satu `<a:tbl>...</a:tbl>` slide 14 (salah
                   satu dari 3 tabel NR/RG/Total, atau — dari pemanggil
                   `_apply_slide14_rect_highlights` — potongan
                   `<p:graphicFrame>` yang membungkus tabel tsb, karena
                   regex `<a:tr\\b.*?</a:tr>` di dalam fungsi ini tetap
                   cocok pada kedua bentuk itu).

    Return: tuple `(row_blocks, resolved)` — `row_blocks` adalah list
    string XML tiap `<a:tr>` (dipakai ulang oleh pemanggil untuk mencari
    posisi baris berdasarkan indeks yang sama), dan `resolved[i]` adalah
    key region kanonis untuk baris ke-i (atau `None` untuk baris yang
    bukan baris data, mis. header, atau yang tidak bisa diselesaikan).
    Dipanggil dari `_apply_slide14_tables` (untuk tahu baris mana perlu
    diisi nilai apa) dan `_apply_slide14_rect_highlights` (untuk tahu
    posisi vertikal baris `region` demi memindahkan kotak highlight).
    """
    row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    resolved = []
    seen = set()
    placeholder_idx = []

    for idx, rb in enumerate(row_blocks):
        texts = re.findall(r'<a:t>([^<]*)</a:t>', rb)
        if len(texts) < 3:
            resolved.append(None)
            continue
        n = len(texts) - 2   # 2 sel terakhir selalu nilai YTD25/YTD26
        label = ''.join(texts[:n])
        canon = _SLIDE14_LABEL_LOOKUP.get(label)
        if canon is not None:
            seen.add(canon)
            resolved.append(canon)
        elif label == "REGION":
            resolved.append("__PLACEHOLDER__")
            placeholder_idx.append(idx)
        else:
            resolved.append(None)

    missing = [r for r in REGIONS if r not in seen]
    if len(placeholder_idx) == 1 and len(missing) == 1:
        resolved[placeholder_idx[0]] = missing[0]
    else:
        for idx in placeholder_idx:
            resolved[idx] = None   # ambigu — biarkan tak terselesaikan daripada menebak

    return row_blocks, resolved


def _apply_slide14_tables(slide_xml, data):
    """
    Mengisi 3 tabel YoY attrition NR / RG / Total di slide 14 — yaitu 3
    blok `<a:tbl>` PERTAMA pada slide (blok ke-4 adalah tabel "Top 3 Reason
    Out", ditangani terpisah oleh `_apply_slide14_reason_table` dan tidak
    disentuh di sini).

    Cara kerja: XML slide dipecah jadi potongan-potongan lewat
    `re.split(r'(<a:tbl>.*?</a:tbl>)', ...)` (grup tertangkap disertakan,
    jadi setiap tabel jadi elemen terpisah dalam `parts`, diselingi teks
    non-tabel). Untuk 3 tabel pertama, `_resolve_slide14_table_rows(part)`
    menentukan region kanonis tiap barisnya, lalu tabel itu diisi lewat
    `update_table_rows()` (`xml_updater.py`) dengan:
      - `row_key_fn`: memakai `counter` closure untuk mengambil elemen
        `resolved[idx]` yang cocok secara BERURUTAN dengan baris yang
        sedang diproses `update_table_rows` (karena `update_table_rows`
        memproses baris satu per satu dengan `re.sub`, urutan pemanggilan
        `row_key_fn` PASTI sama dengan urutan baris fisik di tabel —
        sehingga `counter` yang naik 1 tiap panggilan tetap sinkron dengan
        indeks `resolved`).
      - `row_values_fn`: mengambil baris `attrition[canon]` (dari
        `data["attrition"]`, hasil `load_attrition()` di
        `data_loader.py`), memformat nilai YTD25/YTD26 dengan `_fmt_pct_id`
        memakai key metrik yang sesuai tabel ini (`_SLIDE14_METRIC_KEYS[table_idx]`
        — NR untuk tabel 0, RG untuk tabel 1, Total untuk tabel 2). Kalau
        label baris masih literal placeholder "REGION" (kasus yang
        diresolusi lewat eliminasi tadi), label itu DIGANTI dengan singkatan
        region-nya (`REGION_ABBREV`) — kalau bukan placeholder, label yang
        sudah ada (nama/singkatan region lain) dibiarkan (`None` di posisi
        itu berarti "tidak diubah").

    Parameter: `slide_xml` (bytes XML slide 14, sebelum diisi), `data`
    (dict `load_all()`, dipakai key `"attrition"`).

    Return: bytes XML slide 14 dengan 3 tabel YoY sudah terisi. Dipanggil
    dari `generate_pptx_for_region` (langkah 3), sebelum
    `_apply_slide14_rect_highlights` dan `_apply_slide14_reason_table`
    dijalankan berantai di atas hasilnya.
    """
    attrition = data.get("attrition", {})
    text = slide_xml.decode('utf-8')
    parts = re.split(r'(<a:tbl>.*?</a:tbl>)', text, flags=re.DOTALL)

    table_idx = 0
    for i, part in enumerate(parts):
        if not part.startswith('<a:tbl>'):
            continue
        if table_idx < 3:
            _, resolved = _resolve_slide14_table_rows(part)
            v25_key, v26_key = _SLIDE14_METRIC_KEYS[table_idx]
            counter = [0]

            def row_key_fn(cur_texts, resolved=resolved, counter=counter):
                idx = counter[0]
                counter[0] += 1
                return resolved[idx] if idx < len(resolved) else None

            def row_values_fn(canon, cur_texts, v25_key=v25_key, v26_key=v26_key):
                row = attrition.get(canon)
                if row is None:
                    return None
                n = len(cur_texts) - 2
                label = ''.join(cur_texts[:n])
                v25 = _fmt_pct_id(row.get(v25_key))
                v26 = _fmt_pct_id(row.get(v26_key))
                if label == "REGION":
                    override = REGION_ABBREV.get(canon, canon)
                    labels = [override] + [None] * (n - 1)
                else:
                    labels = [None] * n
                return labels + [v25, v26]

            updated = update_table_rows(part.encode('utf-8'), row_key_fn, row_values_fn)
            parts[i] = updated.decode('utf-8')
        table_idx += 1

    return ''.join(parts).encode('utf-8')


def _apply_slide14_rect_highlights(slide_xml, region):
    """
    Memindahkan masing-masing dari 3 kotak highlight merah slide 14 ke
    baris yang cocok dengan `region` di TABEL yang sesuai — dengan teknik
    posisi-baris yang sama seperti slide 10 (tinggi kumulatif baris), tapi
    di sini ada langkah tambahan: mencocokkan tiap kotak highlight dengan
    TABEL mana yang menjadi induknya, karena ada 3 tabel berdampingan.

    Cara kerja:
      1. Untuk tiap `<p:graphicFrame>` yang memuat `<a:tbl>` (3 tabel
         pertama saja — pengecekan `table_idx >= 3` melewati tabel reason-
         out ke-4), catat offset-x tabel itu (`off_m`, posisi horizontal
         tabel di slide), region tiap barisnya (lewat
         `_resolve_slide14_table_rows(gf)`), dan tinggi kumulatif tiap
         baris (`cum`, dihitung sama seperti slide 10). Semua ini disimpan
         di `table_cols` sebagai `(off_x, keys, cum)`.
      2. Untuk tiap shape `<p:sp>` yang mengandung warna merah "FF0000"
         (kotak highlight), tabel induknya ditentukan dengan mencari entri
         `table_cols` yang offset-x-nya PALING DEKAT dengan offset-x kotak
         itu sendiri (`min(..., key=lambda t: abs(t[0] - rect_x))`) — ini
         cara mengaitkan kotak dengan tabelnya secara struktural/posisional,
         bukan dengan asumsi urutan tabel = urutan kotak.
      3. Setelah tabel induknya ketemu, delta posisi y dihitung sama
         seperti slide 10: selisih tinggi kumulatif antara baris `region`
         dan baris "Jawa Tengah" (baris acuan template) di tabel itu.

    Parameter: `slide_xml` (bytes XML slide 14, hasil dari
    `_apply_slide14_tables` — sehingga baris-baris sudah berisi data
    terbaru saat posisi barisnya dibaca ulang di sini), `region`.

    Return: bytes XML slide 14 dengan posisi ketiga kotak highlight sudah
    diperbarui. Dipanggil dari `generate_pptx_for_region` (langkah 3),
    hasilnya diteruskan ke `_apply_slide14_reason_table`.
    """
    text = slide_xml.decode('utf-8')

    table_cols = []  # (off_x, keys, tinggi_kumulatif_awal_tiap_baris)
    gfs = re.findall(r'<p:graphicFrame>.*?</p:graphicFrame>', text, re.DOTALL)
    table_idx = 0
    for gf in gfs:
        if '<a:tbl>' not in gf:
            continue
        if table_idx >= 3:
            table_idx += 1
            continue
        off_m = re.search(r'<a:off x="(\d+)" y="\d+"/>', gf)
        if not off_m:
            table_idx += 1
            continue
        _, keys = _resolve_slide14_table_rows(gf)
        heights = [int(h) for h in re.findall(r'<a:tr h="(\d+)"', gf)]
        cum = [0]
        for h in heights:
            cum.append(cum[-1] + h)
        table_cols.append((int(off_m.group(1)), keys, cum))
        table_idx += 1

    if not table_cols:
        return slide_xml

    sp_blocks = re.findall(r'<p:sp\b.*?</p:sp>', text, re.DOTALL)
    for sp in sp_blocks:
        if 'FF0000' not in sp:
            continue
        off_m = re.search(r'<a:off x="(\d+)" y="(\d+)"', sp)
        if not off_m:
            continue
        rect_x, old_y = int(off_m.group(1)), int(off_m.group(2))
        _, keys, cum = min(table_cols, key=lambda t: abs(t[0] - rect_x))
        try:
            target_idx = keys.index(region)
        except ValueError:
            continue
        anchor_idx = keys.index("Jawa Tengah") if "Jawa Tengah" in keys else target_idx
        delta = cum[target_idx] - cum[anchor_idx]
        if delta == 0:
            continue
        new_y = old_y + delta
        new_sp = sp.replace(f'y="{old_y}"', f'y="{new_y}"', 1)
        text = text.replace(sp, new_sp, 1)

    return text.encode('utf-8')


_TCPR_TAIL_RE = re.compile(r'(</a:lnB>)(?:<a:noFill/>|<a:solidFill>.*?</a:solidFill>)(</a:tcPr>)', re.DOTALL)


def _fmt_reason_pct(v):
    """Format persentase 2 desimal gaya Indonesia dari nilai fraksi (0.1234 → "12,34%")."""
    return f"{v * 100:.2f}%".replace(".", ",")


def _apply_slide14_reason_table(slide_xml, data, region):
    """
    Mengisi tabel "Top 3 Reason Out Attrition" di slide 14 (blok `<a:tbl>`
    KE-4, setelah 3 tabel NR/RG/Total yang ditangani `_apply_slide14_tables`)
    dengan nilai `region` yang sedang diproses — diambil dari
    `data["attrition"][region]["reason_out_rows"]`, hasil komputasi
    `load_slide14_reason_out(region)` di `data_loader.py` (jadi PER REGION,
    bukan nilai yang sama untuk semua region).

    Selain mengisi nilai, fungsi ini juga menghitung highlight (fill kuning
    "FFFF00") pada baris alasan (dari 6 baris individual, bukan baris
    header/agregat — `leaf_idxs`) yang punya nilai TERTINGGI di masing-
    masing dari 3 kolom nilai (non_regret, regret, total) — meniru
    highlight contoh statis yang sudah ada di template (warna sama), tapi
    di sini dihitung secara dinamis per region lewat `max_idx`.

    Cara kerja detail: tabel ke-4 diambil lewat `tbl_matches[3]`. Validasi
    struktural `len(row_blocks) != len(reason_rows) + 1` memastikan jumlah
    baris tabel di template cocok persis dengan jumlah baris data
    (+1 header) sebelum mengubah apa pun — kalau tidak cocok, tabel
    dibiarkan apa adanya. Untuk tiap baris, `_rebuild_row_cells` dipakai
    (bukan search-and-replace biasa) untuk mengisi ulang label + 3 nilai
    kolom, dan sekaligus menyisipkan/menghapus fill kuning pada `<a:tcPr>`
    sel yang bersangkutan lewat `_TCPR_TAIL_RE` (regex yang menyasar bagian
    ekor `<a:tcPr>`, tepat setelah `</a:lnB>`, tempat elemen fill berada).

    Parameter: `slide_xml` (bytes XML slide 14, hasil rantai
    `_apply_slide14_tables` → `_apply_slide14_rect_highlights`), `data`
    (dict `load_all()`), `region`.

    Return: bytes XML slide 14 final. Dipanggil dari
    `generate_pptx_for_region` (langkah 3), hasilnya langsung disimpan
    lewat `editor.update(slide14_path, slide14_xml)`.
    """
    reason_rows = data.get("attrition", {}).get(region, {}).get("reason_out_rows", [])
    if not reason_rows:
        return slide_xml

    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) < 4:
        return slide_xml
    table_xml = tbl_matches[3].group(0)

    row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    if len(row_blocks) != len(reason_rows) + 1:   # +1 untuk baris header
        return slide_xml

    leaf_idxs = [i for i, r in enumerate(reason_rows) if not r["is_header"]]
    max_idx = {col: max(leaf_idxs, key=lambda i: reason_rows[i][col])
               for col in ("non_regret", "regret", "total")}
    value_cols = [None, "non_regret", "regret", "total"]

    new_table_xml = table_xml
    for ridx, data_row in enumerate(reason_rows):
        row_xml = row_blocks[ridx + 1]

        new_values = [data_row["label"],
                      _fmt_reason_pct(data_row["non_regret"]),
                      _fmt_reason_pct(data_row["regret"]),
                      _fmt_reason_pct(data_row["total"])]

        def transform(ci, tc, new_values=new_values, data_row=data_row, ridx=ridx):
            new_tc = re.sub(r'(<a:t>)[^<]*(</a:t>)',
                             lambda m, v=new_values[ci]: m.group(1) + v.replace('&', '&amp;') + m.group(2),
                             tc, count=1)
            col = value_cols[ci]
            if col is not None and not data_row["is_header"]:
                fill = '<a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>' if max_idx[col] == ridx else '<a:noFill/>'
                new_tc = _TCPR_TAIL_RE.sub(rf'\1{fill}\2', new_tc, count=1)
            return new_tc

        new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
        if num_tcs != 4:
            continue
        new_table_xml = new_table_xml.replace(row_xml, new_row_xml, 1)

    text = text.replace(table_xml, new_table_xml, 1)
    return text.encode('utf-8')


# ---------------------------------------------------------------------------
# Slide 15 – Attrition Report by Function (2 tabel + indikator naik/turun/tetap)
# ---------------------------------------------------------------------------
_SLIDE15_INDICATOR_GROUPS = [
    "Group 2", "Group 11", "Group 18", "Group 23",
    "Group 27", "Group 31", "Group 35", "Group 39", "Group 43",
]
_SLIDE15_PCT_KEYS = ["pct_nr", "pct_rg", "pct_total"]


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


def _apply_slide15_tables(slide_xml, data, region):
    """
    Mengisi kedua tabel "Attrition Report by Function" di slide 15 (Mei-26
    di atas, Mei-25 di bawah) dengan data asli `region`. Urutan baris dan
    labelnya tetap (fixed) di template — jadi fungsi ini HANYA menimpa 7
    sel nilai per baris (label dan header tidak disentuh, ditandai `None`
    di posisi 0 pada list `values`).

    Cara kerja: kedua tabel diambil lewat `tbl_matches[0]` (key "y26") dan
    `tbl_matches[1]` (key "y25") dari `data["s15_func"][region]`. Untuk
    tiap tabel, validasi `len(row_blocks) != len(func_rows) + 1` memastikan
    jumlah baris template cocok dengan jumlah baris data (+1 header)
    sebelum menulis apa pun — kalau tidak cocok tabel itu dilewati
    (`continue`), tabel lainnya tetap diproses. Tiap baris data diisi lewat
    `_rebuild_row_cells` (bukan search-and-replace biasa — lihat alasannya
    di docstring `_rebuild_row_cells`, karena banyak sel yang nilainya
    kebetulan sama, mis. "0"). Setelah tabel pertama diproses, XML slide di
    `text` diperbarui dan `tbl_matches` DIHITUNG ULANG (parsing ulang) agar
    posisi tabel kedua yang dicari tetap valid (posisi karakter string bisa
    bergeser setelah tabel pertama diubah panjangnya).

    Parameter: `slide_xml` (bytes XML slide 15), `data` (dict `load_all()`,
    key `"s15_func"`), `region`.

    Return: bytes XML slide 15 dengan kedua tabel terisi. Dipanggil dari
    `generate_pptx_for_region` (langkah 4), hasilnya diteruskan ke
    `_apply_slide15_indicators` sebelum `editor.update(...)`.
    """
    s15 = data.get("s15_func", {}).get(region, {})
    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) < 2:
        return slide_xml

    for ti, key in enumerate(("y26", "y25")):
        func_rows = s15.get(key, [])
        table_xml = tbl_matches[ti].group(0)
        row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
        if len(row_blocks) != len(func_rows) + 1:   # +1 baris header
            continue

        new_table_xml = table_xml
        for ridx, fr in enumerate(func_rows):
            row_xml = row_blocks[ridx + 1]
            values = [None,
                      _fmt_slide15_num(fr["avg_hc"]),
                      _fmt_slide15_num(fr["out_nr"]),
                      _fmt_slide15_num(fr["out_rg"]),
                      _fmt_slide15_num(fr["out_total"]),
                      _fmt_slide15_pct(fr["pct_nr"]),
                      _fmt_slide15_pct(fr["pct_rg"]),
                      _fmt_slide15_pct(fr["pct_total"])]

            def transform(ci, tc, values=values):
                if values[ci] is None:
                    return tc
                return re.sub(r'(<a:t>)[^<]*(</a:t>)',
                               lambda m, v=values[ci]: m.group(1) + v + m.group(2),
                               tc, count=1)

            new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
            if num_tcs != 8:
                continue
            new_table_xml = new_table_xml.replace(row_xml, new_row_xml, 1)

        text = text.replace(table_xml, new_table_xml, 1)
        tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))

    return text.encode('utf-8')


def _compute_slide15_indicator_map(slide_xml):
    """
    Menemukan, untuk tiap shape indikator naik/turun/tetap di slide 15,
    posisinya sebagai `(row_idx, metric_idx)` — row_idx = baris fungsi ke
    berapa (0-based), metric_idx = kolom metrik ke berapa (0=non-regret,
    1=regret, 2=total, sesuai `_SLIDE15_PCT_KEYS`).

    KENAPA STRUKTURAL, BUKAN HARDCODE: posisi ini diturunkan dari struktur
    XML slide itu sendiri, bukan dari daftar nama shape yang ditulis
    manual, supaya tetap benar meskipun PowerPoint menomori ulang nama
    shape (mis. "Isosceles Triangle 12") ketika file di-resave — nama shape
    semacam itu TIDAK STABIL antar-resave, sedangkan struktur grup (grpSp
    mana yang menaungi shape ini) dan urutan posisi horizontalnya stabil.

    Cara kerja: hanya grpSp yang namanya ada di `_SLIDE15_INDICATOR_GROUPS`
    (nama grup ini SENDIRI juga berupa nama shape PowerPoint, tapi grup
    lebih stabil karena tidak diubah-ubah pengguna/tidak digambar ulang
    seperti triangle-nya) yang diproses; indeksnya di list itu = `row_idx`
    (baris fungsi ke berapa — 1 grup = 1 baris tabel fungsi). Di dalam tiap
    grup, semua shape anak langsung (`./p:sp`) diambil beserta offset-x
    nya, lalu DIURUTKAN berdasarkan offset-x (`shapes.sort()`) — urutan
    kiri-ke-kanan pada slide SELALU sama dengan urutan metrik non-regret →
    regret → total di layout template, sehingga urutan-setelah-sort itulah
    yang menjadi `metric_idx` (0, 1, 2).

    Parameter: `slide_xml` (bytes XML slide 15, di-parse dengan
    `etree.fromstring` — BUKAN string/regex seperti kebanyakan fungsi lain
    di file ini, karena di sini navigasi ancestor/descendant XML yang
    sesungguhnya dibutuhkan, bukan sekadar pencarian pola teks).

    Return: dict `{nama_shape (str): (row_idx, metric_idx)}`. Dipanggil
    dari `_apply_slide15_indicators`, yang lalu mencocokkan nama tiap shape
    `<p:sp>` di slide terhadap dict ini untuk tahu baris/metrik mana yang
    diwakili shape tersebut.
    """
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
          'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    root = etree.fromstring(slide_xml)
    mapping = {}
    for grp in root.iter('{%s}grpSp' % ns['p']):
        cnv = grp.find('./p:nvGrpSpPr/p:cNvPr', ns)
        gname = cnv.get('name') if cnv is not None else None
        if gname not in _SLIDE15_INDICATOR_GROUPS:
            continue
        row_idx = _SLIDE15_INDICATOR_GROUPS.index(gname)
        shapes = []
        for sp in grp.findall('./p:sp', ns):
            name = sp.find('.//p:cNvPr', ns).get('name')
            off = sp.find('.//a:xfrm/a:off', ns)
            shapes.append((int(off.get('x')), name))
        shapes.sort()
        for metric_idx, (_, name) in enumerate(shapes):
            mapping[name] = (row_idx, metric_idx)
    return mapping


_XFRM_OPEN_RE = re.compile(r'<a:xfrm([^>]*)>')
_PRSTGEOM_RE = re.compile(r'<a:prstGeom prst="[^"]*"')
_DIRECT_FILL_RE = re.compile(r'<a:solidFill><a:srgbClr val="[0-9A-Fa-f]{6}"/></a:solidFill>|<a:grpFill/>')


def _set_slide15_indicator_style(sp_xml, direction):
    """
    Mengubah TAMPILAN satu shape indikator (bentuk + orientasi + warna)
    sesuai `direction`, TANPA mengubah posisi maupun ukurannya (atribut lain
    di `<a:xfrm>` dan `<a:ext>` dibiarkan) — jadi shape yang di template-nya
    kebetulan sebuah segitiga bisa "berubah" jadi lingkaran (dan sebaliknya)
    hanya dengan mengganti elemen `<a:prstGeom>`-nya, tetap di kotak
    pembatas (bounding box) yang sama persis.

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
        sp_xml = _XFRM_OPEN_RE.sub(lambda m: set_flip(m, True), sp_xml, count=1)
        sp_xml = _PRSTGEOM_RE.sub('<a:prstGeom prst="triangle"', sp_xml, count=1)
        sp_xml = _DIRECT_FILL_RE.sub(fill("92D050"), sp_xml, count=1)
    elif direction == "up":
        sp_xml = _XFRM_OPEN_RE.sub(lambda m: set_flip(m, False), sp_xml, count=1)
        sp_xml = _PRSTGEOM_RE.sub('<a:prstGeom prst="triangle"', sp_xml, count=1)
        sp_xml = _DIRECT_FILL_RE.sub(fill("FF0000"), sp_xml, count=1)
    else:
        sp_xml = _PRSTGEOM_RE.sub('<a:prstGeom prst="ellipse"', sp_xml, count=1)
        sp_xml = _DIRECT_FILL_RE.sub(fill("FFC000"), sp_xml, count=1)
    return sp_xml


def _apply_slide15_indicators(slide_xml, data, region):
    """
    Menyetel 3 indikator naik/turun/tetap tiap baris fungsi di slide 15,
    dengan membandingkan nilai %Attr tahun berjalan `region` (y26) terhadap
    tahun lalu (y25), per metrik (non-regret / regret / total).

    Cara kerja: `y26`/`y25` diambil dari `data["s15_func"][region]` (list
    dict per baris fungsi, urutannya harus sama antara y26 dan y25 —
    divalidasi lewat `len(y26) != len(y25)`). Posisi tiap shape indikator
    (baris & metrik mana yang diwakilinya) diperoleh dari
    `_compute_slide15_indicator_map(slide_xml)`. Untuk tiap shape `<p:sp>`
    yang namanya cocok dengan salah satu entri di map itu, nilai
    `y26[row_idx][key]` dibandingkan dengan `y25[row_idx][key]`
    (`key` = `_SLIDE15_PCT_KEYS[metric_idx]`) untuk menentukan `direction`
    ("down" jika turun, "up" jika naik, "flat" jika sama), lalu
    `_set_slide15_indicator_style(sp, direction)` menerapkan tampilannya.

    Parameter: `slide_xml` (bytes XML slide 15, hasil dari
    `_apply_slide15_tables` — dipanggil SETELAH tabel diisi, meski secara
    teknis kedua langkah independen karena yang dibaca di sini adalah data
    `data["s15_func"]`, bukan isi tabel XML), `data` (dict `load_all()`),
    `region`.

    Return: bytes XML slide 15 dengan ketiga indikator tiap baris sudah
    disetel. Dipanggil dari `generate_pptx_for_region` (langkah 4), sebagai
    langkah terakhir sebelum `editor.update(slide15_path, slide15_xml)`.
    """
    s15 = data.get("s15_func", {}).get(region, {})
    y26, y25 = s15.get("y26", []), s15.get("y25", [])
    if len(y26) != len(y25):
        return slide_xml

    indicator_map = _compute_slide15_indicator_map(slide_xml)
    if not indicator_map:
        return slide_xml

    text = slide_xml.decode('utf-8')
    sp_blocks = re.findall(r'<p:sp>.*?</p:sp>', text, re.DOTALL)
    for sp in sp_blocks:
        m = re.search(r'name="([^"]*)"', sp)
        if not m or m.group(1) not in indicator_map:
            continue
        row_idx, metric_idx = indicator_map[m.group(1)]
        if row_idx >= len(y26):
            continue
        key = _SLIDE15_PCT_KEYS[metric_idx]
        v26, v25 = y26[row_idx].get(key), y25[row_idx].get(key)
        if v26 is None or v25 is None:
            continue
        direction = "down" if v26 < v25 else ("up" if v26 > v25 else "flat")
        new_sp = _set_slide15_indicator_style(sp, direction)
        if new_sp != sp:
            text = text.replace(sp, new_sp, 1)

    return text.encode('utf-8')


# ---------------------------------------------------------------------------
# Slide 16 – Field Regretted Attrition (tabel per-cabang/cluster + tabel reason-out)
# ---------------------------------------------------------------------------
_TR_OPEN_RE = re.compile(r'<a:tr h="\d+"')


def _apply_slide16_branch_table(slide_xml, data, region, table_index, data_key, grand_total_key):
    """
    Mengisi salah satu dari 2 tabel per-cabang/cluster di slide 16 (dipanggil
    dua kali oleh `generate_pptx_for_region`: sekali untuk Sales/branch,
    sekali untuk Collection/cluster) dengan SEBANYAK baris data yang benar-
    benar dimiliki `region` — data sudah terurut dari %Regret terbesar oleh
    loader (`data_loader.py`), sehingga urutan tampil di sini tinggal ikut
    urutan `branch_rows` apa adanya.

    MASALAH yang dipecahkan: jumlah cabang/cluster tiap region BERBEDA-BEDA,
    sedangkan template punya jumlah baris tabel yang TETAP (dirancang untuk
    region dengan cabang terbanyak). Kalau region ini cabangnya lebih
    sedikit, sisa baris template TIDAK dibiarkan kosong/bernilai 0 (yang
    akan terlihat janggal) — melainkan baris-baris yang tak terpakai
    DIHAPUS SELURUHNYA (`new_table_xml.replace(row_xml, '', 1)`), dan tinggi
    baris-baris yang tersisa (`new_row_height`) DIHITUNG ULANG dengan
    membagi rata total tinggi seluruh baris data yang dihapus/dipertahankan
    (`total_data_height // keep_count`), sehingga tabel tetap mengisi ruang
    vertikal aslinya secara utuh — tidak ada celah kosong di bawah, tidak
    ada baris kosong/bernilai 0.

    Baris grand-total ("REGION – ..." — baris terakhir, tetap, tidak ikut
    dihapus/digeser) hanya nilai-nilainya yang diperbarui, memakai agregat
    yang SAMA dengan fungsi yang cocok di slide 15 (`grand_total_key`, mis.
    "Sales Officer"/"Collection Officer", dicocokkan dengan label baris
    `data["s15_func"][region]["y26"]`) — labelnya sendiri dibiarkan
    (`gt_transform` melewatkan `ci == 0`).

    Parameter:
        slide_xml: bytes XML slide 16.
        data: dict `load_all()` — dipakai `data["s16"][region][data_key]`
              (list baris cabang/cluster) dan `data["s15_func"][region]["y26"]`
              (untuk baris grand-total).
        region: region yang sedang diproses.
        table_index: indeks tabel (0-based) di antara `<a:tbl>` slide 16 —
                     0 untuk tabel Sales, 1 untuk tabel Collection (dioper
                     oleh `generate_pptx_for_region`).
        data_key: key di `data["s16"][region]` untuk baris cabang, mis.
                  "branch_sales"/"branch_collection".
        grand_total_key: label baris yang dicari di agregat slide-15, mis.
                         "Sales Officer"/"Collection Officer".

    Return: bytes XML slide 16 dengan tabel `table_index` sudah terisi
    (baris tak terpakai dihapus, tinggi baris disesuaikan, grand-total
    diperbarui). Dipanggil dua kali dari `generate_pptx_for_region`
    (langkah 5), hasilnya diteruskan ke pemanggilan berikutnya (untuk tabel
    lain) lalu ke `_apply_slide16_reason_table`.
    """
    branch_rows = data.get("s16", {}).get(region, {}).get(data_key, [])
    grand_total = data.get("s15_func", {}).get(region, {}).get("y26", [])
    grand_total = next((r for r in grand_total if r["label"] == grand_total_key), None)

    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) <= table_index:
        return slide_xml
    table_xml = tbl_matches[table_index].group(0)
    row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    if len(row_blocks) < 2:
        return slide_xml

    data_row_count = len(row_blocks) - 2   # dikurangi baris header dan baris grand-total
    keep_count = min(len(branch_rows), data_row_count)

    def cell_values(label, r):
        if r is None:
            return [label, "0", "0", "0", "0", "0,00%", "0,00%", "0,00%"]
        return [label,
                _fmt_slide15_num(r["avg_hc"]), _fmt_slide15_num(r["out_nr"]),
                _fmt_slide15_num(r["out_rg"]), _fmt_slide15_num(r["out_total"]),
                _fmt_slide15_pct(r["pct_nr"]), _fmt_slide15_pct(r["pct_rg"]), _fmt_slide15_pct(r["pct_total"])]

    # Bagikan ulang total tinggi baris-baris yang dihapus ke baris-baris yang
    # dipertahankan, supaya tinggi keseluruhan tabel tidak berubah (tidak ada
    # ruang kosong yang terbuang).
    data_rows = row_blocks[1:1 + data_row_count]
    total_data_height = sum(int(m.group(0)[9:-1]) for m in
                             (re.match(r'<a:tr h="(\d+)"', r) for r in data_rows) if m)
    new_row_height = (total_data_height // keep_count) if keep_count else 0

    new_table_xml = table_xml
    for i in range(data_row_count):
        row_xml = row_blocks[i + 1]
        if i >= keep_count:
            new_table_xml = new_table_xml.replace(row_xml, '', 1)
            continue
        r = branch_rows[i]
        values = cell_values(r["label"], r)

        def transform(ci, tc, values=values):
            return re.sub(r'(<a:t>)[^<]*(</a:t>)',
                           lambda m, v=values[ci]: m.group(1) + v.replace('&', '&amp;') + m.group(2),
                           tc, count=1)

        new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
        if num_tcs != 8:
            continue
        new_row_xml = _TR_OPEN_RE.sub(f'<a:tr h="{new_row_height}"', new_row_xml, count=1)
        new_table_xml = new_table_xml.replace(row_xml, new_row_xml, 1)

    # Baris grand-total: pertahankan label "REGION – ..." yang sudah ada, hanya perbarui nilainya
    gt_row_xml = row_blocks[-1]
    gt_values = cell_values(None, grand_total)   # index 0 (label) sengaja dibiarkan di bawah

    def gt_transform(ci, tc, gt_values=gt_values):
        if ci == 0:
            return tc
        return re.sub(r'(<a:t>)[^<]*(</a:t>)',
                       lambda m, v=gt_values[ci]: m.group(1) + v + m.group(2),
                       tc, count=1)

    new_gt_row, num_gt_tcs = _rebuild_row_cells(gt_row_xml, gt_transform)
    if num_gt_tcs == 8:
        new_table_xml = new_table_xml.replace(gt_row_xml, new_gt_row, 1)

    text = text.replace(table_xml, new_table_xml, 1)
    return text.encode('utf-8')


def _fmt_slide16_reason_num(v):
    """Format angka bulat (tanpa desimal) untuk tabel reason-out slide 16; nilai tak valid → "0"."""
    try:
        return str(int(v))
    except (TypeError, ValueError):
        return "0"


_TCPR_TAIL_OPTIONAL_RE = re.compile(r'(</a:lnB>)(?:<a:noFill/>|<a:solidFill>.*?</a:solidFill>)?(</a:tcPr>)', re.DOTALL)


def _apply_slide16_reason_table(slide_xml, data, region):
    """
    Mengisi tabel reason-out slide 16 (blok `<a:tbl>` KE-3, setelah 2 tabel
    cabang/cluster yang ditangani `_apply_slide16_branch_table`) dengan data
    asli `region`, dan menghitung highlight (fill merah) pada baris alasan
    (dari 5 baris, TIDAK termasuk baris "Total Regret"/header — dikeluarkan
    dari perhitungan max lewat `leaf_idxs = [... if not r["is_header"]]`)
    yang punya nilai tertinggi di masing-masing dari 3 kolom (Sales Officer,
    Collection Officer, Total).

    Cara kerja: validasi struktural `len(row_blocks) != len(reason_rows) + 1`
    memastikan jumlah baris template cocok dengan jumlah baris data (+1
    header) sebelum menulis apa pun. Tiap baris diproses dengan
    `_rebuild_row_cells`; `transform` melewatkan sel label (`ci == 0` —
    dibiarkan apa adanya, walau isinya bisa terdiri >1 run teks, mis.
    "Better " + "Job&Benefit") dan mengisi 3 sel nilai berikutnya, sekaligus
    menyisipkan/menghapus fill merah lewat `_TCPR_TAIL_OPTIONAL_RE` (varian
    `_TCPR_TAIL_RE` yang membuat elemen fill OPSIONAL dalam pola regex-nya,
    karena baris di template ini mungkin belum punya elemen fill sama
    sekali sebelum baris data pertama disorot).

    Parameter: `slide_xml` (bytes XML slide 16, hasil rantai
    `_apply_slide16_branch_table` × 2), `data` (dict `load_all()`, key
    `data["s16"][region]["reason_out"]`), `region`.

    Return: bytes XML slide 16 final. Dipanggil dari
    `generate_pptx_for_region` (langkah 5), hasilnya langsung disimpan
    lewat `editor.update(slide16_path, slide16_xml)`.
    """
    reason_rows = data.get("s16", {}).get(region, {}).get("reason_out", [])
    if not reason_rows:
        return slide_xml

    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) < 3:
        return slide_xml
    table_xml = tbl_matches[2].group(0)
    row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    if len(row_blocks) != len(reason_rows) + 1:
        return slide_xml

    leaf_idxs = [i for i, r in enumerate(reason_rows) if not r["is_header"]]
    max_idx = {col: max(leaf_idxs, key=lambda i: reason_rows[i][col])
               for col in ("sales_officer", "collection_officer", "total")}

    value_cols = [None, "sales_officer", "collection_officer", "total"]

    new_table_xml = table_xml
    for ridx, rr in enumerate(reason_rows):
        row_xml = row_blocks[ridx + 1]

        def transform(ci, tc, rr=rr, ridx=ridx):
            # Selalu tepat 4 sel: 1 sel label (yang bisa saja terdiri dari
            # >1 run teks, mis. "Better " + "Job&Benefit" — tidak disentuh
            # di sini) + 3 sel nilai.
            if ci == 0:
                return tc
            col = value_cols[ci]
            value = _fmt_slide16_reason_num(rr[col])
            new_tc = re.sub(r'(<a:t>)[^<]*(</a:t>)',
                             lambda m, v=value: m.group(1) + v + m.group(2),
                             tc, count=1)
            if not rr["is_header"]:
                fill = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' if max_idx[col] == ridx else ''
                new_tc = _TCPR_TAIL_OPTIONAL_RE.sub(rf'\1{fill}\2', new_tc, count=1)
            return new_tc

        new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
        if num_tcs != 4:
            continue
        new_table_xml = new_table_xml.replace(row_xml, new_row_xml, 1)

    text = text.replace(table_xml, new_table_xml, 1)
    return text.encode('utf-8')


# ---------------------------------------------------------------------------
# Slide 18 – Ringkasan Fraud Rate, indikator, dan tabel Worst-5 cabang/cluster
# ---------------------------------------------------------------------------
def _fmt_rupiah_short(value):
    """Menyesuaikan gaya penulisan template: '240 JT' (tanpa desimal) / '1,18 M' (2 desimal, pakai koma)."""
    if value >= 1e9:
        return f"{value / 1e9:.2f} M".replace(".", ",")
    return f"{value / 1e6:.0f} JT"


def _fmt_diff_pct(new, old):
    """
    Format persentase perbedaan (`(new-old)/old*100`) dengan tanda +/- eksplisit
    dan dibulatkan ke integer, mis. "+63%"/"-12%". `old` bernilai 0/None/falsy
    → "+0%" (menghindari pembagian dengan nol).
    """
    if not old:
        return "+0%"
    pct = (new - old) / old * 100
    sign = "+" if pct >= 0 else ""
    return f"{sign}{round(pct)}%"


def _fmt_id_thousands(value):
    """Format angka dengan pemisah ribuan gaya Indonesia ('.'), dibulatkan ke integer; nilai tak valid → "0"."""
    try:
        return f"{int(round(value)):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "0"


def _apply_slide18_summary(slide_xml, data, region):
    """
    Mengganti angka ringkasan fraud YTD-25/YTD-26 di slide 18 (Kasus,
    Potloss, Potloss-per-Kasus, dan 3 label Diff%-nya), serta menyetel 3
    segitiga indikator naik/turun yang bersesuaian.

    Cara kerja PENGGANTIAN TEKS: berbeda dari kebanyakan fungsi lain di
    file ini yang mencocokkan sel/shape secara struktural, di sini
    penggantian dilakukan dengan mencari nilai LITERAL placeholder yang
    sudah ada di template (mis. teks "24", "39", "240 JT", "+392%" — nilai
    contoh milik region template) dan menggantinya satu-per-satu memakai
    `re.sub(..., count=1)` pada pola `<a:t>{old}</a:t>` — ini AMAN dilakukan
    di sini karena tiap string placeholder itu diasumsikan MUNCUL TEPAT
    SEKALI di slide (beda dengan sel tabel yang isinya sering berulang/
    identik antar baris, di mana `_rebuild_row_cells` dibutuhkan).

    Nilai kasus/potloss diambil dari `data["fraud"][region]`
    (`kasus_25/26`, `potloss_25/26`), lalu potloss-per-kasus (`ppk_25/26`)
    dihitung di sini (bukan dari loader) sebagai `potloss / kasus` (dengan
    penjagaan pembagian nol). Nilai-nilai itu diformat lewat
    `_fmt_rupiah_short` (kasus/potloss) dan `_fmt_diff_pct` (label Diff%),
    lalu dipasangkan dengan literal placeholder yang akan digantikannya
    dalam list `replacements`.

    Untuk indikator: arah (`direction`) tiap segitiga ditentukan dari
    perbandingan nilai 26 vs 25 (naik/turun/tetap), dipetakan ke NAMA SHAPE
    template yang sudah diketahui ("Isosceles Triangle 31/55/57" —
    di sini nama shape dipakai langsung karena stabil/sudah diverifikasi
    untuk slide ini, berbeda dengan pendekatan struktural
    `_compute_slide15_indicator_map` yang dipilih untuk slide 15 karena di
    sana ada 9 grup indikator yang lebih rawan renumbering). Tampilan tiap
    shape diterapkan lewat `_set_slide15_indicator_style` (fungsi yang
    sama dipakai ulang dari bagian slide 15, karena logika bentuk/warna
    indikatornya identik).

    Parameter: `slide_xml` (bytes XML slide 18), `data` (dict `load_all()`,
    key `"fraud"`), `region`.

    Return: bytes XML slide 18 dengan ringkasan dan indikator sudah
    diperbarui. Dipanggil dari `generate_pptx_for_region` (langkah 6),
    hasilnya diteruskan berantai ke dua pemanggilan
    `_apply_slide18_worst5_table` sebelum `editor.update(...)`.
    """
    f = data.get("fraud", {}).get(region, {})
    kasus_25, potloss_25 = f.get("kasus_25", 0), f.get("potloss_25", 0)
    kasus_26, potloss_26 = f.get("kasus_26", 0), f.get("potloss_26", 0)
    ppk_25 = (potloss_25 / kasus_25) if kasus_25 else 0
    ppk_26 = (potloss_26 / kasus_26) if kasus_26 else 0

    replacements = [
        ("240 JT", _fmt_rupiah_short(potloss_25)),
        ("24", str(kasus_25)),
        ("39", str(kasus_26)),
        ("1,18 M", _fmt_rupiah_short(potloss_26)),
        ("+392%", _fmt_diff_pct(potloss_26, potloss_25)),
        ("10 JT", _fmt_rupiah_short(ppk_25)),
        ("30 JT", _fmt_rupiah_short(ppk_26)),
        ("+63%", _fmt_diff_pct(kasus_26, kasus_25)),
        ("+203%", _fmt_diff_pct(ppk_26, ppk_25)),
    ]

    text = slide_xml.decode('utf-8')
    for old, new in replacements:
        text = re.sub(rf'(<a:t>){re.escape(old)}(</a:t>)', lambda m, v=new: m.group(1) + v + m.group(2),
                       text, count=1)

    def direction(new, old):
        if new > old:
            return "up"
        if new < old:
            return "down"
        return "flat"

    indicator_names = {
        "Isosceles Triangle 31": direction(potloss_26, potloss_25),
        "Isosceles Triangle 55": direction(kasus_26, kasus_25),
        "Isosceles Triangle 57": direction(ppk_26, ppk_25),
    }
    sp_blocks = re.findall(r'<p:sp\b.*?</p:sp>', text, re.DOTALL)
    for sp in sp_blocks:
        m = re.search(r'name="([^"]*)"', sp)
        if not m or m.group(1) not in indicator_names:
            continue
        new_sp = _set_slide15_indicator_style(sp, indicator_names[m.group(1)])
        if new_sp != sp:
            text = text.replace(sp, new_sp, 1)

    return text.encode('utf-8')


def _apply_slide18_worst5_table(slide_xml, data, region, table_index, data_key):
    """
    Mengisi salah satu dari 2 tabel Worst-5 potloss di slide 18 (Branch SSD
    / Cluster Collection — dipanggil dua kali oleh `generate_pptx_for_region`
    dengan `table_index`/`data_key` berbeda). Berbeda dari
    `_apply_slide16_branch_table`, di sini jumlah baris template SUDAH pas
    5 baris (tidak ada logika hapus-baris/redistribusi tinggi) — kalau data
    `region` kurang dari 5 baris, sisa baris diisi nilai kosong/0
    (`("", 0)` sebagai fallback saat `i >= len(rows_data)`).

    Parameter:
        slide_xml: bytes XML slide 18 (untuk pemanggilan kedua, ini adalah
                   hasil dari pemanggilan pertama, karena keduanya dirantai
                   berurutan di `generate_pptx_for_region`).
        data: dict `load_all()`, dipakai `data["fraud"][region][data_key]`
              (list `(nama, total_potloss)` — sudah terurut/dipilih
              Worst-5 oleh loader).
        region: region yang sedang diproses.
        table_index: indeks tabel (0-based) di antara `<a:tbl>` slide 18.
        data_key: key list Worst-5 di `data["fraud"][region]`, mis.
                  "worst5_branch_25"/"worst5_cluster_25".

    Return: bytes XML slide 18 dengan tabel `table_index` sudah terisi
    (nama diisi apa adanya, potloss diformat lewat `_fmt_id_thousands`;
    sel "No"/nomor urut di kolom pertama dilewati — `values[0] = None`).
    Dipanggil dari `generate_pptx_for_region` (langkah 6).
    """
    rows_data = data.get("fraud", {}).get(region, {}).get(data_key, [])

    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) <= table_index:
        return slide_xml
    table_xml = tbl_matches[table_index].group(0)
    row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    if len(row_blocks) < 2:
        return slide_xml

    new_table_xml = table_xml
    for i in range(len(row_blocks) - 1):
        row_xml = row_blocks[i + 1]
        name, total = rows_data[i] if i < len(rows_data) else ("", 0)
        values = [None, name, _fmt_id_thousands(total)]   # lewati sel peringkat "No"

        def transform(ci, tc, values=values):
            if values[ci] is None:
                return tc
            return re.sub(r'(<a:t>)[^<]*(</a:t>)',
                           lambda m, v=values[ci]: m.group(1) + v.replace('&', '&amp;') + m.group(2),
                           tc, count=1)

        new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
        if num_tcs != 3:
            continue
        new_table_xml = new_table_xml.replace(row_xml, new_row_xml, 1)

    text = text.replace(table_xml, new_table_xml, 1)
    return text.encode('utf-8')


def _build_slide_replacements(slide_idx, data, region):
    """
    Menyusun daftar tuple string `(old, new)` untuk penggantian teks generik
    satu slide tertentu (`slide_idx`, 0-indexed) — hasilnya diteruskan
    langsung ke `_replace_text_in_slide`/`replace_text_in_slide` (string
    replace polos, tanpa parsing XML).

    Setiap slide MINIMAL mendapat penggantian nama region dasar
    (`base = [(TEMPLATE_REGION, region)]`). Beberapa slide (indeks 3,
    "Slide 4") juga mendapat penggantian tambahan berupa angka total HC
    (`get_slide4_total_hc`, dari `chart_data.py`) menimpa placeholder
    `TEMPLATE_TOTAL_HC` (dari `config.py`). Slide lain (bercabang eksplisit
    di `if/elif` di bawah, mis. slide 3, 5-7, 10, 12, 14-16, 18) memang
    sengaja hanya mendapat `base` — cabang-cabang itu dipisah eksplisit
    (bukan satu `else` polos) sebagai dokumentasi implisit slide mana saja
    yang SUDAH DIVERIFIKASI cukup diproses dengan penggantian nama region
    generik saja (slide-slide dengan tabel/highlight, mis. 14-16 dan 18,
    isinya sudah ditangani terpisah oleh fungsi `_apply_slideNN_*` sebelum
    fungsi ini dipanggil, sehingga di sini tinggal mengganti nama region
    di teks judul/label yang tersisa).

    Parameter: `slide_idx` (int, index slide 0-based), `data` (dict
    `load_all()`, dipakai hanya untuk cabang slide 4), `region`.

    Return: list `(old, new)`. Dipanggil dari `generate_pptx_for_region`
    (langkah 7) untuk SETIAP slide di presentasi; hasilnya (kalau tidak
    kosong) diteruskan ke `_replace_text_in_slide`.
    """
    tr = TEMPLATE_REGION

    # TEMPLATE_ABBREV ("Jateng") tidak dipakai di sini: template saat ini
    # hanya membawa satu token "REGION", bukan token nama-singkat terpisah.
    # REGION_ABBREV masih dipakai untuk teks label chart (lihat
    # get_slide3_annotations dsb) tapi tidak lagi menggerakkan pencarian/
    # penggantian teks slide.
    base = [(tr, region)]

    if slide_idx == 0:   # Slide 1
        return base
    elif slide_idx == 2:  # Slide 3
        return base
    elif slide_idx == 3:  # Slide 4
        hc = get_slide4_total_hc(data, region)
        return base + [(TEMPLATE_TOTAL_HC, hc)]
    elif slide_idx in (4, 5, 6):  # Slides 5-7
        return base
    elif slide_idx == 9:  # Slide 10
        return base
    elif slide_idx == 11:  # Slide 12
        return base
    elif slide_idx in (13, 14, 15):  # Slides 14-16
        return base
    elif slide_idx == 17:  # Slide 18
        return base
    else:
        return base


# ---------------------------------------------------------------------------
# Generator PPTX utama per region
# ---------------------------------------------------------------------------
def generate_pptx_for_region(region, data, verbose=True):
    """
    TITIK MASUK (entry point) modul ini: menghasilkan satu file PPTX HR
    Dashboard untuk satu `region`, dipanggil dari `main.py`
    (`generate_pptx_for_region(region, data, verbose=True)`) sekali per
    region dalam sebuah loop.

    Cara kerja keseluruhan: membuka `TEMPLATE_PATH` (dari `config.py`) lewat
    `PptxEditor` (context manager di `xml_updater.py` yang membaca seluruh
    isi zip PPTX ke memori dan menampung perubahan sebagai
    `{path: bytes baru}` sampai `editor.save()` dipanggil), lalu menjalankan
    10 langkah berurutan yang masing-masing memanggil satu atau beberapa
    fungsi `_apply_slideNN_*`/`_compute_slideNN_*` yang didefinisikan di
    bagian atas file ini. Urutan antar-langkah TIDAK sepenuhnya bebas —
    lihat catatan pada langkah 3 di bawah soal kenapa slide 14 harus
    diproses SEBELUM langkah 7 (penggantian teks generik).

    Parameter:
        region: nama region (str) yang sedang di-generate, mis. "Jawa
                Tengah", "Jawa Timur", dsb (nilai dari `REGIONS` di
                `config.py`, dioper oleh loop pemanggil di `main.py`).
        data: dict besar hasil `load_all()` (`data_loader.py`), berisi
              semua sumber data per slide (`slide3`, `s4_npat`, `s4_wc`,
              `s10`, `attrition`, `s15_func`, `s16`, `fraud`, dsb) — dioper
              apa adanya ke hampir semua fungsi `_apply_*`/`_compute_*`
              yang lalu mengambil sub-key yang relevan untuk slide masing-
              masing.
        verbose: jika True, mencetak progres ke stdout (path output, jumlah
                 chart ditemukan) — tidak memengaruhi logika, murni logging.

    Return: path (str) file PPTX yang sudah disimpan
    (`OUTPUT_DIR/HR_Dashboard_{region_dengan_underscore}.pptx`).
    """
    safe_name = region.replace(" ", "_")
    out_path = os.path.join(OUTPUT_DIR, f"HR_Dashboard_{safe_name}.pptx")

    if verbose:
        print(f"  Generating: {out_path}")

    with PptxEditor(TEMPLATE_PATH) as editor:
        # 1. Menemukan semua pasangan file chart/workbook ter-embed di dalam
        #    paket PPTX (lihat _discover_charts) — dilakukan secara dinamis
        #    lewat penelusuran isi zip, bukan daftar nomor chart yang
        #    di-hardcode, supaya tetap benar walau jumlah/nomor chart di
        #    template berubah.
        chart_map = _discover_charts(editor)
        if verbose:
            print(f"    Found {len(chart_map)} charts")

        # 2. Memperbarui isi tiap chart (data kategori/series) + workbook
        #    Excel ter-embed di baliknya (lihat _update_chart). Kegagalan
        #    pada satu chart (mis. data region tidak lengkap untuk chart
        #    itu) hanya dicatat sebagai warning dan tidak menghentikan
        #    proses chart-chart lainnya.
        for chart_num, chart_info in sorted(chart_map.items()):
            try:
                _update_chart(editor, chart_num, chart_info, data, region)
            except Exception as e:
                print(f"    [WARN] chart{chart_num}: {e}")

        # 3. Mengisi 3 tabel YoY attrition NR/RG/Total slide 14
        #    (_apply_slide14_tables), memindahkan 3 kotak highlightnya
        #    (_apply_slide14_rect_highlights), lalu mengisi tabel "Top 3
        #    Reason Out" (_apply_slide14_reason_table).
        #    HARUS dijalankan SEBELUM pass penggantian teks generik di
        #    langkah 7 di bawah: salah satu sel tabel masih berupa
        #    placeholder literal "{REGION}" yang perlu diselesaikan
        #    ("Jateng", label baris tetapnya) lewat logika eliminasi kita
        #    sendiri (_resolve_slide14_table_rows) TERLEBIH DAHULU —
        #    kalau tidak, substitusi generik (TEMPLATE_REGION, region) di
        #    langkah 7 akan menimpanya dengan nama region OUTPUT saat ini,
        #    sehingga baris region itu jadi terduplikasi (dua baris dengan
        #    nama region yang sama).
        slide_paths_early = _get_slide_paths(editor)
        if len(slide_paths_early) > 13:
            slide14_path = slide_paths_early[13]
            slide14_xml = editor.read(slide14_path)
            slide14_xml = _apply_slide14_tables(slide14_xml, data)
            slide14_xml = _apply_slide14_rect_highlights(slide14_xml, region)
            slide14_xml = _apply_slide14_reason_table(slide14_xml, data, region)
            editor.update(slide14_path, slide14_xml)

        # 4. Mengisi 2 tabel "Attrition Report by Function" slide 15
        #    (_apply_slide15_tables) lalu menyetel 9×3 indikator naik/
        #    turun/tetapnya (_apply_slide15_indicators).
        if len(slide_paths_early) > 14:
            slide15_path = slide_paths_early[14]
            slide15_xml = editor.read(slide15_path)
            slide15_xml = _apply_slide15_tables(slide15_xml, data, region)
            slide15_xml = _apply_slide15_indicators(slide15_xml, data, region)
            editor.update(slide15_path, slide15_xml)

        # 5. Mengisi tabel field-attrition slide 16: tabel per-cabang Sales
        #    (table_index=0) dan per-cluster Collection (table_index=1)
        #    lewat _apply_slide16_branch_table (dua kali, dengan
        #    data_key/grand_total_key berbeda), lalu tabel reason-out
        #    (_apply_slide16_reason_table).
        if len(slide_paths_early) > 15:
            slide16_path = slide_paths_early[15]
            slide16_xml = editor.read(slide16_path)
            slide16_xml = _apply_slide16_branch_table(
                slide16_xml, data, region, table_index=0,
                data_key="branch_sales", grand_total_key="Sales Officer")
            slide16_xml = _apply_slide16_branch_table(
                slide16_xml, data, region, table_index=1,
                data_key="branch_collection", grand_total_key="Collection Officer")
            slide16_xml = _apply_slide16_reason_table(slide16_xml, data, region)
            editor.update(slide16_path, slide16_xml)

        # 6. Mengisi ringkasan fraud + indikatornya slide 18
        #    (_apply_slide18_summary), lalu 2 tabel Worst-5 potloss
        #    (_apply_slide18_worst5_table): table_index=0 untuk Branch SSD,
        #    table_index=1 untuk Cluster Collection.
        if len(slide_paths_early) > 17:
            slide18_path = slide_paths_early[17]
            slide18_xml = editor.read(slide18_path)
            slide18_xml = _apply_slide18_summary(slide18_xml, data, region)
            slide18_xml = _apply_slide18_worst5_table(slide18_xml, data, region, table_index=0, data_key="worst5_branch_25")
            slide18_xml = _apply_slide18_worst5_table(slide18_xml, data, region, table_index=1, data_key="worst5_cluster_25")
            editor.update(slide18_path, slide18_xml)

        # 7. Pass penggantian teks generik untuk SEMUA slide: nama region
        #    (dan angka total HC khusus slide 4) lewat
        #    _build_slide_replacements + _replace_text_in_slide. Slide yang
        #    tabelnya sudah diisi di langkah 3-6 di atas (14, 15, 16, 18)
        #    tetap ikut lewat sini untuk mengganti teks judul/label lain di
        #    luar tabel (mis. nama region di judul slide) — bukan berarti
        #    tabelnya diproses ulang di sini.
        slide_paths = _get_slide_paths(editor)
        for slide_idx, slide_path in enumerate(slide_paths):
            replacements = _build_slide_replacements(slide_idx, data, region)
            if not replacements:
                continue
            slide_xml = editor.read(slide_path)
            updated = _replace_text_in_slide(slide_xml, replacements)
            editor.update(slide_path, updated)

        # 8. Memindahkan kotak highlight merah slide 3
        #    (_compute_slide3_rect_positions + _apply_slide3_rect_updates)
        #    dan segitiga indikator + text box YoY%-nya
        #    (_compute_slide3_indicator_updates + _apply_slide3_indicator_updates)
        #    ke posisi bar milik `region` pada masing-masing dari 6 chart.
        #    Dijalankan setelah langkah 7 (bukan sebelum, seperti slide 14)
        #    karena slide 3 tidak punya masalah placeholder ambigu yang
        #    sama — urutan di sini tidak kritikal terhadap langkah 7.
        if len(slide_paths) > 2:
            slide3_path = slide_paths[2]
            slide3_xml = editor.read(slide3_path)

            rect_updates = _compute_slide3_rect_positions(data, region)
            if rect_updates:
                slide3_xml = _apply_slide3_rect_updates(slide3_xml, rect_updates)

            indicator_updates = _compute_slide3_indicator_updates(data, region)
            if indicator_updates:
                slide3_xml = _apply_slide3_indicator_updates(slide3_xml, indicator_updates)

            slide3_xml = _apply_slide3_highlights(slide3_xml, data, region)

            editor.update(slide3_path, slide3_xml)

        # 9. Mengisi tabel training BSC slide 10 (_apply_slide10_table —
        #    sama untuk semua region) lalu memindahkan kotak highlight
        #    merahnya ke baris `region` (_apply_slide10_rect_highlight).
        if len(slide_paths) > 9:
            slide10_path = slide_paths[9]
            slide10_xml = editor.read(slide10_path)
            slide10_xml = _apply_slide10_table(slide10_xml, data)
            slide10_xml = _apply_slide10_rect_highlight(slide10_xml, region)
            editor.update(slide10_path, slide10_xml)

        # 10. Simpan seluruh perubahan yang terkumpul di `editor` sebagai
        #     file PPTX baru di `out_path` (menulis ulang zip, menjaga tipe
        #     kompresi tiap entry seperti aslinya — lihat PptxEditor.save
        #     di xml_updater.py).
        editor.save(out_path)

    return out_path
