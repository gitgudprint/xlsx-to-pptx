"""
Slide 18 — Ringkasan Fraud Rate + indikator, 2 tabel Worst-5 potloss
(Branch SSD & Cluster Collection), dan highlight/insight otomatisnya.
"""
import re

from ..pptx_common import _rebuild_row_cells, _apply_highlight_tokens, _set_slide15_indicator_style
from ..chart_data import get_slide18_highlights

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


def _apply_slide18_highlights(slide_xml, data, region):
    """
    Mengisi token angka {HIGHLIGHT_FRAUDRATE_YOY} dan 4 slot
    {SLIDE18_HIGHLIGHT_A..D} di slide 18 dengan teks draf hasil
    `get_slide18_highlights()` di chart_data.py, lewat
    `_apply_highlight_tokens` (dipanggil dua kali dengan prefix berbeda,
    karena {HIGHLIGHT_FRAUDRATE_YOY} tidak memakai prefix "SLIDE18_"
    seperti 4 slot lainnya — beda penamaan token ini murni ikut apa yang
    sudah ada di template).

    Parameter: `slide_xml` (bytes XML slide 18), `data` (dict load_all()),
    `region`.

    Return: bytes XML slide 18 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 6), setelah
    `_apply_slide18_worst5_table`.
    """
    highlights = get_slide18_highlights(data, region)
    slide_xml = _apply_highlight_tokens(slide_xml, "HIGHLIGHT", {"FRAUDRATE_YOY": highlights["FRAUDRATE_YOY"]})
    slide_xml = _apply_highlight_tokens(slide_xml, "SLIDE18_HIGHLIGHT",
                                         {k: v for k, v in highlights.items() if k != "FRAUDRATE_YOY"})
    return slide_xml


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


