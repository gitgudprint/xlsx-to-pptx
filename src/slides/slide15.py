"""
Slide 15 — Attrition Report by Function (2 tabel + 9x3 indikator
naik/turun/tetap) dan highlight/insight otomatisnya.
"""
import re
from lxml import etree

from ..pptx_common import (
    _rebuild_row_cells, _apply_highlight_tokens,
    _fmt_slide15_num, _fmt_slide15_pct, _set_slide15_indicator_style,
)
from ..chart_data import get_slide15_highlights

_SLIDE15_INDICATOR_GROUPS = [
    "Group 2", "Group 11", "Group 18", "Group 23",
    "Group 27", "Group 31", "Group 35", "Group 39", "Group 43",
]
_SLIDE15_PCT_KEYS = ["pct_nr", "pct_rg", "pct_total"]


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

    Return: dict `{nama_shape (str): (row_idx, metric_idx, group_flip)}` —
    `group_flip` (bool) menandai apakah grup pembungkus shape ini punya
    flipV="1" di level grup (lihat catatan di atas). Dipanggil dari
    `_apply_slide15_indicators`, yang lalu mencocokkan nama tiap shape
    `<p:sp>` di slide terhadap dict ini untuk tahu baris/metrik mana yang
    diwakili shape tersebut, dan mengompensasi `group_flip` saat menyetel
    flip individual shape itu.
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
        # Sebagian grup (mis. "Group 18" / baris Sales Support) punya
        # flipV="1" pada `<p:grpSpPr><a:xfrm>` GRUP itu sendiri — beda
        # dengan grup lain yang tidak punya flip di level grup. Flip di
        # level grup ini menambah SATU flip lagi di atas flip masing-masing
        # segitiga anak, jadi kalau tidak dikompensasi, arah panah
        # (naik/turun) yang terlihat jadi TERBALIK khusus untuk baris itu,
        # meskipun warnanya sendiri (dihitung terpisah dari flip) tetap
        # benar. `group_flip` dicatat di sini supaya
        # `_apply_slide15_indicators` bisa membalik nilai flip yang
        # di-set ke tiap anak sebagai kompensasi.
        grp_xfrm = grp.find('./p:grpSpPr/a:xfrm', ns)
        group_flip = grp_xfrm is not None and grp_xfrm.get('flipV') == '1'
        shapes = []
        for sp in grp.findall('./p:sp', ns):
            name = sp.find('.//p:cNvPr', ns).get('name')
            off = sp.find('.//a:xfrm/a:off', ns)
            shapes.append((int(off.get('x')), name))
        shapes.sort()
        for metric_idx, (_, name) in enumerate(shapes):
            mapping[name] = (row_idx, metric_idx, group_flip)
    return mapping


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
    `_set_slide15_indicator_style(sp, direction, group_flip)` menerapkan
    tampilannya — `group_flip` diteruskan apa adanya dari
    `_compute_slide15_indicator_map` untuk mengompensasi flip di level
    grup (lihat catatan di fungsi itu dan di `_set_slide15_indicator_style`).

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
        row_idx, metric_idx, group_flip = indicator_map[m.group(1)]
        if row_idx >= len(y26):
            continue
        key = _SLIDE15_PCT_KEYS[metric_idx]
        v26, v25 = y26[row_idx].get(key), y25[row_idx].get(key)
        if v26 is None or v25 is None:
            continue
        direction = "down" if v26 < v25 else ("up" if v26 > v25 else "flat")
        new_sp = _set_slide15_indicator_style(sp, direction, group_flip)
        if new_sp != sp:
            text = text.replace(sp, new_sp, 1)

    return text.encode('utf-8')


def _apply_slide15_highlights(slide_xml, data, region):
    """
    Mengisi 3 placeholder highlight slide 15 ({SLIDE15_HIGHLIGHT_A/B/C} —
    tren YoY + penyebab, fungsi yang bergerak berlawanan arah dari Grand
    Total, slot kosong) dengan teks draf hasil `get_slide15_highlights()`
    di chart_data.py, lewat `_apply_highlight_tokens`.

    Parameter: `slide_xml` (bytes XML slide 15 — dipanggil setelah
    `_apply_slide15_indicators`, meski independen), `data` (dict
    load_all()), `region`.

    Return: bytes XML slide 15 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 4).
    """
    return _apply_highlight_tokens(slide_xml, "SLIDE15_HIGHLIGHT", get_slide15_highlights(data, region))


# ---------------------------------------------------------------------------
# Slide 16 – Field Regretted Attrition (tabel per-cabang/cluster + tabel reason-out)
# ---------------------------------------------------------------------------
