"""
Slide 4 — Ringkasan HC/NPAT/Function/Frontliners/Span-of-Control: 3 tabel
statis persentase AGE/EDU/LOS (sebelumnya tidak pernah diisi ulang per
region) dan highlight/insight otomatisnya.
"""
import re

from ..pptx_common import _rebuild_row_cells, _apply_highlight_tokens
from ..chart_data import get_slide4_highlights, get_slide4_lea_pct, _fmt_slide4_pct

def _apply_slide4_lea_tables(slide_xml, data, region):
    """
    Mengisi 3 tabel statis persentase AGE/EDU/LOS di slide 4 (tepat di
    bawah chart histogram 10/11/12) dengan angka hasil
    `get_slide4_lea_pct()` di chart_data.py — tabel-tabel ini SEBELUMNYA
    TIDAK PERNAH disentuh kode apa pun, jadi selalu menampilkan nilai
    contoh statis milik region demo template untuk SEMUA region.

    Struktur & urutan tabel (ditemukan lewat urutan kemunculan `<a:tbl>`
    di XML, BUKAN nama shape — tabel-tabel ini tidak bernama unik):
      tbl[0] = AGE: 2 baris (NAS lalu REG), masing-masing label + 5 sel
               nilai (urutan AGE_ORDER menaik).
      tbl[1] = EDU: sama seperti AGE tapi 4 sel nilai (EDU_ORDER menaik).
      tbl[2] = LOS: 1 baris header ("REG"/"NAS", dilewati) + 6 baris data,
               2 sel per baris (REG, NAS) — urutan barisnya MENAIK
               (a. <1 thn -> f. > 20th), KEBALIKAN dari urutan LOS_ORDER
               yang dipakai chart10. `get_slide4_lea_pct()` SUDAH
               membalik urutan ini sebelum dikembalikan, jadi di sini
               tinggal dipasangkan apa adanya sesuai posisi baris —
               JANGAN membalik lagi di sini, supaya tidak tertukar/
               terbalik dua kali.

    Validasi jumlah sel per baris dicek sebelum ditulis (`num_tcs`
    hasil `_rebuild_row_cells`); kalau tidak cocok dengan jumlah nilai
    yang tersedia, baris itu dilewati apa adanya (tidak menghentikan
    baris/tabel lain).

    Parameter: `slide_xml` (bytes XML slide 4), `data` (dict load_all()),
    `region`.

    Return: bytes XML slide 4 dengan ketiga tabel sudah diisi. Dipanggil
    dari `generate_pptx_for_region` (langkah 3a), sebelum
    `_apply_slide4_highlights` (independen, tapi tabel ini dibaca ulang
    oleh insight highlight sehingga urutan tidak masalah).
    """
    pct = get_slide4_lea_pct(data, region)
    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) < 3:
        return slide_xml

    def fill_row(row_xml, values, skip_first=True):
        def transform(ci, tc, values=values, skip_first=skip_first):
            vi = ci - 1 if skip_first else ci
            if skip_first and ci == 0:
                return tc
            if vi < 0 or vi >= len(values):
                return tc
            return re.sub(r'(<a:t>)[^<]*(</a:t>)',
                           lambda m, v=_fmt_slide4_pct(values[vi]): m.group(1) + v + m.group(2),
                           tc, count=1)
        new_row, _ = _rebuild_row_cells(row_xml, transform)
        return new_row

    # tbl[0]=AGE, tbl[1]=EDU: baris 0 = NAS, baris 1 = REG.
    for tbl_idx, key in ((0, "AGE"), (1, "EDU")):
        table_xml = tbl_matches[tbl_idx].group(0)
        rows = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
        if len(rows) != 2:
            continue
        new_table_xml = table_xml
        new_table_xml = new_table_xml.replace(rows[0], fill_row(rows[0], pct[key]["nas"]), 1)
        new_table_xml = new_table_xml.replace(rows[1], fill_row(rows[1], pct[key]["reg"]), 1)
        text = text.replace(table_xml, new_table_xml, 1)

    # tbl[2]=LOS: baris 0 = header (dilewati), baris 1-6 = data (REG, NAS).
    table_xml = tbl_matches[2].group(0)
    rows = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    if len(rows) == 7:
        new_table_xml = table_xml
        for i in range(6):
            row_values = [pct["LOS"]["reg"][i], pct["LOS"]["nas"][i]]
            new_row = fill_row(rows[i + 1], row_values, skip_first=False)
            new_table_xml = new_table_xml.replace(rows[i + 1], new_row, 1)
        text = text.replace(table_xml, new_table_xml, 1)

    return text.encode('utf-8')


def _apply_slide4_highlights(slide_xml, data, region):
    """
    Mengisi 3 placeholder highlight slide 4 ({SLIDE4_HIGHLIGHT_A/B/C})
    dengan teks draf hasil `get_slide4_highlights()` di chart_data.py,
    lewat `_apply_highlight_tokens`.

    Parameter: `slide_xml` (bytes XML slide 4), `data` (dict load_all()),
    `region`.

    Return: bytes XML slide 4 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 3b).
    """
    return _apply_highlight_tokens(slide_xml, "SLIDE4_HIGHLIGHT", get_slide4_highlights(data, region))


