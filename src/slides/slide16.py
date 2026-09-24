"""
Slide 16 — Field Regretted Attrition: tabel per-cabang/cluster Sales &
Collection (dengan dukungan region berbagang/lebih-sedikit dari baris
template), tabel reason-out, dan highlight/insight otomatisnya.
"""
import re

from ..pptx_common import (
    _rebuild_row_cells, _apply_highlight_tokens,
    _fmt_slide15_num, _fmt_slide15_pct, _TCPR_TAIL_OPTIONAL_RE,
)
from ..chart_data import get_slide16_highlights

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
    sedangkan template punya jumlah baris tabel yang TETAP, dan tabelnya
    harus SELALU MUAT pada tinggi total yang sama (`total_data_height`,
    dijumlah dari tinggi baris data asli template) supaya tidak menabrak
    konten lain di slide. Dua arah:
      - Region dengan cabang LEBIH SEDIKIT dari baris template: sisa baris
        tak terpakai DIHAPUS SELURUHNYA (bukan dibiarkan kosong/bernilai
        0), dan tinggi baris yang dipertahankan DIHITUNG ULANG lebih besar
        dari baseline (`total_data_height // n_branches`) supaya tabel
        tetap mengisi ruang vertikal aslinya secara utuh.
      - Region dengan cabang LEBIH BANYAK dari baris template (mis. Jabar
        punya 34 cluster Collection tapi template cuma 15 baris data —
        bug nyata yang pernah terjadi sebelum perbaikan ini): 2 baris
        TERAKHIR di template di-CLONE bergantian (bukan cuma 1 baris
        diulang — itu akan merusak pola warna selang-seling/banding
        antar-baris, lihat catatan di dekat `alt_templates`) sebanyak
        kekurangannya supaya SEMUA baris data region ini tampil (bukan
        diam-diam dibuang), TAPI
        tinggi tiap baris (termasuk baris asli) ikut MENGECIL dari
        baseline (`total_data_height // n_branches` juga, sekarang lebih
        kecil dari baseline karena `n_branches` lebih besar dari
        `data_row_count`) supaya tinggi total tabel tidak melebihi
        alokasi aslinya. Ukuran font tiap sel ikut diskalakan turun
        sebanding (`font_scale`, dibatasi minimum `_FONT_SCALE_MIN` biar
        tidak sampai tidak terbaca) supaya teksnya tetap muat di baris
        yang lebih pendek itu, bukan terpotong/tumpang tindih.

    KENAPA DIBANGUN ULANG LEWAT POSISI (match span dari `re.finditer`),
    BUKAN `str.replace` berbasis isi seperti versi sebelumnya: baris-baris
    tabel yang BELUM diisi (masih placeholder template) seringkali
    BYTE-IDENTICAL satu sama lain (semua "0"/kosong) — `str.replace(...,
    count=1)` pada teks yang identik akan selalu mengenai kemunculan
    PALING KIRI yang tersisa, jadi kalau langkah HAPUS dan langkah ISI
    dipisah jadi 2 loop berurutan (bukan diproses baris demi baris
    berurutan sesuai posisi), baris yang terhapus/tersisipi bisa jadi
    baris yang SALAH (mis. clone baru malah disisipkan setelah baris
    PERTAMA, bukan baris TERAKHIR). Membangun tabel baru dari potongan
    `table_xml[a:b]` berdasarkan `.start()`/`.end()` match menghindari
    ambiguitas ini sepenuhnya — tidak pernah bergantung pada isi baris
    untuk tahu itu baris yang mana.

    Baris grand-total ("REGION – ..." — baris terakhir, tetap, tidak ikut
    dihapus/di-clone) hanya nilai-nilainya yang diperbarui, memakai agregat
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
    (baris tak terpakai dihapus ATAU baris tambahan di-clone sesuai
    kebutuhan, tinggi baris disesuaikan, grand-total diperbarui). Dipanggil
    dua kali dari `generate_pptx_for_region` (langkah 5), hasilnya
    diteruskan ke pemanggilan berikutnya (untuk tabel lain) lalu ke
    `_apply_slide16_reason_table`.
    """
    branch_rows = data.get("s16", {}).get(region, {}).get(data_key, [])
    grand_total = data.get("s15_func", {}).get(region, {}).get("y26", [])
    grand_total = next((r for r in grand_total if r["label"] == grand_total_key), None)

    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) <= table_index:
        return slide_xml
    table_match = tbl_matches[table_index]
    table_xml = table_match.group(0)
    row_matches = list(re.finditer(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL))
    if len(row_matches) < 2:
        return slide_xml

    data_row_count = len(row_matches) - 2   # dikurangi baris header dan baris grand-total
    n_branches = len(branch_rows)

    def cell_values(label, r):
        if r is None:
            return [label, "0", "0", "0", "0", "0,00%", "0,00%", "0,00%"]
        return [label,
                _fmt_slide15_num(r["avg_hc"]), _fmt_slide15_num(r["out_nr"]),
                _fmt_slide15_num(r["out_rg"]), _fmt_slide15_num(r["out_total"]),
                _fmt_slide15_pct(r["pct_nr"]), _fmt_slide15_pct(r["pct_rg"]), _fmt_slide15_pct(r["pct_total"])]

    # Kolom "Out Regr" (index 3) dan "%Regret" (index 6) di template punya
    # fill merah (FF0000) yang di-hardcode STATIS pada beberapa baris
    # teratas — cocok untuk region demo (Jawa Tengah) yang datanya memang
    # diurutkan menurun berdasar out_rg, tapi TIDAK ikut disesuaikan saat
    # baris diisi ulang untuk region lain, jadi highlight-nya bisa
    # nyasar ke baris yang out_rg/pct_rg-nya sebenarnya 0. Di sini
    # highlight dihitung ULANG per baris sesuai data region yang
    # sebenarnya: merah HANYA kalau nilai kolom itu sendiri bukan 0. Warna
    # teks di dalam sel ikut disetel di sini juga (putih saat merah,
    # hitam saat tidak) — teks hitam di atas fill merah nyaris tidak
    # terbaca, dan karena logika ini dijalankan SERAGAM untuk semua baris
    # (baik baris asli template maupun baris hasil clone untuk region
    # dengan cabang/cluster lebih banyak dari baris template), warna
    # latar+teksnya otomatis ikut benar juga untuk baris tambahan itu.
    _REGRET_HIGHLIGHT_COLS = {3: "out_rg", 6: "pct_rg"}
    _RPR_RE = re.compile(r'(<a:rPr\b[^>]*>)(.*?)(</a:rPr>)', re.DOTALL)

    def set_first_run_color(tc, color_hex):
        def repl(m):
            inner = re.sub(r'<a:solidFill>.*?</a:solidFill>',
                            f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>',
                            m.group(2), count=1)
            return m.group(1) + inner + m.group(3)
        return _RPR_RE.sub(repl, tc, count=1)

    # Skala ukuran font sel data (`sz="NNN"`, satuan ratusan poin) sebanding
    # dengan seberapa jauh `new_row_height` menyusut dari tinggi baris
    # ASLI template (`baseline_row_height`) — supaya saat baris di-clone
    # banyak-banyak (region dengan cabang/cluster jauh lebih banyak dari
    # baris template, mis. Jabar 41 branch Sales vs 30 baris template),
    # tabelnya TETAP MUAT pada tinggi asli yang dialokasikan template
    # (tidak melebar ke bawah menabrak konten lain di slide), dengan teks
    # yang ikut mengecil alih-alih terpotong/tumpang tindih. `_FONT_SCALE_MIN`
    # jadi batas bawah supaya teks tidak sampai tidak terbaca sama sekali
    # untuk region dengan jumlah baris ekstrem.
    _FONT_SCALE_MIN = 0.55
    _SZ_RE = re.compile(r'sz="(\d+)"')

    def scale_fonts(row_xml, scale):
        if scale >= 0.999:
            return row_xml
        return _SZ_RE.sub(lambda m: f'sz="{max(400, round(int(m.group(1)) * scale))}"', row_xml)

    def fill_row(row_xml, height, values, r, font_scale):
        def transform(ci, tc, values=values, r=r):
            new_tc = re.sub(r'(<a:t>)[^<]*(</a:t>)',
                             lambda m, v=values[ci]: m.group(1) + v.replace('&', '&amp;') + m.group(2),
                             tc, count=1)
            field = _REGRET_HIGHLIGHT_COLS.get(ci)
            if field is not None:
                nyala = bool(r is not None and (r.get(field) or 0) != 0)
                fill = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' if nyala else '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
                new_tc = _TCPR_TAIL_OPTIONAL_RE.sub(rf'\1{fill}\2', new_tc, count=1)
                new_tc = set_first_run_color(new_tc, "FFFFFF" if nyala else "000000")
            return new_tc
        new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
        if num_tcs != 8:
            return row_xml
        new_row_xml = scale_fonts(new_row_xml, font_scale)
        return _TR_OPEN_RE.sub(f'<a:tr h="{height}"', new_row_xml, count=1)

    data_row_matches = row_matches[1:1 + data_row_count]
    total_data_height = sum(int(m.group(0)[9:-1]) for m in
                             (re.match(r'<a:tr h="(\d+)"', dm.group(0)) for dm in data_row_matches) if m)
    baseline_row_height = (total_data_height // data_row_count) if data_row_count else 0

    # Tinggi baris SELALU dihitung ulang dari total budget tinggi asli
    # dibagi rata ke `n_branches` baris — supaya tinggi TOTAL tabel selalu
    # sama dengan yang dialokasikan template, baik saat baris dikurangi
    # (row lebih tinggi dari baseline) maupun saat baris ditambah (row
    # lebih pendek dari baseline, dikompensasi lewat `font_scale`).
    new_row_height = (total_data_height // n_branches) if n_branches else 0
    font_scale = max(_FONT_SCALE_MIN, min(1.0, new_row_height / baseline_row_height)) if baseline_row_height else 1.0

    if n_branches >= data_row_count:
        # Sama atau lebih banyak: pakai SEMUA baris template apa adanya
        # sebagai basis, lalu clone baris TAMBAHAN sebanyak kekurangannya.
        #
        # PENTING soal warna selang-seling (banding): baris-baris template
        # berselang-seling warna latar per PARITAS indeksnya (mis. genap =
        # putih/noFill, ganjil = tint pink pucat — lihat cara kerja
        # `_REGRET_HIGHLIGHT_COLS` di atas untuk kolom yang memang selalu
        # diubah; kolom LAINNYA tidak pernah disentuh tcPr-nya sama sekali
        # di `transform` di bawah, jadi warna aslinya ikut baris template
        # yang di-clone). Dulu SEMUA baris tambahan meng-clone baris
        # TERAKHIR saja, sehingga semuanya jadi SATU warna yang sama
        # (parity baris terakhir), merusak pola selang-seling. Di sini
        # baris tambahan mengulang 2 baris template TERAKHIR
        # (`row_templates[-2]` lalu `row_templates[-1]`) secara bergantian
        # — parity baris tepat setelah baris template terakhir SELALU
        # sama dengan `row_templates[-2]` (beda 2 indeks = parity sama),
        # jadi urutan ini melanjutkan pola selang-seling tanpa putus.
        row_templates = [dm.group(0) for dm in data_row_matches]
        extra_needed = n_branches - data_row_count
        if data_row_count >= 2:
            alt_templates = [row_templates[-2], row_templates[-1]]
        elif data_row_count == 1:
            alt_templates = [row_templates[-1]]
        else:
            alt_templates = []
        if alt_templates:
            row_templates += [alt_templates[k % len(alt_templates)] for k in range(extra_needed)]
    else:
        # Lebih sedikit: pakai HANYA `n_branches` baris pertama sebagai
        # basis (sisanya dihapus dengan sendirinya karena tidak diikutkan
        # ke `row_templates`).
        row_templates = [dm.group(0) for dm in data_row_matches[:n_branches]]

    filled_rows = []
    for i in range(n_branches):
        r = branch_rows[i]
        filled_rows.append(fill_row(row_templates[i], new_row_height, cell_values(r["label"], r), r, font_scale))

    gt_row_xml = row_matches[-1].group(0)
    gt_values = cell_values(None, grand_total)   # index 0 (label) sengaja dibiarkan di bawah

    def gt_transform(ci, tc, gt_values=gt_values):
        if ci == 0:
            return tc
        return re.sub(r'(<a:t>)[^<]*(</a:t>)',
                       lambda m, v=gt_values[ci]: m.group(1) + v + m.group(2),
                       tc, count=1)

    new_gt_row, num_gt_tcs = _rebuild_row_cells(gt_row_xml, gt_transform)
    if num_gt_tcs != 8:
        new_gt_row = gt_row_xml

    header_xml = row_matches[0].group(0)
    new_table_xml = (table_xml[:row_matches[0].start()] + header_xml +
                      ''.join(filled_rows) + new_gt_row +
                      table_xml[row_matches[-1].end():])

    text = text[:table_match.start()] + new_table_xml + text[table_match.end():]
    return text.encode('utf-8')


def _fmt_slide16_reason_num(v):
    """Format angka bulat (tanpa desimal) untuk tabel reason-out slide 16; nilai tak valid → "0"."""
    try:
        return str(int(v))
    except (TypeError, ValueError):
        return "0"


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


def _apply_slide16_highlight(slide_xml, data, region):
    """
    Mengisi 1 placeholder insight slide 16 ({SLIDE16_INSIGHT}) dengan teks
    draf hasil `get_slide16_highlights()` di chart_data.py, lewat
    `_apply_highlight_tokens`.

    Parameter: `slide_xml` (bytes XML slide 16), `data` (dict load_all()),
    `region`.

    Return: bytes XML slide 16 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 5), setelah
    `_apply_slide16_reason_table`.
    """
    return _apply_highlight_tokens(slide_xml, "SLIDE16", {"INSIGHT": get_slide16_highlights(data, region)})


# ---------------------------------------------------------------------------
# Slide 18 – Ringkasan Fraud Rate, indikator, dan tabel Worst-5 cabang/cluster
# ---------------------------------------------------------------------------
