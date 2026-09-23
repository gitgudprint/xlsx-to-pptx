"""
Manipulasi XML tingkat rendah: memperbarui cache XML chart dan teks slide.

Semua modifikasi menggunakan regex pada raw byte UTF-8, bukan parse/reserialize,
sehingga deklarasi XML asli, lokasi namespace, dan struktur file tetap
terjaga byte-demi-byte. Ini mencegah munculnya dialog "repaired content" di
PowerPoint yang biasanya dipicu oleh lxml saat menulis ulang deklarasi
namespace atau gaya tanda kutip.
"""
import re
import io
import zipfile
import openpyxl
from openpyxl.utils.cell import range_boundaries


# ---------------------------------------------------------------------------
# XML text helpers
# ---------------------------------------------------------------------------
def _xml_escape(s):
    """
    Meng-escape karakter khusus XML (&, <, >) pada sebuah nilai apa pun.

    Cara kerja: nilai `s` di-cast ke string lalu tiga karakter tersebut
    diganti berurutan dengan entity XML-nya (& lebih dulu, agar entity yang
    baru disisipkan tidak ikut ter-escape ulang).

    Parameter:
        s: nilai apa pun (biasanya str, tapi bisa juga angka) yang akan
           disisipkan ke dalam teks XML — dipanggil dari `_make_str_pts`
           (untuk label kategori/nama seri) dan dari `update_table_rows`
           (untuk nilai sel tabel native).

    Return: str yang sudah aman disisipkan sebagai isi elemen XML (mis. di
    dalam `<c:v>...</c:v>` atau `<a:t>...</a:t>`).
    """
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;'))


def _make_str_pts(labels):
    """
    Membentuk rangkaian elemen `<c:pt idx="i"><c:v>...</c:v></c:pt>` untuk
    cache string (dipakai di dalam `<c:strCache>`), satu elemen per label.

    Cara kerja: melakukan enumerate atas `labels`, melewati (skip) entri
    yang bernilai None (idx-nya ikut dilewati juga, sesuai perilaku asli),
    dan meng-escape tiap label lewat `_xml_escape` sebelum disisipkan.

    Parameter:
        labels: list nama seri (dipanggil dari `_new_str_cache` saat
                membangun `<c:tx>`) atau list nama kategori (dipanggil dari
                `_new_str_cache` saat membangun `<c:cat>`). Nilainya berasal
                dari `series_list`/`categories` yang diteruskan
                `pptx_updater.py` ke `update_chart_xml`.

    Return: satu string berisi gabungan semua elemen `<c:pt>`, digunakan
    langsung sebagai isi `<c:strCache>` oleh `_new_str_cache`.
    """
    return ''.join(
        f'<c:pt idx="{i}"><c:v>{_xml_escape(l)}</c:v></c:pt>'
        for i, l in enumerate(labels) if l is not None
    )


def _make_num_pts(values):
    """
    Membentuk rangkaian elemen `<c:pt idx="i"><c:v>...</c:v></c:pt>` untuk
    cache angka (dipakai di dalam `<c:numCache>`), satu elemen per nilai.

    Cara kerja: melakukan enumerate atas `values`; entri None dilewati, dan
    entri yang gagal dikonversi ke `float` (TypeError/ValueError) juga
    dilewati secara diam-diam (silent skip) agar satu nilai rusak tidak
    menggagalkan seluruh chart.

    Parameter:
        values: list nilai numerik satu seri data (mis. `values` dari tuple
                `(ser_name, values)` di `series_list`), dipanggil dari
                `_new_num_cache` yang pada gilirannya dipanggil oleh
                `_patch_val`/`_patch_xval`/`_patch_yval`/`_patch_cat`.

    Return: satu string berisi gabungan semua elemen `<c:pt>`, digunakan
    langsung sebagai isi `<c:numCache>` oleh `_new_num_cache`.
    """
    parts = []
    for i, v in enumerate(values):
        if v is None:
            continue
        try:
            parts.append(f'<c:pt idx="{i}"><c:v>{float(v)}</c:v></c:pt>')
        except (TypeError, ValueError):
            pass
    return ''.join(parts)


# ---------------------------------------------------------------------------
# Pola regex yang sudah di-compile (prefix namespace c: dipakai di semua
# file XML chart OOXML)
# ---------------------------------------------------------------------------
_SER_RE    = re.compile(r'<c:ser>(.*?)</c:ser>', re.DOTALL)
_TX_RE     = re.compile(r'(<c:tx\b[^>]*>.*?</c:tx>)', re.DOTALL)
_CAT_RE    = re.compile(r'(<c:cat\b[^>]*>.*?</c:cat>)', re.DOTALL)
_VAL_RE    = re.compile(r'(<c:val\b[^>]*>.*?</c:val>)', re.DOTALL)
_XVAL_RE   = re.compile(r'(<c:xVal\b[^>]*>.*?</c:xVal>)', re.DOTALL)
_YVAL_RE   = re.compile(r'(<c:yVal\b[^>]*>.*?</c:yVal>)', re.DOTALL)
_STR_CACHE = re.compile(r'<c:strCache>.*?</c:strCache>', re.DOTALL)
_NUM_CACHE = re.compile(r'<c:numCache>.*?</c:numCache>', re.DOTALL)
_FMT_CODE  = re.compile(r'<c:formatCode>(.*?)</c:formatCode>')


def _new_str_cache(labels):
    """
    Membangun satu blok `<c:strCache>` baru dari nol (termasuk `<c:ptCount>`
    dan seluruh `<c:pt>`-nya lewat `_make_str_pts`), untuk menggantikan
    blok `<c:strCache>` lama secara utuh.

    Parameter:
        labels: list string (nama seri dari `_patch_tx`, atau nama
                kategori dari `_patch_cat`) yang akan menjadi isi cache.

    Return: string XML lengkap `<c:strCache>...</c:strCache>`, dipakai oleh
    `_patch_tx` dan `_patch_cat` sebagai pengganti (via `_STR_CACHE.sub`)
    dari blok `<c:strCache>` yang ada di XML chart.
    """
    return (
        f'<c:strCache>'
        f'<c:ptCount val="{len(labels)}"/>'
        f'{_make_str_pts(labels)}'
        f'</c:strCache>'
    )


def _new_num_cache(values, original_block=''):
    """
    Membangun satu blok `<c:numCache>` baru dari nol, sambil mempertahankan
    `<c:formatCode>` (format angka, mis. "0.0%") dari blok lama jika ada.

    Cara kerja: mencari `<c:formatCode>` pada `original_block` dengan regex
    `_FMT_CODE`; jika ditemukan, format tersebut dipakai ulang, jika tidak
    dipakai default `'General'`. Elemen `<c:pt>` dibangun lewat
    `_make_num_pts`.

    Parameter:
        values: list nilai numerik yang akan menjadi isi cache (dari
                `_patch_val`/`_patch_xval`/`_patch_yval`, atau dari
                `_patch_cat` untuk kategori numerik pada chart scatter).
        original_block: teks XML blok `<c:val>`/`<c:cat>`/dst yang lama,
                dipakai hanya untuk mengambil `<c:formatCode>`-nya;
                default string kosong berarti selalu pakai format
                `'General'`.

    Return: string XML lengkap `<c:numCache>...</c:numCache>`, dipakai oleh
    para pemanggilnya di atas sebagai pengganti (via `_NUM_CACHE.sub`) dari
    blok `<c:numCache>` yang ada di XML chart.
    """
    fmt_m = _FMT_CODE.search(original_block)
    fmt   = fmt_m.group(1) if fmt_m else 'General'
    return (
        f'<c:numCache>'
        f'<c:formatCode>{fmt}</c:formatCode>'
        f'<c:ptCount val="{len(values)}"/>'
        f'{_make_num_pts(values)}'
        f'</c:numCache>'
    )


# ---------------------------------------------------------------------------
# Patcher per-bagian (beroperasi pada teks satu blok <c:ser>)
# ---------------------------------------------------------------------------
def _patch_tx(ser_text, series_name):
    """
    Mengganti nama seri (strCache di dalam `<c:tx>`) pada satu blok `<c:ser>`.

    Cara kerja: `_TX_RE` mencari blok `<c:tx>...</c:tx>` (non-greedy) di
    dalam `ser_text`; begitu ketemu, `_STR_CACHE` mencari `<c:strCache>` di
    dalam blok `<c:tx>` itu dan menggantinya dengan cache baru hasil
    `_new_str_cache([series_name])` (list berisi satu nama seri saja).
    Substitusi dibatasi `count=1` karena tiap `<c:ser>` hanya punya satu
    `<c:tx>`.

    Parameter:
        ser_text: teks XML satu blok `<c:ser>` (didapat dari `m.group(1)`
                  di dalam `update_chart_xml`).
        series_name: nama seri (str), elemen pertama dari tuple
                     `(ser_name, values)` pada `series_list` yang
                     diteruskan `pptx_updater._update_chart` ke
                     `update_chart_xml`.

    Return: teks `ser_text` dengan `<c:tx>` yang sudah diperbarui,
    diteruskan sebagai `content` ke `_patch_cat` berikutnya di dalam
    `update_chart_xml`.
    """
    def _repl(m):
        tx = m.group(1)
        return _STR_CACHE.sub(_new_str_cache([series_name]), tx, count=1)
    return _TX_RE.sub(_repl, ser_text, count=1)


def _patch_cat(ser_text, categories):
    """
    Mengganti cache kategori (strCache atau numCache di dalam `<c:cat>`)
    pada satu blok `<c:ser>`.

    Cara kerja: `_CAT_RE` mencari blok `<c:cat>...</c:cat>`. Di dalamnya,
    fungsi ini mengecek apakah kategori disimpan sebagai `<c:strCache>`
    (kategori teks, kasus umum) atau `<c:numCache>` (kategori numerik,
    dipakai chart tipe scatter/bubble untuk sumbu kategori). Untuk kasus
    numerik, tiap kategori dicoba di-cast ke `float`; jika ada yang gagal
    (TypeError/ValueError), seluruh daftar kategori diganti fallback berupa
    indeks berurutan `0..len(categories)-1` agar chart tetap valid. Jika
    `<c:cat>` tidak memuat salah satu cache tersebut, blok dikembalikan
    apa adanya (tidak ada yang diubah).

    Parameter:
        ser_text: teks `<c:ser>` yang sudah melalui `_patch_tx` (dipanggil
                  berurutan dari `update_chart_xml`).
        categories: list label kategori (str) atau nilai numerik, berasal
                    dari `categories` yang dikembalikan `get_chart_data`
                    dan diteruskan `pptx_updater._update_chart` ke
                    `update_chart_xml`.

    Return: teks `ser_text` dengan `<c:cat>` yang sudah diperbarui,
    diteruskan sebagai `content` ke `_patch_val` berikutnya di dalam
    `update_chart_xml`.
    """
    def _repl(m):
        cat = m.group(1)
        if '<c:strCache>' in cat:
            return _STR_CACHE.sub(_new_str_cache(categories), cat, count=1)
        if '<c:numCache>' in cat:
            try:
                num_cats = [float(c) for c in categories]
            except (TypeError, ValueError):
                num_cats = list(range(len(categories)))
            return _NUM_CACHE.sub(_new_num_cache(num_cats, cat), cat, count=1)
        return cat
    return _CAT_RE.sub(_repl, ser_text, count=1)


def _patch_val(ser_text, values):
    """
    Mengganti cache nilai (numCache di dalam `<c:val>`) pada satu blok
    `<c:ser>`.

    Cara kerja: `_VAL_RE` mencari blok `<c:val>...</c:val>`; di dalamnya
    `_NUM_CACHE` mencari `<c:numCache>` dan menggantinya dengan hasil
    `_new_num_cache(values, val)` — `val` (isi blok `<c:val>` lama)
    diteruskan agar `<c:formatCode>` aslinya tetap dipertahankan.

    Parameter:
        ser_text: teks `<c:ser>` yang sudah melalui `_patch_tx` dan
                  `_patch_cat` (dipanggil dari `update_chart_xml`).
        values: list nilai numerik satu seri, elemen kedua dari tuple
                `(ser_name, values)` pada `series_list` yang diteruskan
                `pptx_updater._update_chart` ke `update_chart_xml`.

    Return: teks `ser_text` final (setelah `<c:tx>`, `<c:cat>`, `<c:val>`
    semuanya diperbarui), dikembalikan oleh `update_chart_xml.replace_ser`
    sebagai isi baru blok `<c:ser>`.
    """
    def _repl(m):
        val = m.group(1)
        return _NUM_CACHE.sub(_new_num_cache(values, val), val, count=1)
    return _VAL_RE.sub(_repl, ser_text, count=1)


def _patch_xval(ser_text, values):
    """
    Mengganti cache sumbu-X (numCache di dalam `<c:xVal>`) pada satu blok
    `<c:ser>` chart scatter/bubble. Cara kerjanya identik dengan
    `_patch_val`, hanya menyasar elemen `<c:xVal>` lewat `_XVAL_RE`.

    Parameter:
        ser_text: teks satu blok `<c:ser>` (dari
                  `update_chart_xml_scatter.replace_ser`).
        values: list nilai sumbu X, yaitu `x_vals` yang diambil dari
                elemen pertama tiap baris `scatter_rows` di dalam
                `update_chart_xml_scatter`.

    Return: teks `ser_text` dengan `<c:xVal>` yang sudah diperbarui,
    diteruskan ke `_patch_yval` di dalam `update_chart_xml_scatter`.
    """
    def _repl(m):
        xv = m.group(1)
        return _NUM_CACHE.sub(_new_num_cache(values, xv), xv, count=1)
    return _XVAL_RE.sub(_repl, ser_text, count=1)


def _patch_yval(ser_text, values):
    """
    Mengganti cache sumbu-Y (numCache di dalam `<c:yVal>`) pada satu blok
    `<c:ser>` chart scatter/bubble. Cara kerjanya identik dengan
    `_patch_val`/`_patch_xval`, hanya menyasar elemen `<c:yVal>` lewat
    `_YVAL_RE`.

    Parameter:
        ser_text: teks `<c:ser>` yang sudah melalui `_patch_xval` (dipanggil
                  berurutan dari `update_chart_xml_scatter`).
        values: list nilai sumbu Y, yaitu `y_vals` yang diambil dari
                elemen kedua tiap baris `scatter_rows` di dalam
                `update_chart_xml_scatter`.

    Return: teks `ser_text` final, dikembalikan oleh
    `update_chart_xml_scatter.replace_ser` sebagai isi baru blok `<c:ser>`.
    """
    def _repl(m):
        yv = m.group(1)
        return _NUM_CACHE.sub(_new_num_cache(values, yv), yv, count=1)
    return _YVAL_RE.sub(_repl, ser_text, count=1)


# ---------------------------------------------------------------------------
# API publik: pembaruan cache XML chart
# ---------------------------------------------------------------------------
def update_chart_xml(chart_xml_bytes, categories, series_list):
    """
    Memperbarui cache data pada chartN.xml (untuk chart non-scatter, mis.
    bar/line/pie) memakai regex atas teks UTF-8 mentah. File asli
    dikembalikan dengan HANYA isi `<c:strCache>` / `<c:numCache>` yang
    diganti — deklarasi XML, namespace, dan struktur lainnya tidak disentuh
    sama sekali, sehingga PowerPoint tidak menampilkan dialog
    "repaired content".

    Cara kerja: `_SER_RE` mencari tiap blok `<c:ser>...</c:ser>` (satu blok
    = satu seri data pada chart). Untuk setiap blok yang ditemukan
    (`counter` melacak indeks urutan seri), diambil pasangan
    `(ser_name, values)` dari `series_list` sesuai indeks tersebut lalu
    dijalankan berantai: `_patch_tx` (ganti nama seri) →
    `_patch_cat` (ganti label kategori) → `_patch_val` (ganti nilai data).
    Jika jumlah blok `<c:ser>` pada chart melebihi panjang `series_list`,
    blok kelebihan itu dikembalikan tanpa perubahan (`m.group(0)`).

    Parameter:
        chart_xml_bytes: isi mentah file `chartN.xml` dalam bytes, dibaca
                         lewat `editor.read(chart_path)` di
                         `pptx_updater._update_chart`.
        categories: list label kategori (sumbu X / legend kategori),
                    berasal dari `get_chart_data(...)` yang dipanggil
                    di `_update_chart`.
        series_list: list tuple `(nama_seri, list_nilai)`, juga berasal
                     dari hasil `get_chart_data(...)` di `_update_chart`.

    Return: bytes XML chart yang sudah diperbarui. Di `_update_chart`,
    hasil ini disimpan lewat `editor.update(chart_path, new_chart_xml)`
    dan juga diteruskan sebagai argumen `chart_xml_bytes` ke
    `update_embedded_workbook` supaya workbook tertanam disinkronkan
    dengan data yang sama.
    """
    xml = chart_xml_bytes.decode('utf-8')
    counter = [0]

    def replace_ser(m):
        idx = counter[0]
        counter[0] += 1
        if idx >= len(series_list):
            return m.group(0)
        ser_name, values = series_list[idx]
        content = m.group(1)
        content = _patch_tx(content, ser_name)
        content = _patch_cat(content, categories)
        content = _patch_val(content, values)
        return f'<c:ser>{content}</c:ser>'

    return _SER_RE.sub(replace_ser, xml).encode('utf-8')


_VALAX_RE = re.compile(r'<c:valAx>.*?</c:valAx>', re.DOTALL)
_AXIS_MAX_RE = re.compile(r'<c:max val="[^"]*"/>')


def set_value_axis_autoscale(chart_xml_bytes):
    """
    Menghapus batas atas (`<c:max>`) sumbu nilai (value axis) pada chart,
    supaya PowerPoint menghitung ulang skala sumbu secara otomatis
    berdasarkan data yang sesungguhnya tampil, bukan memakai batas atas
    yang di-hardcode dari template (nilai contoh region demo saat
    template dibuat).

    KENAPA DIBUTUHKAN: beberapa chart (mis. chart7 slide 4, breakdown per
    fungsi) di-generate ulang untuk 12 region yang jumlah karyawannya
    beda-beda jauh, tapi `<c:max>` sumbu nilainya tetap nilai tunggal yang
    di-hardcode di template (mis. "1100", pas untuk Jawa Tengah tapi
    kekecilan untuk region berkantong headcount besar) — akibatnya bar
    milik region lain terpotong/keluar dari bingkai plot area. Menghapus
    `<c:max>` (min dibiarkan kalau ada, biasanya "0", supaya sumbu tetap
    mulai dari nol) membuat PowerPoint memakai auto-scaling bawaannya,
    yang selalu pas dengan data yang sedang ditampilkan.

    Cara kerja: `_VALAX_RE` mencari tiap blok `<c:valAx>...</c:valAx>`
    (chart dual-axis punya lebih dari satu), lalu di dalam tiap blok itu
    `_AXIS_MAX_RE` menghapus elemen `<c:max val="..."/>` (kalau ada; kalau
    tidak ada, blok dibiarkan apa adanya).

    Parameter: chart_xml_bytes (bytes XML chart, hasil `update_chart_xml`).

    Return: bytes XML chart dengan `<c:max>` sumbu nilai sudah dihapus.
    Dipanggil dari `_update_chart` di pptx_updater.py untuk nomor chart
    tertentu (saat ini chart 7 dan 10-14 di slide 4) SETELAH
    `update_chart_xml`, sebelum `editor.update(...)`.
    """
    xml = chart_xml_bytes.decode('utf-8')
    xml = _VALAX_RE.sub(lambda m: _AXIS_MAX_RE.sub('', m.group(0)), xml)
    return xml.encode('utf-8')


def update_chart_xml_scatter(chart_xml_bytes, scatter_rows):
    """
    Memperbarui cache `xVal`/`yVal` pada chart tipe scatter/bubble.

    Cara kerja: berbeda dari `update_chart_xml`, di sini hanya ada satu
    seri data (`_SER_RE.sub(..., count=1)` — hanya blok `<c:ser>` pertama
    yang diproses). Koordinat X dan Y diekstrak lebih dulu dari
    `scatter_rows` (elemen indeks 0 dan 1 tiap baris; elemen tambahan pada
    baris, kalau ada, diabaikan), lalu dijalankan berantai `_patch_xval`
    diikuti `_patch_yval` pada isi blok `<c:ser>` tersebut.

    Parameter:
        chart_xml_bytes: isi mentah file `chartN.xml` dalam bytes, dibaca
                         lewat `editor.read(chart_path)` di
                         `pptx_updater._update_chart` untuk chart yang
                         nomornya terdaftar di `SCATTER_CHARTS`.
        scatter_rows: list tuple/list `(x, y, ...)`, yaitu `chart_result`
                      yang dikembalikan `get_chart_data(...)` untuk chart
                      scatter — diteruskan apa adanya oleh `_update_chart`.

    Return: bytes XML chart yang sudah diperbarui. Di `_update_chart`,
    hasil ini disimpan lewat `editor.update(chart_path, new_chart_xml)`.
    Catatan: untuk chart scatter, workbook tertanam TIDAK ikut
    disinkronkan (lihat percabangan `if embed_path and chart_num not in
    SCATTER_CHARTS` di `_update_chart`), jadi hasil fungsi ini tidak
    diteruskan ke `update_embedded_workbook`.
    """
    xml = chart_xml_bytes.decode('utf-8')
    x_vals = [r[0] for r in scatter_rows]
    y_vals = [r[1] for r in scatter_rows]

    def replace_ser(m):
        content = m.group(1)
        content = _patch_xval(content, x_vals)
        content = _patch_yval(content, y_vals)
        return f'<c:ser>{content}</c:ser>'

    return _SER_RE.sub(replace_ser, xml, count=1).encode('utf-8')


# ---------------------------------------------------------------------------
# Pembaruan workbook Excel tertanam (embedded)
#
# Cache XML chart (yang diperbarui di atas) adalah yang dirender PowerPoint,
# tetapi setiap chart juga membawa workbook .xlsx tertanam yang menjadi
# fallback PowerPoint setiap kali chart disinkronkan ulang dari sumber
# datanya — misalnya saat pengguna klik "Edit Data", atau saat PowerPoint
# auto-refresh ketika file dibuka/dihitung ulang. Jika workbook itu masih
# menyimpan nilai placeholder dari template, chart akan kembali menampilkan
# data placeholder begitu hal itu terjadi. Karena itu sel-sel workbook
# tertanam harus disinkronkan dengan kategori/seri yang sama dengan yang
# ditulis ke cache XML chart.
# ---------------------------------------------------------------------------
_SER_TX_F  = re.compile(r'<c:tx>\s*<c:strRef>\s*<c:f>([^<]*)</c:f>', re.DOTALL)
_SER_CAT_F = re.compile(r'<c:cat\b[^>]*>.*?<c:f>([^<]*)</c:f>', re.DOTALL)
_SER_VAL_F = re.compile(r'<c:val\b[^>]*>.*?<c:f>([^<]*)</c:f>', re.DOTALL)


def _resolve_sheet(wb, ref):
    """
    Mengembalikan worksheet yang disebutkan dalam referensi sel bergaya
    `'Sheet1!$A$1'` (fallback: sheet yang sedang aktif jika tidak ada nama
    sheet, atau nama sheet-nya tidak ditemukan di workbook).

    Cara kerja: memeriksa apakah `ref` mengandung `!` (pemisah nama sheet
    dari alamat sel); jika ya, nama sheet diambil dari bagian sebelum `!`
    (tanda kutip tunggal di sekelilingnya dilucuti dengan `strip("'")`)
    dan dicari di `wb.sheetnames`.

    Parameter:
        wb: objek `openpyxl.Workbook` hasil `openpyxl.load_workbook` di
            `update_embedded_workbook`.
        ref: string referensi sel/range lengkap dengan nama sheet, mis.
             hasil `tx_m.group(1)` / `cat_m.group(1)` / `val_m.group(1)`
             dari `update_embedded_workbook` (diambil dari elemen `<c:f>`
             pada XML chart).

    Return: objek worksheet (`openpyxl.Worksheet`), dipakai oleh
    `_write_cell` dan `_write_range` untuk menulis nilai ke sel yang tepat.
    """
    if '!' in ref:
        name = ref.split('!', 1)[0].strip("'")
        if name in wb.sheetnames:
            return wb[name]
    return wb.active


def _sync_table_header(ws, row, col, value):
    """
    Jika sel (row, col) adalah sel header dari sebuah Excel Table
    (ListObject), ganti nama `tableColumn` yang bersesuaian agar tetap
    sinkron dengan teks sel. Excel/PowerPoint mengharuskan sel header
    sebuah table persis sama dengan atribut `name` pada `tableColumn`-nya —
    ketidakcocokan ini merusak workbook tertanam (muncul di PowerPoint
    sebagai sumber data chart yang rusak / pesan "linked file isn't
    available").

    Cara kerja: mengiterasi semua `ws.tables` (Excel Table pada sheet
    tersebut); untuk tiap table, batas kolom/baris diambil lewat
    `range_boundaries(table.ref)`. Table dilewati jika baris header table
    itu (`min_row`) bukan `row` yang sedang ditulis, atau jika `col` di
    luar rentang kolom table, atau jika table itu tidak punya baris header
    sama sekali (`headerRowCount` 0). Jika cocok, `tableColumns[col -
    min_col].name` diganti dengan `str(value)`.

    Parameter:
        ws: worksheet target, hasil `_resolve_sheet` di `_write_cell`.
        row, col: koordinat sel (1-based) yang baru saja ditulis, dari
                  `_write_cell` (yaitu `min_row`/`min_col` hasil
                  `range_boundaries` atas referensi sel tunggal).
        value: nilai baru sel tersebut, diteruskan apa adanya dari
               `_write_cell` (biasanya nama seri, dari
               `update_embedded_workbook`).

    Return: None — fungsi ini hanya melakukan efek samping (side effect)
    mengubah nama kolom table pada objek `ws` secara in-place.
    """
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row != row or not (min_col <= col <= max_col):
            continue
        if not table.headerRowCount:
            continue
        table.tableColumns[col - min_col].name = str(value)


def _write_cell(wb, ref, value):
    """
    Menulis satu nilai ke satu sel workbook, ditunjuk oleh referensi
    gaya `'Sheet1!$A$1'`, dan menjaga sinkron nama kolom table jika sel itu
    kebetulan header table.

    Cara kerja: sheet-nya diselesaikan lewat `_resolve_sheet`; tanda `$`
    pada bagian alamat sel dibuang lalu diurai jadi koordinat kolom/baris
    lewat `range_boundaries` (dipakai meski cuma satu sel, karena fungsi
    itu menerima gaya alamat range). Nilai ditulis dengan `ws.cell(...)`,
    lalu `_sync_table_header` dipanggil untuk berjaga-jaga jika sel itu
    memang header sebuah Excel Table.

    Parameter:
        wb: `openpyxl.Workbook`, diteruskan dari `update_embedded_workbook`.
        ref: referensi sel tunggal lengkap dengan nama sheet, mis. dari
             `tx_m.group(1)` (referensi `<c:f>` di dalam `<c:tx>`) pada
             `update_embedded_workbook` — biasanya menunjuk sel judul seri.
        value: nilai baru yang ditulis; dipanggil dengan `ser_name`
               (nama seri) dari `update_embedded_workbook`.

    Return: None — efek sampingnya menulis langsung ke objek `wb` yang
    kemudian di-save oleh `update_embedded_workbook`.
    """
    ws = _resolve_sheet(wb, ref)
    cell_ref = ref.split('!', 1)[-1].replace('$', '')
    min_col, min_row, _, _ = range_boundaries(cell_ref)
    ws.cell(row=min_row, column=min_col, value=value)
    _sync_table_header(ws, min_row, min_col, value)


def _write_range(wb, ref, values):
    """
    Menulis serangkaian nilai ke satu range sel (satu baris atau satu
    kolom), ditunjuk oleh referensi gaya `'Sheet1!$A$1:$A$5'`.

    Cara kerja: sheet-nya diselesaikan lewat `_resolve_sheet`, lalu batas
    range diurai dengan `range_boundaries`. Arah penulisan ditentukan
    otomatis: jika `max_row > min_row` maka range dianggap vertikal (nilai
    ditulis menurun per baris), selain itu dianggap horizontal (nilai
    ditulis menyamping per kolom). Iterasi berhenti lebih awal (`break`)
    begitu indeks melewati batas range asli, sehingga jumlah `values` yang
    lebih panjang dari range template tidak meluber ke sel lain.

    Parameter:
        wb: `openpyxl.Workbook`, diteruskan dari `update_embedded_workbook`.
        ref: referensi range lengkap dengan nama sheet, mis. dari
             `cat_m.group(1)` (referensi `<c:f>` di dalam `<c:cat>`) atau
             `val_m.group(1)` (di dalam `<c:val>`) pada
             `update_embedded_workbook`.
        values: list nilai yang ditulis berurutan ke range tersebut;
                dipanggil dengan `categories` atau dengan `values` (nilai
                seri) dari `update_embedded_workbook`.

    Return: None — efek sampingnya menulis langsung ke objek `wb` yang
    kemudian di-save oleh `update_embedded_workbook`.
    """
    ws = _resolve_sheet(wb, ref)
    range_ref = ref.split('!', 1)[-1].replace('$', '')
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    vertical = max_row > min_row
    for i, v in enumerate(values):
        if vertical:
            row, col = min_row + i, min_col
            if row > max_row:
                break
        else:
            row, col = min_row, min_col + i
            if col > max_col:
                break
        ws.cell(row=row, column=col, value=v)


def update_embedded_workbook(xlsx_bytes, chart_xml_bytes, categories, series_list):
    """
    Menulis kategori/data seri yang sama ke dalam workbook .xlsx tertanam
    milik chart, agar isinya cocok dengan cache XML chart yang sudah
    diperbarui `update_chart_xml`. Hanya untuk chart non-scatter (gaya
    bar/line/pie) — tiap seri pada chart tipe ini membawa triple referensi
    sel `<c:tx>`/`<c:cat>`/`<c:val>` yang di-mirror-kan ke sel-sel workbook.

    Cara kerja: `chart_xml_bytes` (XML chart yang SUDAH diperbarui — lihat
    catatan return di `update_chart_xml`) di-decode, lalu `_SER_RE`
    dipakai lagi untuk mengambil ulang tiap blok `<c:ser>` (kali ini hanya
    untuk membaca referensi sel-nya, bukan mengubah cache). Untuk tiap
    blok, sesuai indeksnya, diambil `(ser_name, values)` dari
    `series_list`, lalu:
      - `_SER_TX_F` mencari referensi sel judul seri di dalam `<c:tx>`
        dan menuliskannya lewat `_write_cell(wb, ref, ser_name)`;
      - `_SER_CAT_F` mencari referensi range kategori di dalam `<c:cat>`
        dan menuliskannya lewat `_write_range(wb, ref, categories)`;
      - `_SER_VAL_F` mencari referensi range nilai di dalam `<c:val>` dan
        menuliskannya lewat `_write_range(wb, ref, values)`.
    Jika salah satu pola `<c:f>` tidak ditemukan pada suatu seri, bagian
    itu dilewati saja (tidak dianggap error). Loop berhenti jika indeks
    seri sudah melebihi panjang `series_list`.

    Parameter:
        xlsx_bytes: isi mentah workbook .xlsx tertanam, dibaca lewat
                    `editor.read(embed_path)` di `pptx_updater._update_chart`.
        chart_xml_bytes: XML chart yang SUDAH diperbarui, yaitu
                         `new_chart_xml` (hasil `update_chart_xml`) yang
                         diteruskan `_update_chart` — bukan XML chart yang
                         lama/asli.
        categories, series_list: sama seperti parameter yang diterima
                                  `update_chart_xml`, diteruskan apa
                                  adanya oleh `_update_chart`.

    Return: bytes .xlsx hasil `wb.save(buf)` ke buffer in-memory. Di
    `_update_chart`, hasil ini disimpan kembali ke editor lewat
    `editor.update(embed_path, new_embed_bytes)`.
    """
    xml = chart_xml_bytes.decode('utf-8')
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

    for idx, ser_text in enumerate(_SER_RE.findall(xml)):
        if idx >= len(series_list):
            break
        ser_name, values = series_list[idx]

        tx_m = _SER_TX_F.search(ser_text)
        if tx_m:
            _write_cell(wb, tx_m.group(1), ser_name)

        cat_m = _SER_CAT_F.search(ser_text)
        if cat_m:
            _write_range(wb, cat_m.group(1), categories)

        val_m = _SER_VAL_F.search(ser_text)
        if val_m:
            _write_range(wb, val_m.group(1), values)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Pembaruan tabel native PowerPoint (mis. tabel training BSC di slide 10)
# ---------------------------------------------------------------------------
_TR_RE = re.compile(r'<a:tr\b.*?</a:tr>', re.DOTALL)
_TABLE_T_RE = re.compile(r'(<a:t>)[^<]*(</a:t>)')


def update_table_rows(slide_xml_bytes, row_key_fn, row_values_fn):
    """
    Mengisi baris-baris data sebuah tabel native `<a:tbl>` tanpa menyentuh
    struktur/format-nya (jumlah baris/kolom, styling run, dsb tetap sama —
    hanya teks di dalam `<a:t>` yang mungkin berubah).

    Cara kerja: `_TR_RE` mencari tiap blok `<a:tr>...</a:tr>` (satu baris
    tabel). Untuk tiap baris, semua teks `<a:t>` di dalamnya diekstrak
    berurutan (`cur_texts`) lalu diserahkan ke `row_key_fn(cur_texts)` agar
    si pemanggil bisa menentukan baris data mana (jika ada) yang cocok
    dengan baris ini — daftar teks LENGKAP diserahkan, bukan cuma teks
    pertama, karena satu label kadang terpecah jadi beberapa run, mis.
    "HO/" + "Sentralisasi". Jika `row_key_fn` mengembalikan None, baris itu
    dibiarkan apa adanya (bukan baris data, mis. baris header). Jika
    mengembalikan sebuah key, `row_values_fn(key, cur_texts)` dipanggil dan
    harus mengembalikan list string pengganti sepanjang `len(cur_texts)`
    (elemen bernilai None berarti sel itu tidak diubah); jika panjangnya
    tidak cocok atau hasilnya None, baris dikembalikan tanpa perubahan.
    Substitusi teks per run dilakukan oleh `_TABLE_T_RE`, dengan tiap nilai
    pengganti di-escape lewat `_xml_escape` sebelum disisipkan.

    Parameter:
        slide_xml_bytes: XML slide (atau, dari pemanggil di baris 617,
                         satu bagian teks slide dalam bytes) yang memuat
                         tabel target.
        row_key_fn: fungsi `(texts) -> key|None` yang disuplai si pemanggil
                    untuk mengenali baris; contoh nyata: `_slide10_row_key`
                    (dipanggil di `pptx_updater.py` sekitar baris 451) atau
                    closure `row_key_fn` yang didefinisikan langsung di
                    tempat pemanggilan (mis. sekitar baris 597).
        row_values_fn: fungsi `(key, texts) -> list[str|None]|None` yang
                       disuplai si pemanggil untuk menghasilkan nilai baru
                       tiap sel pada baris yang cocok.

    Return: bytes XML slide/bagian teks yang sudah diperbarui, langsung
    dipakai pemanggil untuk `editor.update(...)` (lihat pemanggilnya di
    `pptx_updater.py`, mis. hasil dari fungsi yang membungkus pemanggilan
    ini pada baris 451, atau langsung di baris 617).
    """
    xml = slide_xml_bytes.decode('utf-8')

    def repl_row(m):
        row_xml = m.group(0)
        cur_texts = re.findall(r'<a:t>([^<]*)</a:t>', row_xml)
        if not cur_texts:
            return row_xml
        key = row_key_fn(cur_texts)
        if key is None:
            return row_xml
        new_values = row_values_fn(key, cur_texts)
        if new_values is None or len(new_values) != len(cur_texts):
            return row_xml

        counter = [0]

        def repl_t(mm):
            i = counter[0]
            counter[0] += 1
            new_v = new_values[i]
            if new_v is None:
                return mm.group(0)
            return f'{mm.group(1)}{_xml_escape(new_v)}{mm.group(2)}'

        return _TABLE_T_RE.sub(repl_t, row_xml)

    return _TR_RE.sub(repl_row, xml).encode('utf-8')


# ---------------------------------------------------------------------------
# Pembaruan teks slide – murni string replacement, tanpa parse/reserialize XML
# ---------------------------------------------------------------------------
def replace_text_in_slide(slide_xml_bytes, replacements):
    """
    Menerapkan sederet penggantian string (old, new) ke bytes XML slide
    tanpa parsing XML sama sekali.

    Cara kerja: XML slide di-decode ke string, lalu untuk tiap pasangan
    `(old, new)` di `replacements`, dilakukan `text.replace(old, new)`
    (hanya jika `old` memang ada di teks, sebagai optimisasi kecil). Ini
    aman dilakukan sebagai string replace polos (bukan regex/parse) karena
    nilai yang diganti (mis. nama region) hanya pernah muncul sebagai isi
    teks di dalam `<a:t>...</a:t>`, tidak pernah sebagai bagian dari nama
    atribut atau nama elemen XML — sehingga tidak berisiko merusak
    struktur tag.

    Parameter:
        slide_xml_bytes: bytes XML satu slide, dibaca lewat
                         `editor.read(slide_path)` di
                         `pptx_updater.generate_pptx_for_region` lalu
                         diteruskan lewat wrapper lokal
                         `_replace_text_in_slide`.
        replacements: list tuple `(old, new)` string, dibangun oleh
                      `_build_slide_replacements(slide_idx, data, region)`
                      di `pptx_updater.py`.

    Return: bytes XML slide yang sudah diperbarui, disimpan kembali lewat
    `editor.update(slide_path, updated)` di
    `pptx_updater.generate_pptx_for_region`.
    """
    text = slide_xml_bytes.decode('utf-8')
    for old, new in replacements:
        if old in text:
            text = text.replace(old, new)
    return text.encode('utf-8')


# ---------------------------------------------------------------------------
# Operasi tingkat zip untuk PPTX
# ---------------------------------------------------------------------------
class PptxEditor:
    """
    Editor PPTX in-memory; mempertahankan tipe kompresi asli tiap file di
    dalam arsip zip (bukan menekan-ulang semuanya dengan satu default).

    Dipakai sebagai context manager (`with PptxEditor(TEMPLATE_PATH) as
    editor:`) di `pptx_updater.generate_pptx_for_region` — file template
    dibaca satu kali di `__enter__`, seluruh entri chart/slide/workbook
    dibaca dan ditulis ulang lewat `read`/`update` selama blok `with`
    berjalan, lalu hasil akhirnya ditulis ke file output dengan `save`.
    """

    def __init__(self, pptx_path):
        """
        Menyimpan path file PPTX sumber dan menyiapkan state kosong; belum
        membuka file apa pun di sini (pembukaan sebenarnya terjadi di
        `__enter__`).

        Parameter:
            pptx_path: path file .pptx template, diteruskan sebagai
                       `TEMPLATE_PATH` dari
                       `pptx_updater.generate_pptx_for_region`.

        Return: None (constructor).
        """
        self.pptx_path = pptx_path
        self._src_buf  = None
        self._src_zip  = None
        self._updates  = {}   # path → bytes

    def __enter__(self):
        """
        Membuka file PPTX sumber: seluruh isinya dibaca ke memori
        (`self._src_buf`) lalu dibungkus sebagai `zipfile.ZipFile` yang
        siap dibaca (`self._src_zip`), sehingga tidak perlu membuka file
        dari disk berulang kali selama proses edit berlangsung.

        Return: `self` (objek `PptxEditor` itu sendiri), sesuai konvensi
        context manager — inilah nilai yang diikat ke variabel `editor`
        pada `with PptxEditor(...) as editor:`.
        """
        with open(self.pptx_path, 'rb') as f:
            self._src_buf = f.read()
        self._src_zip = zipfile.ZipFile(io.BytesIO(self._src_buf), 'r')
        return self

    def read(self, path):
        """
        Membaca isi satu entri di dalam arsip PPTX berdasarkan path
        internalnya (mis. `'ppt/slides/slide3.xml'`).

        Cara kerja: jika path tersebut sudah pernah ditulis lewat
        `update()` sebelumnya di sesi edit ini, versi yang sudah diperbarui
        itulah yang dikembalikan (bukan versi asli dari zip) — ini
        memastikan pemanggil selalu melihat versi terbaru, termasuk saat
        satu file yang sama diedit bertahap oleh beberapa langkah berturut
        (lihat pola `slide_xml = editor.read(...)` diikuti beberapa fungsi
        `_apply_...` yang saling menumpuk sebelum `editor.update(...)`
        dipanggil, di `generate_pptx_for_region`). Jika belum pernah
        diperbarui, dibaca langsung dari `self._src_zip`.

        Parameter:
            path: path internal entri zip (str), dipanggil dari berbagai
                  tempat di `pptx_updater.py`, mis. `editor.read(chart_path)`
                  di `_update_chart`, atau `editor.read(slide14_path)` di
                  `generate_pptx_for_region`.

        Return: bytes isi entri tersebut, biasanya langsung diteruskan ke
        salah satu fungsi update di modul ini (`update_chart_xml`,
        `update_table_rows`, `replace_text_in_slide`, dst).
        """
        if path in self._updates:
            return self._updates[path]
        return self._src_zip.read(path)

    def update(self, path, data):
        """
        Menyimpan (di memori) konten baru untuk satu entri arsip, untuk
        ditulis ke file output nanti saat `save()` dipanggil.

        Cara kerja: `data` dinormalisasi ke bytes (di-encode ke UTF-8 jika
        yang diberikan berupa str) lalu disimpan ke dict `self._updates`
        dengan `path` sebagai kunci. Menimpa entri yang sama untuk path
        yang sama tidak masalah — hanya nilai terakhir yang dipakai.

        Parameter:
            path: path internal entri zip yang akan diganti isinya.
            data: bytes atau str konten baru; dipanggil misalnya dengan
                  `new_chart_xml` (dari `update_chart_xml`), `new_embed_bytes`
                  (dari `update_embedded_workbook`), atau `updated` (dari
                  `replace_text_in_slide`/`update_table_rows`) di
                  `pptx_updater.py`.

        Return: None — efek sampingnya mengubah `self._updates` secara
        in-place; perubahan baru benar-benar ditulis ke file saat `save()`
        dipanggil.
        """
        self._updates[path] = data if isinstance(data, (bytes, bytearray)) else data.encode('utf-8')

    def names(self):
        """
        Mengembalikan daftar seluruh path entri di dalam arsip PPTX asli.

        Cara kerja: mendelegasikan langsung ke `self._src_zip.namelist()`
        (daftar ini berasal dari struktur zip asli, tidak dipengaruhi oleh
        entri yang sudah di-`update()` karena `update()` hanya mengganti
        isi, bukan menambah/menghapus entri).

        Return: list path (str), dipakai misalnya oleh
        `pptx_updater._discover_charts` (untuk menemukan pasangan file
        chart/embed) dan `pptx_updater._get_slide_paths` (untuk menyaring
        path yang cocok pola `ppt/slides/slideN.xml`).
        """
        return self._src_zip.namelist()

    def save(self, out_path):
        """
        Menulis arsip PPTX final ke `out_path`: setiap entri dari zip asli
        disalin apa adanya, KECUALI entri yang sudah pernah di-`update()`,
        yang datanya diganti dengan versi baru dari `self._updates`.

        Cara kerja: membuka ulang `self._src_buf` sebagai `ZipFile` sumber
        (`src`) dan membuat `ZipFile` tujuan baru (`dst`). Untuk tiap
        `info` (metadata entri) di `src.infolist()`, data yang ditulis
        adalah `self._updates.get(info.filename, src.read(info.filename))`
        — pakai versi baru jika ada, kalau tidak pakai isi asli.
        `zipfile.ZipInfo` baru dibuat dengan nama dan timestamp yang sama
        dengan entri asli, dan `compress_type`-nya (STORED atau DEFLATED)
        sengaja disalin dari entri asli — bukan dibiarkan default — karena
        PowerPoint/OOXML mengharapkan pola kompresi tertentu per jenis file
        dalam paket, dan menyamakannya membantu menghindari file output
        yang dianggap korup.

        Parameter:
            out_path: path file .pptx output, dipanggil dengan `out_path`
                      (`HR_Dashboard_<region>.pptx`) dari
                      `pptx_updater.generate_pptx_for_region`.

        Return: None — efek sampingnya menulis file baru ke `out_path`.
        """
        with zipfile.ZipFile(io.BytesIO(self._src_buf), 'r') as src:
            with zipfile.ZipFile(out_path, 'w') as dst:
                for info in src.infolist():
                    data = self._updates.get(info.filename, src.read(info.filename))
                    # Pertahankan tipe kompresi asli (STORED atau DEFLATED)
                    new_info = zipfile.ZipInfo(info.filename, info.date_time)
                    new_info.compress_type = info.compress_type
                    dst.writestr(new_info, data)

    def __exit__(self, *args):
        """
        Menutup handle `ZipFile` sumber saat blok `with` berakhir (normal
        maupun karena exception), agar file descriptor/buffer terkait
        tidak menggantung.

        Parameter:
            *args: informasi exception standar dari protokol context
                   manager (tipe, nilai, traceback) — tidak dipakai/
                   diperiksa di sini, sehingga exception apa pun tetap
                   diteruskan (tidak ditelan) setelah cleanup ini.

        Return: None (implisit) — karena tidak mengembalikan True,
        exception yang terjadi di dalam blok `with` (jika ada) tetap akan
        di-raise ulang ke pemanggil.
        """
        if self._src_zip:
            self._src_zip.close()
