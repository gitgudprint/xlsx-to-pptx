"""
Mengorkestrasi seluruh pembaruan PPTX per region:
 - Workbook Excel yang ter-embed (data chart)
 - Cache XML chart
 - Teks slide (judul, kotak anotasi)
 - Tabel native (`<a:tbl>`) dan shape highlight (rect/segitiga/indikator)

Modul ini adalah "otak" orkestrasi tipis: logika per-slide sendiri hidup di
`src/slides/slideN.py` (satu modul per slide) dan utilitas yang dipakai
lintas slide hidup di `src/pptx_common.py` — file ini hanya memanggil
fungsi `_apply_slideNN_*` dari modul-modul itu secara berurutan untuk satu
region, dibungkus `_safe_apply_slide` supaya kegagalan di satu slide tidak
menghentikan slide lain (lihat docstring-nya).

Titik masuk (entry point) modul ini adalah `generate_pptx_for_region()` di
bagian paling bawah file, dipanggil sekali per region dari `main.py`.
"""
import re
import os
import time as _time

from .config import (
    TEMPLATE_PATH, OUTPUT_DIR, TEMPLATE_REGION, TEMPLATE_TOTAL_HC,
)
from .xml_updater import replace_text_in_slide, PptxEditor
from .chart_data import get_slide4_total_hc
from .pptx_common import _discover_charts, _update_chart, _get_slide_paths, _replace_text_in_slide
from .perf import metrics as _perf_metrics

from .slides.slide3 import (
    _compute_slide3_rect_positions, _apply_slide3_rect_updates,
    _compute_slide3_indicator_updates, _apply_slide3_indicator_updates,
    _apply_slide3_highlights,
)
from .slides.slide4 import _apply_slide4_lea_tables, _apply_slide4_highlights
from .slides.slide5 import _apply_slide5_pct_boxes, _apply_slide5_highlights
from .slides.slide6_7 import _apply_slide6_highlights, _apply_slide7_highlights
from .slides.slide10 import _apply_slide10_table, _apply_slide10_rect_highlight, _apply_slide10_insight
from .slides.slide14 import (
    _apply_slide14_tables, _apply_slide14_top3_highlights,
    _apply_slide14_rect_highlights, _apply_slide14_reason_table, _apply_slide14_highlights,
)
from .slides.slide15 import _apply_slide15_tables, _apply_slide15_indicators, _apply_slide15_highlights
from .slides.slide16 import _apply_slide16_branch_table, _apply_slide16_reason_table, _apply_slide16_highlight
from .slides.slide18 import _apply_slide18_summary, _apply_slide18_worst5_table, _apply_slide18_highlights

# Peta nomor slide (1-indexed, sesuai penomoran file "slideN.xml" — sudah
# diverifikasi cocok 1:1 dengan urutan tampil di `ppt/presentation.xml`)
# ke rentang nomor chart yang dikandungnya, sesuai komentar bagian di
# `chart_data.py` ("SLIDE N (chart A-B)"). Dipakai oleh
# `generate_pptx_for_region(..., slides=...)` untuk tahu chart mana saja
# yang perlu diperbarui saat hanya sebagian slide yang diminta — slide
# 10/14/15/16 sengaja tidak punya entri di sini karena isinya murni
# tabel/shape (tidak ada chart PowerPoint di slide-slide itu).
SLIDE_CHART_RANGES = {
    3: range(1, 7),      # chart 1-6
    4: range(7, 18),     # chart 7-17
    5: range(18, 24),    # chart 18-23
    6: range(24, 33),    # chart 24-32
    7: range(33, 42),    # chart 33-41
    12: range(42, 44),   # chart 42-43
    18: range(44, 47),   # chart 44-46
}


def _slide_and_chart_filters(slides):
    """
    Menerjemahkan parameter `slides` publik (`generate_pptx_for_region`)
    jadi 2 nilai internal yang gampang dicek berulang kali:
      - `wanted_slides`: `None` kalau `slides` adalah `None`/kosong
        (artinya "semua slide", perilaku baku/lama), atau `set[int]` nomor
        slide 1-indexed yang diminta.
      - `wanted_charts`: `None` kalau semua chart diproses, atau
        `set[int]` nomor chart yang perlu diperbarui (gabungan
        `SLIDE_CHART_RANGES` untuk tiap nomor di `wanted_slides` — slide
        yang tidak punya chart, mis. 10/14/15/16, otomatis menyumbang
        himpunan kosong lewat `.get(s, ())`).

    Parameter: `slides` — `None`, atau iterable of int (nomor slide
    1-indexed, mis. `{3, 14, 16}` atau `[3, 14, 16]`).

    Return: tuple `(wanted_slides, wanted_charts)`. Dipanggil sekali di
    awal `generate_pptx_for_region`, hasilnya dipakai untuk melewati
    (skip) chart/slide yang tidak diminta di seluruh langkah berikutnya.
    """
    if not slides:
        return None, None
    wanted_slides = set(slides)
    wanted_charts = {c for s in wanted_slides for c in SLIDE_CHART_RANGES.get(s, ())}
    return wanted_slides, wanted_charts


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
def _safe_apply_slide(editor, label, slide_path, apply_fn):
    """
    Menjalankan `apply_fn(slide_xml) -> slide_xml_baru` untuk satu slide
    dengan aman: kalau `apply_fn` melempar exception APA PUN (data yang
    diharapkan kosong/tidak ada, placeholder/token yang diharapkan tidak
    ditemukan di template, dst.), errornya DICATAT lewat `print` (diberi
    prefix "[WARN]" + `label` supaya jelas slide/langkah mana yang
    gagal), TAPI TIDAK menghentikan seluruh proses `generate_pptx_for_region`
    — slide itu dibiarkan dalam kondisi sebelum `apply_fn` dipanggil (baik
    itu masih asli dari template, atau hasil langkah-langkah sebelumnya
    yang berhasil), dan region tetap lanjut diproses ke slide/langkah
    berikutnya serta tetap menghasilkan file PPTX di akhir.

    KENAPA DIBUTUHKAN: sebelum ada helper ini, satu-satunya bagian
    `generate_pptx_for_region` yang punya penanganan galat semacam ini
    hanya loop update chart (langkah 2) — SEMUA pemanggilan
    `_apply_slideNN_*` lainnya (tabel, highlight, dst. — langkah 3
    dst.) TIDAK dibungkus apa pun, jadi satu error tak terduga di
    mana pun (mis. data region kosong untuk satu slide tertentu, atau
    template yang sedikit berubah sehingga sebuah token/placeholder yang
    dicari tidak ketemu) akan menghentikan TOTAL pemrosesan region itu —
    tidak ada file PPTX apa pun yang dihasilkan untuknya, padahal
    slide-slide lain semestinya tetap bisa berhasil diisi dengan data
    yang memang tersedia.

    Parameter:
        editor: instance `PptxEditor` (dipakai untuk `editor.read(slide_path)`
                sebelum memanggil `apply_fn`, dan `editor.update(...)`
                sesudahnya kalau berhasil).
        label: str, nama pendek slide/langkah untuk pesan warning (mis.
               "slide14", "slide16 branch_sales").
        slide_path: str, path XML slide di dalam paket (mis.
                    "ppt/slides/slide14.xml").
        apply_fn: fungsi `(slide_xml: bytes) -> bytes` — biasanya closure
                  kecil yang merantai beberapa `_apply_slideNN_*` untuk
                  satu slide (lihat pemanggil di `generate_pptx_for_region`).

    Return: None. Efeknya berupa `editor.update(slide_path, ...)` kalau
    `apply_fn` berhasil, atau `print` warning kalau gagal (slide tidak
    diubah sama sekali untuk pemanggilan ini). Dipanggil dari
    `generate_pptx_for_region` untuk setiap blok slide (langkah 3a dst.).
    Setiap panggilan juga dicatat ke `perf.metrics` (`record_slide`) untuk
    laporan performa (lihat `src/perf.py`).
    """
    try:
        slide_xml = editor.read(slide_path)
        slide_xml = apply_fn(slide_xml)
        editor.update(slide_path, slide_xml)
        _perf_metrics.record_slide(True)
    except Exception as e:
        print(f"    [WARN] {label}: {e}")
        _perf_metrics.record_slide(False, f"{label}: {type(e).__name__}")


def generate_pptx_for_region(region, data, slides=None, verbose=True):
    """
    TITIK MASUK (entry point) modul ini: menghasilkan satu file PPTX HR
    Dashboard untuk satu `region`, dipanggil dari `main.py`
    (`generate_pptx_for_region(region, data, slides=None, verbose=True)`)
    sekali per region dalam sebuah loop.

    Cara kerja keseluruhan: membuka `TEMPLATE_PATH` (dari `config.py`) lewat
    `PptxEditor` (context manager di `xml_updater.py` yang membaca seluruh
    isi zip PPTX ke memori dan menampung perubahan sebagai
    `{path: bytes baru}` sampai `editor.save()` dipanggil), lalu menjalankan
    10 langkah berurutan yang masing-masing memanggil satu atau beberapa
    fungsi `_apply_slideNN_*`/`_compute_slideNN_*` yang didefinisikan di
    bagian atas file ini. Urutan antar-langkah TIDAK sepenuhnya bebas —
    lihat catatan pada langkah 3 di bawah soal kenapa slide 14 harus
    diproses SEBELUM langkah 7 (penggantian teks generik).

    KETAHANAN TERHADAP ERROR PER SLIDE: mulai dari langkah 3a, tiap blok
    slide dijalankan lewat `_safe_apply_slide` (lihat docstring-nya) —
    kalau satu slide gagal diisi (mis. data region-nya kosong untuk
    metrik tertentu, atau template kehilangan token/placeholder yang
    diharapkan), errornya dicetak ke stdout sebagai `[WARN] <label>: ...`
    dan slide itu dibiarkan apa adanya (isi sebelumnya, atau asli
    template), TAPI proses region ini TETAP LANJUT ke slide-slide
    berikutnya dan tetap menghasilkan file PPTX di `out_path` — bukan
    berhenti total. Pola yang sama sudah lebih dulu dipakai untuk loop
    update chart di langkah 2.

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
        slides: `None` (baku) untuk memproses SEMUA slide seperti biasa,
                atau iterable of int berisi nomor slide 1-indexed yang mau
                diproses saja (mis. `{3, 14, 16}` atau `[3, 14, 16]`,
                lihat `_slide_and_chart_filters`). Chart yang bukan milik
                slide-slide yang diminta (lihat `SLIDE_CHART_RANGES`) juga
                ikut dilewati di langkah 2, supaya benar-benar hanya
                bagian yang diminta yang diproses ulang.

                PENTING: PPTX yang dihasilkan TETAP UTUH 19 slide (dibuka
                dari `TEMPLATE_PATH` seperti biasa) — slide yang TIDAK
                diminta bukan dihapus, melainkan dibiarkan dalam kondisi
                ASLI TEMPLATE (belum diisi data region apa pun), karena
                fungsi ini selalu membangun dari template dari awal, bukan
                menambal file hasil generate sebelumnya. Cocok untuk
                "lihat cepat" satu/dua slide tanpa menunggu seluruh 19
                slide diproses; kalau yang dibutuhkan adalah menambal
                sebagian slide di PPTX yang SUDAH jadi tanpa mengubah
                slide lain yang sudah terisi, itu perlu alur kerja
                terpisah (tidak didukung fungsi ini).
        verbose: jika True, mencetak progres ke stdout (path output, jumlah
                 chart ditemukan) — tidak memengaruhi logika, murni logging.

    Return: path (str) file PPTX yang sudah disimpan
    (`OUTPUT_DIR/HR_Dashboard_{region_dengan_underscore}.pptx`).
    """
    safe_name = region.replace(" ", "_")
    out_path = os.path.join(OUTPUT_DIR, f"HR_Dashboard_{safe_name}.pptx")
    wanted_slides, wanted_charts = _slide_and_chart_filters(slides)

    if verbose:
        print(f"  Generating: {out_path}")
        if wanted_slides is not None:
            print(f"    Slides filter: {sorted(wanted_slides)}")

    with _perf_metrics.phase("rendering"):
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
            #    proses chart-chart lainnya. Kalau `slides` diisi, chart yang
            #    bukan milik slide yang diminta (`wanted_charts`) dilewati.
            for chart_num, chart_info in sorted(chart_map.items()):
                if wanted_charts is not None and chart_num not in wanted_charts:
                    continue
                _chart_t0 = _time.perf_counter()
                try:
                    _update_chart(editor, chart_num, chart_info, data, region)
                    _perf_metrics.record_chart(True, elapsed=_time.perf_counter() - _chart_t0, chart_num=chart_num)
                except Exception as e:
                    print(f"    [WARN] chart{chart_num}: {e}")
                    _perf_metrics.record_chart(False, f"chart{chart_num}: {type(e).__name__}", elapsed=_time.perf_counter() - _chart_t0, chart_num=chart_num)

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

            # 3a. Mengisi 3 tabel statis persentase AGE/EDU/LOS slide 4
            #     (_apply_slide4_lea_tables) lalu highlight-nya
            #     (_apply_slide4_highlights). Dibungkus `_safe_apply_slide`
            #     (lihat docstring-nya) supaya kalau salah satu langkah di
            #     slide ini gagal (mis. data region kosong untuk metrik
            #     tertentu, atau token/placeholder yang diharapkan tidak
            #     ketemu di template), errornya dicetak sebagai warning dan
            #     PPTX tetap lanjut dibuat dengan slide/data lain yang
            #     berhasil — bukan menghentikan seluruh proses region ini.
            if len(slide_paths_early) > 3 and (wanted_slides is None or 4 in wanted_slides):
                def _step_slide4(xml, data=data, region=region):
                    xml = _apply_slide4_lea_tables(xml, data, region)
                    return _apply_slide4_highlights(xml, data, region)
                _safe_apply_slide(editor, "slide4", slide_paths_early[3], _step_slide4)

            # 3b. Mengisi 24 kotak highlight persentase slide 5
            #     (_apply_slide5_pct_boxes) — lihat penjelasan pemetaan shape
            #     di get_slide5_pct_boxes (chart_data.py).
            if len(slide_paths_early) > 4 and (wanted_slides is None or 5 in wanted_slides):
                def _step_slide5(xml, data=data, region=region):
                    xml = _apply_slide5_pct_boxes(xml, data, region)
                    return _apply_slide5_highlights(xml, data, region)
                _safe_apply_slide(editor, "slide5", slide_paths_early[4], _step_slide5)

            if len(slide_paths_early) > 13 and (wanted_slides is None or 14 in wanted_slides):
                def _step_slide14(xml, data=data, region=region):
                    xml = _apply_slide14_tables(xml, data)
                    xml = _apply_slide14_top3_highlights(xml, data)
                    xml = _apply_slide14_rect_highlights(xml, region)
                    xml = _apply_slide14_reason_table(xml, data, region)
                    return _apply_slide14_highlights(xml, data, region)
                _safe_apply_slide(editor, "slide14", slide_paths_early[13], _step_slide14)

            # 4. Mengisi 2 tabel "Attrition Report by Function" slide 15
            #    (_apply_slide15_tables) lalu menyetel 9×3 indikator naik/
            #    turun/tetapnya (_apply_slide15_indicators).
            if len(slide_paths_early) > 14 and (wanted_slides is None or 15 in wanted_slides):
                def _step_slide15(xml, data=data, region=region):
                    xml = _apply_slide15_tables(xml, data, region)
                    xml = _apply_slide15_indicators(xml, data, region)
                    return _apply_slide15_highlights(xml, data, region)
                _safe_apply_slide(editor, "slide15", slide_paths_early[14], _step_slide15)

            # 5. Mengisi tabel field-attrition slide 16: tabel per-cabang Sales
            #    (table_index=0) dan per-cluster Collection (table_index=1)
            #    lewat _apply_slide16_branch_table (dua kali, dengan
            #    data_key/grand_total_key berbeda), lalu tabel reason-out
            #    (_apply_slide16_reason_table).
            if len(slide_paths_early) > 15 and (wanted_slides is None or 16 in wanted_slides):
                def _step_slide16(xml, data=data, region=region):
                    xml = _apply_slide16_branch_table(
                        xml, data, region, table_index=0,
                        data_key="branch_sales", grand_total_key="Sales Officer")
                    xml = _apply_slide16_branch_table(
                        xml, data, region, table_index=1,
                        data_key="branch_collection", grand_total_key="Collection Officer")
                    xml = _apply_slide16_reason_table(xml, data, region)
                    return _apply_slide16_highlight(xml, data, region)
                _safe_apply_slide(editor, "slide16", slide_paths_early[15], _step_slide16)

            # 6. Mengisi ringkasan fraud + indikatornya slide 18
            #    (_apply_slide18_summary), lalu 2 tabel Worst-5 potloss
            #    (_apply_slide18_worst5_table): table_index=0 untuk Branch SSD,
            #    table_index=1 untuk Cluster Collection.
            if len(slide_paths_early) > 17 and (wanted_slides is None or 18 in wanted_slides):
                def _step_slide18(xml, data=data, region=region):
                    xml = _apply_slide18_summary(xml, data, region)
                    xml = _apply_slide18_worst5_table(xml, data, region, table_index=0, data_key="worst5_branch_25")
                    xml = _apply_slide18_worst5_table(xml, data, region, table_index=1, data_key="worst5_cluster_25")
                    return _apply_slide18_highlights(xml, data, region)
                _safe_apply_slide(editor, "slide18", slide_paths_early[17], _step_slide18)

            # 7. Pass penggantian teks generik untuk SEMUA slide: nama region
            #    (dan angka total HC khusus slide 4) lewat
            #    _build_slide_replacements + _replace_text_in_slide. Slide yang
            #    tabelnya sudah diisi di langkah 3-6 di atas (14, 15, 16, 18)
            #    tetap ikut lewat sini untuk mengganti teks judul/label lain di
            #    luar tabel (mis. nama region di judul slide) — bukan berarti
            #    tabelnya diproses ulang di sini. Tiap slide dibungkus
            #    `_safe_apply_slide` sendiri-sendiri supaya kegagalan di satu
            #    slide tidak ikut menggagalkan penggantian teks di slide lain.
            #    Kalau `slides` diisi, slide di luar `wanted_slides` dilewati.
            slide_paths = _get_slide_paths(editor)
            for slide_idx, slide_path in enumerate(slide_paths):
                if wanted_slides is not None and (slide_idx + 1) not in wanted_slides:
                    continue
                replacements = _build_slide_replacements(slide_idx, data, region)
                if not replacements:
                    continue
                _safe_apply_slide(editor, f"slide{slide_idx + 1} text replacements", slide_path,
                                   lambda xml, replacements=replacements: _replace_text_in_slide(xml, replacements))

            # 8. Memindahkan kotak highlight merah slide 3
            #    (_compute_slide3_rect_positions + _apply_slide3_rect_updates)
            #    dan segitiga indikator + text box YoY%-nya
            #    (_compute_slide3_indicator_updates + _apply_slide3_indicator_updates)
            #    ke posisi bar milik `region` pada masing-masing dari 6 chart.
            #    Dijalankan setelah langkah 7 (bukan sebelum, seperti slide 14)
            #    karena slide 3 tidak punya masalah placeholder ambigu yang
            #    sama — urutan di sini tidak kritikal terhadap langkah 7.
            if len(slide_paths) > 2 and (wanted_slides is None or 3 in wanted_slides):
                def _step_slide3(xml, data=data, region=region):
                    rect_updates = _compute_slide3_rect_positions(data, region)
                    if rect_updates:
                        xml = _apply_slide3_rect_updates(xml, rect_updates)
                    indicator_updates = _compute_slide3_indicator_updates(data, region)
                    if indicator_updates:
                        xml = _apply_slide3_indicator_updates(xml, indicator_updates)
                    return _apply_slide3_highlights(xml, data, region)
                _safe_apply_slide(editor, "slide3", slide_paths[2], _step_slide3)

            # 9. Mengisi tabel training BSC slide 10 (_apply_slide10_table —
            #    sama untuk semua region), memindahkan kotak highlight
            #    merahnya ke baris `region` (_apply_slide10_rect_highlight),
            #    lalu mengisi draf insight otomatisnya (_apply_slide10_insight).
            if len(slide_paths) > 9 and (wanted_slides is None or 10 in wanted_slides):
                def _step_slide10(xml, data=data, region=region):
                    xml = _apply_slide10_table(xml, data)
                    xml = _apply_slide10_rect_highlight(xml, region)
                    return _apply_slide10_insight(xml, data, region)
                _safe_apply_slide(editor, "slide10", slide_paths[9], _step_slide10)

            # 9b. Mengisi highlight PA Sales slide 6 dan PA Collection slide 7
            #     (_apply_slide6_highlights / _apply_slide7_highlights).
            if len(slide_paths) > 5 and (wanted_slides is None or 6 in wanted_slides):
                _safe_apply_slide(editor, "slide6", slide_paths[5],
                                   lambda xml, data=data, region=region: _apply_slide6_highlights(xml, data, region))

            if len(slide_paths) > 6 and (wanted_slides is None or 7 in wanted_slides):
                _safe_apply_slide(editor, "slide7", slide_paths[6],
                                   lambda xml, data=data, region=region: _apply_slide7_highlights(xml, data, region))

            # 10. Simpan seluruh perubahan yang terkumpul di `editor` sebagai
            #     file PPTX baru di `out_path` (menulis ulang zip, menjaga tipe
            #     kompresi tiap entry seperti aslinya — lihat PptxEditor.save
            #     di xml_updater.py).
            editor.save(out_path)

    return out_path
