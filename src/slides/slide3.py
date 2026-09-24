"""
Slide 3 — Perbandingan YoY per fungsi (HC, SSD, COLL, CREDIT, LAR, BISNIS),
6 chart bar horizontal. Modul ini menangani 3 shape overlay yang harus
dipindah/diisi ulang per region: kotak highlight merah, segitiga indikator
naik/turun, dan text box %YoY — semuanya independen dari isi chart itu
sendiri (chart-nya sendiri diperbarui lewat pptx_common._update_chart via
CHART_DATA_FN di chart_data.py).
"""
import re

from ..chart_data import (
    get_slide3_highlights,
    chart1_hc, chart2_ssd, chart3_coll, chart4_credit, chart5_lar, chart6_bisnis,
)

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
    1: 3057837,   # HC (nilai template sebelumnya, 3079791, sudah tidak
                  # cocok lagi dengan template saat ini — template.pptx
                  # sempat di-resave dengan sedikit pergeseran koordinat;
                  # nilai ini diverifikasi ulang langsung dari shape
                  # segitiga di "Group 38" pada template saat ini)
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

# Pergeseran horizontal (x, EMU) segitiga indikator + text box YoY% ke
# KANAN, supaya tidak lagi menimpa area chart (bar terpanjang + data
# label bawaannya bisa muncul di mana saja di sepanjang lebar chart).
#
# KENAPA DIBUTUHKAN: posisi x template untuk kedua shape ini ternyata
# ada JAUH DI DALAM area chart (bukan di sebelah kanannya seperti
# kelihatannya dari koordinat mentah) — grpSp pembungkus tiap chart (mis.
# "Group 27" untuk LAR, "Group 31" untuk BISNIS) punya transform sendiri
# yang menggeser posisi slide-absolute jauh dari yang terlihat di
# `<a:off>` mentahnya. Setelah posisi slide-absolute chart dihitung ulang
# (menggabungkan `off`/`ext`/`chOff`/`chExt` grup pembungkusnya), tepi
# kanan tiap chart ternyata 700 ribuan-1,3 jutaan EMU SETELAH posisi x
# indikator saat ini — cukup untuk menutupi label data bar yang panjang.
#
# NILAI: percobaan pertama menggeser indikator sampai TEPAT MELEWATI tepi
# kanan chart (tepi_kanan_chart - x_lama_textbox + margin), tapi ini
# TERLALU JAUH — indikatornya jadi terlihat terpisah/mengambang dari
# chart-nya, bukan lagi "menempel" di ujung bar (feedback pengguna: "out
# of the chart box"). Nilai sekarang jauh lebih kecil (nudge, bukan
# relokasi penuh) — cukup untuk memberi sedikit jarak dari ujung bar
# terpanjang tanpa membuat indikator terlihat lepas dari chart-nya. Sama
# untuk keenam chart (dulu per-chart, tapi nilainya memang selalu sama
# sejak jadi nudge kecil, jadi disederhanakan jadi satu angka). Kalau
# label data masih sesekali tertutup, nilai ini bisa dinaikkan bertahap;
# kalau masih terlihat terlalu jauh, diturunkan.
_SLIDE3_INDICATOR_X_SHIFT = 250000

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

    PENTING soal `v25` kosong/0 (mis. Pasima pada chart4/CREDIT — region
    itu tidak punya data CREDIT tahun lalu sama sekali, `v25=None`):
    `delta` (posisi) DIHITUNG DAN DIKEMBALIKAN TETAP, karena `delta` cuma
    bergantung pada RANKING (`v_k`, posisi region di `cats`), bukan pada
    nilai `v25`/`v26` itu sendiri — jadi indikator TETAP harus pindah ke
    bar yang benar meski %YoY-nya tidak bisa dihitung. Dulu seluruh chart
    itu di-`continue` (dilewati total) begitu `v25` kosong, sehingga
    indikatornya macet di posisi default template (bug nyata — laporan
    pengguna "ada indikator yang tidak berpindah tempat"). Sekarang hanya
    `pct_str`/`is_increase` yang di-fallback ("N/A", tidak naik/turun),
    posisi tetap ikut pindah seperti seharusnya.

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

        v_k = N - 1 - data_idx
        v_k_jt = _SLIDE3_TEMPLATE_V_K_JT[chart_num]
        bar_step = _SLIDE3_EMPIRICAL_BAR_STEP[chart_num]
        delta = int(round((v_k - v_k_jt) * bar_step))

        if not v25:
            # %YoY tidak bisa dihitung (pembagi kosong/0) — posisi tetap
            # dipindah (lihat catatan di atas), teks/arah di-fallback ke
            # nilai netral yang aman daripada mengarang angka.
            pct_str = "N/A"
            is_increase = bool(v26)
        else:
            yoy_pct = (v26 - v25) / v25 * 100
            pct_str = f"{yoy_pct:.1f}".replace(".", ",") + "%"
            is_increase = yoy_pct >= 0
        results.append((chart_num, delta, pct_str, is_increase))

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

    Selain y, x KEDUA shape ini JUGA digeser ke kanan sejauh
    `_SLIDE3_INDICATOR_X_SHIFT` (nilai tetap, sama untuk semua chart dan
    semua region — lihat catatan lengkap di konstanta itu) supaya sedikit
    menjauh dari ujung bar/data-label bawaan chart. Rasio skala x grup
    pembungkus kotak highlight untuk kelima chart sekunder (bukan HC)
    sudah diverifikasi 1:1 (ext.cx == chExt.cx), begitu juga hampir 1:1
    untuk HC, jadi nilai shift EMU yang sama bisa dipakai langsung tanpa
    konversi skala untuk segitiga (koordinat-anak grup) maupun text box
    (koordinat slide) — sama seperti alasan `delta` y bisa dipakai
    langsung tanpa konversi.

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

    # Membangun tabel lookup berdasarkan nilai y template lama. Pergeseran
    # x (`_SLIDE3_INDICATOR_X_SHIFT`) sama untuk semua chart, jadi tidak
    # perlu ikut disimpan per-entri di sini — cukup dipakai langsung
    # sebagai konstanta di kedua cabang `repl` di bawah.
    tri_map  = {}   # old_tri_y  → (new_tri_y, is_increase)
    text_map = {}   # old_text_y → (new_text_y, pct_str)
    for chart_num, delta, pct_str, is_increase in indicator_updates:
        tri_map[_SLIDE3_TEMPLATE_TRI_Y[chart_num]]  = (
            _SLIDE3_TEMPLATE_TRI_Y[chart_num]  + delta, is_increase)
        text_map[_SLIDE3_TEMPLATE_TEXT_Y[chart_num]] = (
            _SLIDE3_TEMPLATE_TEXT_Y[chart_num] + delta, pct_str)

    SP_RE = re.compile(r'(<p:sp\b.*?</p:sp>)', re.DOTALL)

    def shift_x(m):
        return f'x="{int(m.group(1)) + _SLIDE3_INDICATOR_X_SHIFT}"'

    def repl(m):
        sp = m.group(1)

        if 'prst="triangle"' in sp:
            for old_y, (new_y, is_increase) in tri_map.items():
                if f'y="{old_y}"' not in sp:
                    continue
                # Bangun ulang seluruh elemen xfrm: geser x menjauh dari area
                # chart, perbarui y, dan atur arah secara eksplisit.
                # Naik ▲: rot=10800000 flipV=1  (saling meniadakan → menghadap atas, sesuai template)
                # Turun ▼: rot=10800000 saja     (rotasi 180° → menghadap bawah)
                def make_xfrm_repl(oy, ny, inc):
                    def xfrm_repl(xm):
                        inner = xm.group(1).replace(f'y="{oy}"', f'y="{ny}"', 1)
                        inner = re.sub(r'x="(\d+)"', shift_x, inner, count=1)
                        rot_attrs = 'rot="10800000" flipV="1"' if inc else 'rot="10800000"'
                        return f'<a:xfrm {rot_attrs}>{inner}</a:xfrm>'
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

                def off_repl(om, new_y=new_y):
                    new_x = int(om.group(1)) + _SLIDE3_INDICATOR_X_SHIFT
                    return f'<a:off x="{new_x}" y="{new_y}"/>'
                sp = re.sub(rf'<a:off x="(\d+)" y="{old_y}"/>', off_repl, sp, count=1)
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
        escaped = value.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        token = f"{{SLIDE3_HIGHLIGHT_{key}}}"
        text = text.replace(f'<a:t>{token}</a:t>', f'<a:t>{escaped}</a:t>', 1)
    return text.encode('utf-8')


