"""
Mendefinisikan array data untuk tiap nomor chart, berdasarkan data yang sudah
di-load per region (lihat data_loader.load_all()).
Mengembalikan (categories, series_list) atau (scatter_data) yang siap
diteruskan ke XML updater (src/xml_updater.py).
"""
import math

from .config import (
    REGIONS, REGION_WILAYAH, REGION_ABBREV, FUNCTION_ORDER,
    LOS_ORDER, EDU_ORDER, AGE_ORDER,
)


def _round_half_up(v):
    """
    Bulatkan `v` ke integer terdekat dengan konvensi Excel (0,5 selalu naik
    lewat `floor(v + 0.5)`), BUKAN `round()` bawaan Python yang pakai
    aturan half-to-even. Dipakai untuk nilai chart yang harusnya bilangan
    bulat (headcount) tapi datang sebagai rata-rata/float (mis. "AVG
    NASIONAL" pada chart7_function, "nas_sales_fl" dkk pada
    chart16_fl_sales/chart17_fl_coll), supaya angka yang tampil di chart
    (data label) berupa bilangan bulat yang wajar, bukan pecahan panjang
    seperti 845.4166666666666.

    Return: int, atau 0 kalau `v` None.
    """
    if v is None:
        return 0
    return int(math.floor(float(v) + 0.5))


def _safe_pct(val):
    """Mengembalikan nilai float persentase; None diperlakukan sebagai 0.0."""
    if val is None:
        return 0.0
    return float(val)


def _fmt_pct(v):
    """Format angka desimal (fraksi) menjadi teks anotasi '+X.X%' / '-X.X%'."""
    if v is None:
        return "0.0%"
    pct = v * 100 if abs(v) <= 1.5 else v  # menangani bentuk fraksi maupun bentuk persen
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct:.1f}%"


# ---------------------------------------------------------------------------
# SLIDE 3 (chart 1-6) - Perbandingan YoY, SEMUA 12 region ditampilkan per chart
# Data chart ini identik di semua file PPTX output.
# ---------------------------------------------------------------------------
def chart1_hc(d):        return _slide3_chart(d["slide3"]["HC"])
def chart2_ssd(d):       return _slide3_chart(d["slide3"]["SSD"])
def chart3_coll(d):      return _slide3_chart(d["slide3"]["COLL"])
def chart4_credit(d):    return _slide3_chart(d["slide3"]["CREDIT"])
def chart5_lar(d):       return _slide3_chart(d["slide3"]["LAR"])
def chart6_bisnis(d):    return _slide3_chart(d["slide3"]["BISNIS"])


def _slide3_chart(entries):
    """
    Helper bersama untuk chart1_hc..chart6_bisnis. `entries` berisi daftar
    (region_name, val_2026, val_2025) yang sudah terurut ascending, diambil
    dari salah satu sub-key d["slide3"][<metrik>].
    Membentuk ulang menjadi (categories, series_list) dengan dua series
    ("Mei-26", "Mei-25"). Karena chart 1-6 menampilkan semua region sekaligus,
    fungsi pemanggilnya (chart1_hc dst.) hanya menerima `d` tanpa `region`.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    cats  = [e[0] for e in entries]
    v26   = [e[1] for e in entries]
    v25   = [e[2] for e in entries]
    return cats, [("Mei-26", v26), ("Mei-25", v25)]


# ---------------------------------------------------------------------------
# SLIDE 4 (chart 7-17)
# ---------------------------------------------------------------------------
def chart7_function(d, region):
    """
    Ambil breakdown headcount per fungsi (urutan tetap dari FUNCTION_ORDER)
    dari d["s4_func"][region] dibandingkan dengan d["s4_func"]["AVG NASIONAL"].
    Categories = daftar fungsi, series = ("Region", ...) dan ("AVG NAS", ...).
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    func_d  = d["s4_func"].get(region, {})
    nas_d   = d["s4_func"].get("AVG NASIONAL", {})
    cats    = FUNCTION_ORDER
    reg_vals = [_round_half_up(func_d.get(f)) for f in cats]
    nas_vals = [_round_half_up(nas_d.get(f)) for f in cats]
    return cats, [("Region", reg_vals), ("AVG NAS", nas_vals)]


def chart8_hc_npat(d, region):
    """
    Ambil HC dan NPAT per tahun dari d["s4_npat"][region], untuk 4 periode
    tetap (FY23, FY24, FY25, 5M2026). Nilai NPAT untuk periode 5M2026 belum
    tersedia sehingga diisi None. Series: ("HC Managed", ...) dan
    ("NPAT", ...).
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    npat = d["s4_npat"].get(region, {})
    cats  = ["FY23", "FY24", "FY25", "5M2026"]
    hc    = [npat.get("FY23", {}).get("HC", 0) or 0,
             npat.get("FY24", {}).get("HC", 0) or 0,
             npat.get("FY25", {}).get("HC", 0) or 0,
             npat.get("5M2026", {}).get("HC", 0) or 0]
    npat_v = [npat.get("FY23", {}).get("NPAT", 0) or 0,
              npat.get("FY24", {}).get("NPAT", 0) or 0,
              npat.get("FY25", {}).get("NPAT", 0) or 0,
              None]  # NPAT untuk 5M2026 belum tersedia
    return cats, [("HC Managed", hc), ("NPAT", npat_v)]


def chart9_work_contract(d, region):
    """
    Ambil persentase status kontrak kerja (Permanent/Contract/Outsources)
    dari d["s4_wc"][region] (nilai region) vs field "nas_*" (nilai
    nasional) di dict yang sama. `_safe_pct` dipakai untuk menangani nilai
    None.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    wc = d["s4_wc"].get(region, {})
    cats = ["Permanent", "Contract", "Outsources"]
    reg_v = [_safe_pct(wc.get("pct_permanent")),
             _safe_pct(wc.get("pct_contract")),
             _safe_pct(wc.get("pct_os"))]
    nas_v = [_safe_pct(wc.get("nas_permanent")),
             _safe_pct(wc.get("nas_contract")),
             _safe_pct(wc.get("nas_os"))]
    return cats, [("Region", reg_v), ("NAS", nas_v)]


def chart10_los(d, region):
    """
    Histogram jumlah pegawai per kelompok masa kerja (LOS), diambil dari
    d["s4_lea"][region]["LOS"] (list of (label, count)).
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    los = d["s4_lea"].get(region, {}).get("LOS", [])
    cats  = [x[0] for x in los]
    vals  = [x[1] for x in los]
    return cats, [("Count", vals)]


def chart11_edu(d, region):
    """
    Histogram jumlah pegawai per kelompok pendidikan (EDU), diambil dari
    d["s4_lea"][region]["EDU"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    edu = d["s4_lea"].get(region, {}).get("EDU", [])
    cats = [x[0] for x in edu]
    vals = [x[1] for x in edu]
    return cats, [("Count", vals)]


def chart12_age(d, region):
    """
    Histogram jumlah pegawai per kelompok usia (AGE), diambil dari
    d["s4_lea"][region]["AGE"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    age = d["s4_lea"].get(region, {}).get("AGE", [])
    cats = [x[0] for x in age]
    vals = [x[1] for x in age]
    return cats, [("Count", vals)]


def chart13_soc_sales(d, region):
    """
    Span of control untuk jalur SALES (SO-SH, SH-MGR), REG vs NAS, diambil
    dari d["s4_soc"][region] dan d["s4_soc"]["NAS"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    soc = d["s4_soc"].get(region, {})
    nas = d["s4_soc"].get("NAS", {})
    cats  = ["SO-SH", "SH-MGR"]
    reg_v = [soc.get("sales_so_sh", 0), soc.get("sales_sh_mgr", 0)]
    nas_v = [nas.get("sales_so_sh", 0), nas.get("sales_sh_mgr", 0)]
    return cats, [("Region", reg_v), ("NAS", nas_v)]


def chart14_soc_coll(d, region):
    """
    Span of control untuk jalur COLL (CO-CH, CH-MGR), REG vs NAS, diambil
    dari d["s4_soc"][region] dan d["s4_soc"]["NAS"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    soc = d["s4_soc"].get(region, {})
    nas = d["s4_soc"].get("NAS", {})
    cats  = ["CO-CH", "CH-MGR"]
    reg_v = [soc.get("coll_so_sh", 0), soc.get("coll_sh_mgr", 0)]
    nas_v = [nas.get("coll_so_sh", 0), nas.get("coll_sh_mgr", 0)]
    return cats, [("Region", reg_v), ("NAS", nas_v)]


def chart15_npat_hc_line(d, region):
    """
    Line chart rasio NPAT/HC per tahun (FY23-FY25), diambil dari
    d["s4_npat"][region][<tahun>]["NPAT_HC"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    npat = d["s4_npat"].get(region, {})
    cats  = ["FY23", "FY24", "FY25"]
    vals  = [npat.get("FY23", {}).get("NPAT_HC", 0) or 0,
             npat.get("FY24", {}).get("NPAT_HC", 0) or 0,
             npat.get("FY25", {}).get("NPAT_HC", 0) or 0]
    return cats, [("NPAT/HC", vals)]


def chart16_fl_sales(d, region):
    """
    Proporsi frontliner vs non-frontliner untuk jalur SALES, REG vs NAS,
    diambil dari d["s4_fl"][region] (field "sales_fl"/"sales_non_fl" untuk
    region, "nas_sales_fl"/"nas_sales_non_fl" untuk nasional).
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    fl = d["s4_fl"].get(region, {})
    cats  = ["Frontliners", "Non-Frontliners"]
    reg_v = [_round_half_up(fl.get("sales_fl")), _round_half_up(fl.get("sales_non_fl"))]
    nas_v = [_round_half_up(fl.get("nas_sales_fl")), _round_half_up(fl.get("nas_sales_non_fl"))]
    return cats, [("Region", reg_v), ("NAS", nas_v)]


def chart17_fl_coll(d, region):
    """
    Proporsi frontliner vs non-frontliner untuk jalur COLL, REG vs NAS,
    diambil dari d["s4_fl"][region] (field "coll_fl"/"coll_non_fl" untuk
    region, "nas_coll_fl"/"nas_coll_non_fl" untuk nasional).
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    fl = d["s4_fl"].get(region, {})
    cats  = ["Frontliners", "Non-Frontliners"]
    reg_v = [_round_half_up(fl.get("coll_fl")), _round_half_up(fl.get("coll_non_fl"))]
    nas_v = [_round_half_up(fl.get("nas_coll_fl")), _round_half_up(fl.get("nas_coll_non_fl"))]
    return cats, [("Region", reg_v), ("NAS", nas_v)]


# ---------------------------------------------------------------------------
# SLIDE 5 (chart 18-23)
# ---------------------------------------------------------------------------
def chart18_edu_fl_nat(d, region):
    """
    Breakdown pendidikan frontliner NASIONAL (Field Sales vs Field Coll),
    diambil dari d["s5_nat"]["EDU"] (list of (label, val_sales, val_coll));
    parameter `region` tidak dipakai karena data ini tetap/nasional.

    PENTING — chart18 vs chart23 di slide 5:
    Pada template, blok "Pendidikan Frontliners" menempatkan label
    "Regional" di atas chart23 dan label "Nasional" di atas chart18 (urut
    dari atas: label Regional → chart23 → label Nasional → chart18). Jadi
    chart18 adalah chart NASIONAL dan chart23 adalah chart REGIONAL —
    kebalikan dari nomor urutnya, dan kebalikan dari pemetaan lama yang
    membuat kedua chart tertukar isinya.

    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s5_nat"]
    edu = nat.get("EDU", [])
    cats   = [x[0] for x in edu]
    vsales = [x[1] for x in edu]
    vcoll  = [x[2] for x in edu]
    return cats, [("Field Sales", vsales), ("Field Coll", vcoll)]


def chart19_age_fl_reg(d, region):
    """
    Breakdown usia frontliner REGION (Field Sales vs Field Coll), diambil
    dari d["s57"][region]["age_sales"] dan ["age_coll"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    age_sales = s.get("age_sales", [])
    age_coll  = s.get("age_coll", [])
    cats   = [x[0] for x in age_sales]
    vsales = [x[1] for x in age_sales]
    vcoll  = [x[1] for x in age_coll]
    return cats, [("Field Sales", vsales), ("Field Coll", vcoll)]


def chart20_los_fl_reg(d, region):
    """
    Breakdown masa kerja frontliner REGION (Field Sales vs Field Coll),
    diambil dari d["s57"][region]["los_sales"] dan ["los_coll"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    los_sales = s.get("los_sales", [])
    los_coll  = s.get("los_coll", [])
    cats   = [x[0] for x in los_sales]
    vsales = [x[1] for x in los_sales]
    vcoll  = [x[1] for x in los_coll]
    return cats, [("Field Sales", vsales), ("Field Coll", vcoll)]


def chart21_los_fl_nat(d, region):
    """
    Breakdown masa kerja frontliner NASIONAL (sama untuk semua region),
    diambil dari d["s5_nat"]["LOS"] (list berisi (label, val_sales,
    val_coll)); parameter `region` tidak dipakai untuk memilih data karena
    chart ini bersifat nasional/tetap.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s5_nat"]
    los = nat.get("LOS", [])
    cats   = [x[0] for x in los]
    vsales = [x[1] for x in los]
    vcoll  = [x[2] for x in los]
    return cats, [("Field Sales", vsales), ("Field Coll", vcoll)]


def chart22_age_fl_nat(d, region):
    """
    Breakdown usia frontliner NASIONAL, diambil dari d["s5_nat"]["AGE"];
    parameter `region` tidak dipakai karena data ini tetap/nasional.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s5_nat"]
    age = nat.get("AGE", [])
    cats   = [x[0] for x in age]
    vsales = [x[1] for x in age]
    vcoll  = [x[2] for x in age]
    return cats, [("Field Sales", vsales), ("Field Coll", vcoll)]


def chart23_edu_fl_reg(d, region):
    """
    Breakdown pendidikan frontliner REGION (Field Sales vs Field Coll),
    diambil dari d["s57"][region]["edu_sales"] dan ["edu_coll"] (list of
    (label, count)).

    PENTING: chart23 adalah chart REGIONAL, bukan nasional — lihat catatan
    lengkap di `chart18_edu_fl_nat()` soal posisi label "Regional"/
    "Nasional" pada blok Pendidikan di template slide 5.

    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    edu_sales = s.get("edu_sales", [])
    edu_coll  = s.get("edu_coll", [])
    cats  = [x[0] for x in edu_sales] if edu_sales else []
    vsales = [x[1] for x in edu_sales]
    vcoll  = [x[1] for x in edu_coll]
    return cats, [("Field Sales", vsales), ("Field Coll", vcoll)]


# ---------------------------------------------------------------------------
# SLIDE 6 (chart 24-32) - PA Sales
# ---------------------------------------------------------------------------
def chart24_pa_sales_los_reg(d, region):
    """
    PA (performance appraisal) Sales berdasarkan masa kerja (LOS) - REGION,
    3 bar bertumpuk (PA2, PA3, PA>=4), diambil dari
    d["s57"][region]["pa_sales_los"] (list of (label, v2, v3, v4)).
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("pa_sales_los", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]
    v3 = [x[2] for x in rows]
    v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart25_count_sales_los_line(d, region):
    """
    Line chart overlay jumlah pegawai (count) untuk PA Sales berdasarkan
    LOS, diambil dari d["s57"][region]["count_sales_los"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("count_sales_los", [])
    cats = [x[0] for x in rows]
    vals = [x[1] for x in rows]
    return cats, [("Count", vals)]


def chart26_pa_sales_age_reg(d, region):
    """
    PA Sales berdasarkan usia (AGE) - REGION, diambil dari
    d["s57"][region]["pa_sales_age"]. Bentuk sama seperti chart24.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("pa_sales_age", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart27_count_sales_age_line(d, region):
    """
    Line chart overlay count untuk PA Sales berdasarkan usia, diambil dari
    d["s57"][region]["count_sales_age"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("count_sales_age", [])
    cats = [x[0] for x in rows]
    return cats, [("Count", [x[1] for x in rows])]


def chart28_pa_sales_edu_reg(d, region):
    """
    PA Sales berdasarkan pendidikan (EDU) - REGION, diambil dari
    d["s57"][region]["pa_sales_edu"]. Bentuk sama seperti chart24.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("pa_sales_edu", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart29_count_sales_edu_line(d, region):
    """
    Line chart overlay count untuk PA Sales berdasarkan pendidikan, diambil
    dari d["s57"][region]["count_sales_edu"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("count_sales_edu", [])
    cats = [x[0] for x in rows]
    return cats, [("Count", [x[1] for x in rows])]


def chart30_pa_sales_edu_nat(d, region):
    """
    PA Sales berdasarkan pendidikan - NASIONAL, diambil dari
    d["s6_nat"]["EDU"]; parameter `region` tidak dipakai karena data ini
    tetap/nasional.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s6_nat"]
    rows = nat.get("EDU", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart31_pa_sales_age_nat(d, region):
    """
    PA Sales berdasarkan usia - NASIONAL, diambil dari d["s6_nat"]["AGE"];
    parameter `region` tidak dipakai karena data ini tetap/nasional.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s6_nat"]
    rows = nat.get("AGE", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart32_pa_sales_los_nat(d, region):
    """
    PA Sales berdasarkan masa kerja - NASIONAL, diambil dari
    d["s6_nat"]["LOS"]; parameter `region` tidak dipakai karena data ini
    tetap/nasional.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s6_nat"]
    rows = nat.get("LOS", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


# ---------------------------------------------------------------------------
# SLIDE 7 (chart 33-41) - PA Coll
# ---------------------------------------------------------------------------
def chart33_pa_coll_los_reg(d, region):
    """
    PA Coll berdasarkan masa kerja (LOS) - REGION, diambil dari
    d["s57"][region]["pa_coll_los"]. Bentuk sama seperti chart24 tapi untuk
    jalur Collection.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("pa_coll_los", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart34_count_coll_los_line(d, region):
    """
    Line chart overlay count untuk PA Coll berdasarkan LOS, diambil dari
    d["s57"][region]["count_coll_los"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("count_coll_los", [])
    return [x[0] for x in rows], [("Count", [x[1] for x in rows])]


def chart35_pa_coll_age_reg(d, region):
    """
    PA Coll berdasarkan usia (AGE) - REGION, diambil dari
    d["s57"][region]["pa_coll_age"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("pa_coll_age", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart36_count_coll_age_line(d, region):
    """
    Line chart overlay count untuk PA Coll berdasarkan usia, diambil dari
    d["s57"][region]["count_coll_age"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("count_coll_age", [])
    return [x[0] for x in rows], [("Count", [x[1] for x in rows])]


def chart37_pa_coll_edu_reg(d, region):
    """
    PA Coll berdasarkan pendidikan (EDU) - REGION, diambil dari
    d["s57"][region]["pa_coll_edu"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("pa_coll_edu", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart38_count_coll_edu_line(d, region):
    """
    Line chart overlay count untuk PA Coll berdasarkan pendidikan, diambil
    dari d["s57"][region]["count_coll_edu"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    s = d["s57"].get(region, {})
    rows = s.get("count_coll_edu", [])
    return [x[0] for x in rows], [("Count", [x[1] for x in rows])]


def chart39_pa_coll_edu_nat(d, region):
    """
    PA Coll berdasarkan pendidikan - NASIONAL, diambil dari
    d["s7_nat"]["EDU"]; parameter `region` tidak dipakai karena data ini
    tetap/nasional.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s7_nat"]
    rows = nat.get("EDU", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart40_pa_coll_age_nat(d, region):
    """
    PA Coll berdasarkan usia - NASIONAL, diambil dari d["s7_nat"]["AGE"];
    parameter `region` tidak dipakai karena data ini tetap/nasional.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s7_nat"]
    rows = nat.get("AGE", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


def chart41_pa_coll_los_nat(d, region):
    """
    PA Coll berdasarkan masa kerja - NASIONAL, diambil dari
    d["s7_nat"]["LOS"]; parameter `region` tidak dipakai karena data ini
    tetap/nasional.
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    nat = d["s7_nat"]
    rows = nat.get("LOS", [])
    cats = [x[0] for x in rows]
    v2 = [x[1] for x in rows]; v3 = [x[2] for x in rows]; v4 = [x[3] for x in rows]
    return cats, [("PA 2", v2), ("PA 3", v3), ("PA >= 4", v4)]


# ---------------------------------------------------------------------------
# SLIDE 12 (chart 42-43) - Scatter Productivity
# ---------------------------------------------------------------------------
def chart42_scatter_region(d, region):
    """
    Data scatter tingkat REGION (mencakup 12 region, sama untuk semua file
    output), diambil langsung dari d["s12"]["ALL_REGIONS"]. Parameter
    `region` tidak dipakai karena data ini tetap untuk semua output.
    Berbeda dari chart lain: fungsi ini mengembalikan data scatter mentah
    (list of tuple), BUKAN (categories, series_list) — chart 42 terdaftar
    di SCATTER_CHARTS sehingga _update_chart() di pptx_updater.py
    meneruskan hasilnya ke update_chart_xml_scatter() di xml_updater.py,
    bukan ke update_chart_xml().
    """
    return d["s12"].get("ALL_REGIONS", [])


def chart43_scatter_branch(d, region):
    """
    Data scatter tingkat cabang (branch), spesifik per region, diambil dari
    d["s12"][region].
    Berbeda dari chart lain: fungsi ini mengembalikan data scatter mentah
    (list of tuple), BUKAN (categories, series_list) — chart 43 terdaftar
    di SCATTER_CHARTS sehingga _update_chart() di pptx_updater.py
    meneruskan hasilnya ke update_chart_xml_scatter() di xml_updater.py,
    bukan ke update_chart_xml().
    """
    return d["s12"].get(region, [])


# ---------------------------------------------------------------------------
# SLIDE 18 (chart 44-46) - Fraud Rate
# ---------------------------------------------------------------------------
def chart44_fraud_summary(d, region):
    """
    Ringkasan jumlah Kasus dan Potloss (dikonversi ke Juta rupiah) YoY,
    diambil dari d["fraud"][region] (field "kasus_25"/"kasus_26" dan
    "potloss_25"/"potloss_26"). Dua series: ("2025", ...) dan ("2026", ...).
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    f = d["fraud"].get(region, {})
    cats  = ["Kasus", "Potloss (Jt)"]
    v25   = [f.get("kasus_25", 0), (f.get("potloss_25", 0) or 0) / 1e6]
    v26   = [f.get("kasus_26", 0), (f.get("potloss_26", 0) or 0) / 1e6]
    return cats, [("2025", v25), ("2026", v26)]


def chart45_fraud_by_location(d, region):
    """
    Potloss fraud berdasarkan lokasi (Branch SSD, Cluster Collection) YoY,
    dikonversi ke Juta rupiah, diambil dari d["fraud"][region]
    ["by_location_25"] dan ["by_location_26"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    f = d["fraud"].get(region, {})
    cats  = ["Branch SSD", "Cluster Collection"]
    v25   = [
        (f.get("by_location_25", {}).get("Branch SSD", 0) or 0) / 1e6,
        (f.get("by_location_25", {}).get("Cluster Collection", 0) or 0) / 1e6,
    ]
    v26   = [
        (f.get("by_location_26", {}).get("Branch SSD", 0) or 0) / 1e6,
        (f.get("by_location_26", {}).get("Cluster Collection", 0) or 0) / 1e6,
    ]
    return cats, [("2025", v25), ("2026", v26)]


def chart46_fraud_by_ie(d, region):
    """
    Potloss fraud Internal vs External YoY, dikonversi ke Juta rupiah,
    diambil dari d["fraud"][region]["by_ie_25"] dan ["by_ie_26"].
    Dipanggil lewat CHART_DATA_FN (lihat get_chart_data) dari _update_chart()
    di pptx_updater.py; hasil diteruskan ke update_chart_xml() di
    xml_updater.py.
    """
    f = d["fraud"].get(region, {})
    cats  = ["Internal", "External"]
    v25   = [
        (f.get("by_ie_25", {}).get("Internal", 0) or 0) / 1e6,
        (f.get("by_ie_25", {}).get("External", 0) or 0) / 1e6,
    ]
    v26   = [
        (f.get("by_ie_26", {}).get("Internal", 0) or 0) / 1e6,
        (f.get("by_ie_26", {}).get("External", 0) or 0) / 1e6,
    ]
    return cats, [("2025", v25), ("2026", v26)]


# ---------------------------------------------------------------------------
# Tabel dispatch utama: nomor_chart -> function(d, region)
# ---------------------------------------------------------------------------
CHART_DATA_FN = {
    # Slide 3
    1:  lambda d, r: chart1_hc(d),
    2:  lambda d, r: chart2_ssd(d),
    3:  lambda d, r: chart3_coll(d),
    4:  lambda d, r: chart4_credit(d),
    5:  lambda d, r: chart5_lar(d),
    6:  lambda d, r: chart6_bisnis(d),
    # Slide 4
    7:  chart7_function,
    8:  chart8_hc_npat,
    9:  chart9_work_contract,
    10: chart10_los,
    11: chart11_edu,
    12: chart12_age,
    13: chart13_soc_sales,
    14: chart14_soc_coll,
    15: chart15_npat_hc_line,
    16: chart16_fl_sales,
    17: chart17_fl_coll,
    # Slide 5 — perhatikan chart18=EDU nasional, chart23=EDU regional
    # (kebalikan dari nomor urutnya; lihat catatan di chart18_edu_fl_nat)
    18: chart18_edu_fl_nat,
    19: chart19_age_fl_reg,
    20: chart20_los_fl_reg,
    21: chart21_los_fl_nat,
    22: chart22_age_fl_nat,
    23: chart23_edu_fl_reg,
    # Slide 6
    24: chart24_pa_sales_los_reg,
    25: chart25_count_sales_los_line,
    26: chart26_pa_sales_age_reg,
    27: chart27_count_sales_age_line,
    28: chart28_pa_sales_edu_reg,
    29: chart29_count_sales_edu_line,
    30: chart30_pa_sales_edu_nat,
    31: chart31_pa_sales_age_nat,
    32: chart32_pa_sales_los_nat,
    # Slide 7
    33: chart33_pa_coll_los_reg,
    34: chart34_count_coll_los_line,
    35: chart35_pa_coll_age_reg,
    36: chart36_count_coll_age_line,
    37: chart37_pa_coll_edu_reg,
    38: chart38_count_coll_edu_line,
    39: chart39_pa_coll_edu_nat,
    40: chart40_pa_coll_age_nat,
    41: chart41_pa_coll_los_nat,
    # Slide 12
    42: chart42_scatter_region,
    43: chart43_scatter_branch,
    # Slide 18
    44: chart44_fraud_summary,
    45: chart45_fraud_by_location,
    46: chart46_fraud_by_ie,
}

# Chart yang mengembalikan data scatter (list of tuple), bukan (cats, series)
SCATTER_CHARTS = {42, 43}


def get_chart_data(chart_num, data, region):
    """
    Titik masuk dispatch: mencari fungsi chart di CHART_DATA_FN berdasarkan
    `chart_num`, lalu memanggilnya dengan (data, region). Mengembalikan None
    jika nomor chart tidak dikenal. Dipanggil dari _update_chart() di
    pptx_updater.py; hasilnya diteruskan ke update_chart_xml() atau (untuk
    chart di SCATTER_CHARTS) update_chart_xml_scatter() di xml_updater.py.
    """
    fn = CHART_DATA_FN.get(chart_num)
    if fn is None:
        return None
    return fn(data, region)


def get_slide3_annotations(data, region):
    """
    Menghasilkan dict teks anotasi (format '+X.X%'/'-X.X%') per metrik
    (HC, BISNIS, SSD, COLL, CREDIT, LAR) untuk text box di slide 3, diambil
    dari data["slide3"]["_dif_pct"][region]. Dipanggil dari
    pptx_updater.py (lihat baris di sekitar `ann = get_slide3_annotations(...)`)
    untuk mengisi teks anotasi perubahan YoY pada slide 3.
    """
    dif = data["slide3"].get("_dif_pct", {}).get(region, {})

    def fmt(v):
        if v is None:
            return "0.0%"
        pct = v * 100 if abs(v) < 2 else v
        sign = "+" if pct >= 0 else ""
        return f"{sign}{pct:.1f}%"

    return {
        "HC":     fmt(dif.get("HC")),
        "BISNIS": fmt(dif.get("BISNIS")),
        "SSD":    fmt(dif.get("SSD")),
        "COLL":   fmt(dif.get("COLL")),
        "CREDIT": fmt(dif.get("CREDIT")),
        "LAR":    fmt(dif.get("LAR")),
    }


# Nama tampilan untuk tiap fungsi pada kalimat highlight slide 3 (huruf awal
# kapital agar enak dibaca dalam kalimat, bukan huruf besar semua seperti di
# label chart/tabel).
_SLIDE3_FUNCTION_DISPLAY = {
    "BISNIS": "Bisnis", "SSD": "SSD", "COLL": "Collection", "CREDIT": "Credit", "LAR": "LAR",
}
_SLIDE3_HC_RANK_BUCKETS = (3, 5, 8)


def _join_indonesian(items):
    """Menggabungkan list string jadi satu frasa Indonesia: 'A', 'A dan B', atau 'A, B, dan C'."""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} dan {items[1]}"
    return ", ".join(items[:-1]) + f", dan {items[-1]}"


def get_slide3_highlights(data, region):
    """
    Membuat draf teks highlight/insight otomatis untuk slide 3 (4 slot:
    {SLIDE3_HIGHLIGHT_A..D}), HANYA berdasarkan angka yang benar-benar ada
    di data — tidak menebak alasan bisnis (mis. migrasi cabang, program
    sentralisasi) yang tidak tercatat di mana pun dalam sumber data.

    Cara kerja tiap slot:
      A — Tren HC Managed vs tahun lalu, plus fungsi mana yang berlawanan
          arah dari tren HC (pengecualian). Dihitung dari
          `data["slide3"]["_dif_pct"][region]`, membandingkan tanda
          (naik/turun) HC terhadap tanda tiap fungsi lain (BISNIS, SSD,
          COLL, CREDIT, LAR).
      B — SENGAJA DIKOSONGKAN. Contoh pada Template_Highlight.txt untuk
          slot ini ("migrasi 9 cabang HI", "sentralisasi PPS") adalah
          narasi sebab-akibat yang butuh konteks bisnis nyata, dan itu
          tidak ada di data manapun yang dibaca data_loader.py — mengarang
          alasan seperti itu berisiko salah, jadi dibiarkan kosong untuk
          diisi manual.
      C — Peringkat `region` berdasarkan HC 2026, dibanding 12 region
          lain. Dihitung dari `data["slide3"]["HC"]` (list (region, v26,
          v25) terurut naik) — posisi region dicari, lalu dikonversi jadi
          peringkat dari atas dan dibulatkan ke bucket terdekat (top 3/5/8)
          supaya bunyinya alami seperti pada contoh ("top 3 region").
      D — Insight tambahan sederhana yang bisa dihitung dari data (bukan
          dari contoh Template_Highlight.txt): perubahan HC absolut
          (jumlah orang) dibanding tahun lalu.

    Parameter:
        data: dict hasil `load_all()` (lihat data_loader.py) — hanya
              memakai key `data["slide3"]`.
        region: nama region (str), sama seperti dipakai di seluruh
                pptx_updater.py.

    Return: dict `{"A": str, "B": str, "C": str, "D": str}`. String kosong
    berarti slot itu dikosongkan (lihat alasan B di atas, atau karena data
    yang dibutuhkan tidak ditemukan). Dipanggil dari
    `_apply_slide3_highlights()` di pptx_updater.py, yang mengganti tiap
    token `{SLIDE3_HIGHLIGHT_<key>}` pada slide 3 dengan isi dict ini.
    """
    dif = data.get("slide3", {}).get("_dif_pct", {}).get(region, {})
    hc_entries = data.get("slide3", {}).get("HC", [])  # (region, v26, v25), terurut naik oleh v26

    result = {"A": "", "B": "", "C": "", "D": ""}

    # --- A: tren HC + fungsi yang jadi pengecualian ---
    hc_dif = dif.get("HC")
    if hc_dif is not None:
        hc_naik = hc_dif >= 0
        arah_hc = "kenaikan" if hc_naik else "penurunan"
        arah_lawan = "penurunan" if hc_naik else "kenaikan"

        pengecualian = []
        for fn in ("BISNIS", "SSD", "COLL", "CREDIT", "LAR"):
            v = dif.get(fn)
            if v is not None and (v >= 0) != hc_naik:
                pengecualian.append(_SLIDE3_FUNCTION_DISPLAY.get(fn, fn))

        if pengecualian:
            result["A"] = (
                f"Secara YoY, total HC Managed {region} mengalami {arah_hc} di berbagai fungsi, "
                f"kecuali {_join_indonesian(pengecualian)} yang mengalami {arah_lawan}"
            )
        else:
            result["A"] = f"Secara YoY, total HC Managed {region} mengalami {arah_hc} di seluruh fungsi"

    # --- C & D: butuh posisi region di daftar HC 2026 ---
    idx = next((i for i, e in enumerate(hc_entries) if e[0] == region), None)
    if idx is not None:
        n = len(hc_entries)
        peringkat = n - idx   # hc_entries terurut naik → peringkat dari atas = total - index

        for bucket in _SLIDE3_HC_RANK_BUCKETS:
            if peringkat <= bucket:
                result["C"] = (f"Secara jumlah headcount di 2026, {region} merupakan "
                                f"top {bucket} region dari jumlah headcount")
                break
        else:
            result["C"] = (f"Secara jumlah headcount di 2026, {region} berada di peringkat "
                            f"ke-{peringkat} dari {n} region")

        _, v26, v25 = hc_entries[idx]
        if v26 is not None and v25 is not None:
            delta = v26 - v25
            arah_delta = "bertambah" if delta >= 0 else "berkurang"
            result["D"] = (f"Dibandingkan tahun lalu, headcount {region} {arah_delta} "
                            f"sebanyak {abs(int(round(delta)))} orang")

    return result


def get_slide10_insight(data, region):
    """
    Membuat draf teks insight otomatis untuk slide 10 (1 slot:
    {SLIDE10_INSIGHT}), berdasarkan `data["s10"]` (hasil
    `load_slide10_training()` di data_loader.py) — membandingkan % YTD ACH
    realisasi training `region` terhadap % YTD ACH nasional.

    BEDA DENGAN CONTOH di Template_Highlight.txt: contoh untuk Jawa Tengah
    menyebut "masih di bawah average Nasional" sebagai kalimat statis —
    fungsi ini TIDAK mengasumsikan arahnya selalu "di bawah", melainkan
    membandingkan angka sungguhan tiap kali dipanggil (bisa jadi "di atas",
    "di bawah", atau "setara dengan" tergantung region dan data terkini),
    supaya tetap benar untuk semua 12 region dan kalau datanya berubah.

    Parameter:
        data: dict hasil load_all() — memakai key data["s10"].
        region: nama region (str).

    Return: str (bisa "" kalau data ytd_ach region atau nasional tidak
    ditemukan). Dipanggil dari _apply_slide10_insight() di pptx_updater.py
    untuk mengisi token {SLIDE10_INSIGHT}.
    """
    s10 = data.get("s10", {})
    reg_ach = s10.get(region, {}).get("ytd_ach")
    nat_ach = s10.get("National", {}).get("ytd_ach")
    if reg_ach is None or nat_ach is None:
        return ""

    if reg_ach < nat_ach:
        posisi = "masih di bawah"
    elif reg_ach > nat_ach:
        posisi = "sudah di atas"
    else:
        posisi = "setara dengan"

    return (f"Realisasi Business Related Training {region} {posisi} average "
            f"Nasional ({reg_ach * 100:.0f}% vs {nat_ach * 100:.0f}%)")


# ---------------------------------------------------------------------------
# Slide 6 & 7 – Highlight/insight PA Sales & PA Collection
# ---------------------------------------------------------------------------
_PA_LABELS = {0: "2", 1: "3", 2: ">=4"}


def _pa_dominant(pa_row):
    """
    Dari satu baris `pa_breakdown()`/`_edu_pa()` (tuple `(label, pa2, pa3,
    pa4)`), cari kategori PA mana yang persentasenya paling tinggi.

    Return: (dominant_label_display str mis. "2"/"3"/">=4", dominant_value
    float, dominant_idx int 0/1/2) — dipakai oleh fungsi highlight slide 6/7
    untuk menyusun frasa "didominasi Avg. PA {label} ({value:.0%})".
    """
    values = pa_row[1:4]
    idx = max(range(3), key=lambda i: values[i])
    return _PA_LABELS[idx], values[idx], idx


def _pa_rows_all_dominant_same(pa_rows):
    """
    Cek apakah SEMUA baris di `pa_rows` (list hasil `pa_breakdown`) punya
    kategori PA dominan yang SAMA. Return label dominan (str) kalau ya,
    None kalau tidak seragam (dipakai untuk memilih frasa "serupa
    didominasi Avg. PA X" vs frasa fallback yang menyebutkan tiap
    kelompok terpisah bila polanya bervariasi).
    """
    if not pa_rows:
        return None
    labels = [_pa_dominant(r)[0] for r in pa_rows]
    return labels[0] if len(set(labels)) == 1 else None


def _pa_compare_to_national(region_row, national_row):
    """
    Membandingkan satu baris PA `region_row` terhadap baris nasional yang
    sepadan (`national_row`, label kategori sama) untuk menyimpulkan apakah
    distribusi PA region "lebih baik", "kurang baik", atau "setara" dibanding
    nasional.

    Definisi "lebih baik" dipakai di sini: kalau kategori dominan baris itu
    adalah Avg. PA 2 (performa rendah), region LEBIH BAIK kalau porsi PA
    2-nya LEBIH RENDAH dari nasional (berarti proporsi karyawan berkinerja
    rendah lebih sedikit). Kalau kategori dominannya Avg. PA >=4 (performa
    ideal), region LEBIH BAIK kalau porsinya LEBIH TINGGI dari nasional.
    Kalau dominannya Avg. PA 3 (tengah, ambigu baik/buruk), perbandingan
    tidak dilakukan (return None) supaya tidak mengarang kesimpulan kualitas
    yang tidak jelas arahnya.

    Return: "lebih baik" / "kurang baik" / "setara" / None (None berarti
    tidak bisa disimpulkan, mis. kategori dominan PA 3, atau data
    nasional/region tidak ada). Dipakai oleh get_slide6_highlights() dan
    get_slide7_highlights() untuk kalimat "Terhadap nasional, ...".
    """
    if region_row is None or national_row is None:
        return None
    _, reg_val, idx = _pa_dominant(region_row)
    nat_val = national_row[1 + idx]
    if idx == 2:      # Avg. PA >=4 — makin tinggi makin baik
        better = reg_val > nat_val
        worse = reg_val < nat_val
    elif idx == 0:    # Avg. PA 2 — makin rendah makin baik
        better = reg_val < nat_val
        worse = reg_val > nat_val
    else:             # Avg. PA 3 — ambigu, tidak disimpulkan
        return None
    if better:
        return "lebih baik"
    if worse:
        return "kurang baik"
    return "setara"


def _pa_los_narrative(pa_rows, count_rows, national_rows, function_label):
    """
    Menyusun kalimat slot LOS (masa kerja) untuk highlight slide 6/7 (mis.
    "Di Sales, performance PA masa kerja 1 tahun ke atas serupa didominasi
    Avg. PA >=4. Sedangkan masa kerja < 1 tahun masih didominasi Avg. PA 2
    (52%). Populasi SO/CO terbesar ada di 1 - 5 tahun (244 orang). Terhadap
    nasional, distribusi PA masa kerja < 1 tahun region {lebih/kurang} baik
    dari nasional.") berdasarkan `pa_rows`/`count_rows` region (indeks 0 =
    "a. <1 thn", sesuai LOS_ORDER_FRONTLINERS) dan `national_rows` yang
    sepadan.

    Parameter:
      pa_rows: data["s57"][region]["pa_sales_los"/"pa_coll_los"].
      count_rows: data["s57"][region]["count_sales_los"/"count_coll_los"].
      national_rows: data["s6_nat"]["LOS"] atau data["s7_nat"]["LOS"].
      function_label: "SO" (Sales Officer) atau "CO" (Collection Officer),
        dipakai di frasa "Populasi {label} terbesar".

    Return: str kalimat gabungan, atau "" kalau data kosong.
    """
    if not pa_rows or not count_rows:
        return ""
    new_hire, veterans = pa_rows[0], pa_rows[1:]

    parts = []
    dominant_all = _pa_rows_all_dominant_same(veterans)
    if dominant_all:
        parts.append(f"performance PA masa kerja 1 tahun ke atas serupa didominasi Avg. PA {dominant_all}")
    elif veterans:
        per_grup = ", ".join(f"{r[0]} didominasi Avg. PA {_pa_dominant(r)[0]}" for r in veterans)
        parts.append(f"performance PA masa kerja 1 tahun ke atas bervariasi ({per_grup})")

    nh_label, nh_val, _ = _pa_dominant(new_hire)
    parts.append(f"sedangkan masa kerja < 1 tahun masih didominasi Avg. PA {nh_label} ({nh_val:.0%})")

    max_label, max_count = max(count_rows, key=lambda r: r[1])
    parts.append(f"Populasi {function_label} terbesar ada di {max_label.split('. ', 1)[-1]} ({max_count} orang)")

    def cap_first(s):
        return s[0].upper() + s[1:] if s else s

    kalimat = f"{cap_first(parts[0])}. {cap_first(parts[1])}. {parts[2]}."

    national_new_hire = national_rows[0] if national_rows else None
    komparasi = _pa_compare_to_national(new_hire, national_new_hire)
    if komparasi:
        kalimat += f" Terhadap nasional, distribusi PA masa kerja < 1 tahun region {komparasi} dari nasional."
    return kalimat


def _pa_age_narrative(pa_rows, national_rows):
    """
    Menyusun kalimat slot usia untuk highlight slide 6/7 (mis. "Berdasarkan
    usia, distribusi PA ideal dominan pada usia 37 tahun ke atas, sedangkan
    kelompok usia di bawah 36 tahun masih perlu peningkatan. Terhadap
    nasional, distribusi PA secara usia di region {lebih/kurang} baik
    dibanding angka nasional.").

    Cara kerja: AGE_ORDER dibagi jadi 2 kelompok — "muda" (a. <26 thn, b.
    26-36 thn) dan "37 ke atas" (c, d, e) — lalu dicek apakah kelompok
    "37 ke atas" mayoritas didominasi Avg. PA >=4 (ideal) sementara
    kelompok muda tidak. Perbandingan ke nasional memakai rata-rata pa4%
    seluruh kelompok usia (bukan cuma satu kelompok, beda dari slot LOS
    yang fokus ke "<1 thn" saja, karena contoh slot ini bicara "secara
    usia" secara umum, bukan satu kelompok spesifik).

    Parameter: pa_rows = data["s57"][region]["pa_sales_age"/"pa_coll_age"],
    national_rows = data["s6_nat"]["AGE"]/data["s7_nat"]["AGE"].

    Return: str, atau "" kalau data kosong.
    """
    if not pa_rows:
        return ""
    muda, tua = pa_rows[:2], pa_rows[2:]
    tua_ideal = tua and all(_pa_dominant(r)[2] == 2 for r in tua)
    muda_belum_ideal = muda and not all(_pa_dominant(r)[2] == 2 for r in muda)

    if tua_ideal and muda_belum_ideal:
        kalimat = ("Berdasarkan usia, distribusi PA ideal dominan pada usia 37 tahun ke atas, "
                   "sedangkan kelompok usia di bawah 36 tahun masih perlu peningkatan.")
    elif tua_ideal:
        kalimat = "Berdasarkan usia, distribusi PA ideal dominan pada usia 37 tahun ke atas."
    else:
        kalimat = "Berdasarkan usia, distribusi PA bervariasi di tiap kelompok usia."

    if national_rows and len(national_rows) == len(pa_rows):
        reg_avg = sum(_pa_dominant(r)[1] if _pa_dominant(r)[2] == 2 else r[3] for r in pa_rows) / len(pa_rows)
        nat_avg = sum(r[3] for r in national_rows) / len(national_rows)
        if reg_avg > nat_avg:
            komparasi = "lebih baik dibanding"
        elif reg_avg < nat_avg:
            komparasi = "kurang baik dibanding"
        else:
            komparasi = "setara dengan"
        kalimat += f" Terhadap nasional, distribusi PA secara usia di region {komparasi} angka nasional."
    return kalimat


def _pa_edu_narrative(pa_rows, national_rows):
    """
    Menyusun kalimat slot pendidikan untuk highlight slide 6/7 (mis.
    "Secara Pendidikan, proporsi PA Pendidikan SMA tidak lebih baik dari
    proporsi PA Nasional. Terhadap nasional, distribusi PA Pendidikan
    Diploma region lebih baik dari angka nasional.").

    Cara kerja: untuk tiap kategori pendidikan (SLTA/Diploma/Sarjana/Pasca
    Sarjana), bandingkan porsi Avg. PA >=4 region vs nasional (>=4 dipilih
    sebagai satu-satunya metrik pembanding lintas kategori, supaya "lebih
    baik" punya arah yang konsisten — makin tinggi porsi >=4 makin baik).
    Kategori dengan selisih PALING NEGATIF (region kalah paling jauh) dan
    PALING POSITIF (region unggul paling jauh) masing-masing disebut satu
    kalimat, meniru pola contoh (satu kategori "tidak lebih baik", satu
    kategori lain "lebih baik").

    Parameter: pa_rows = data["s57"][region]["pa_sales_edu"/"pa_coll_edu"],
    national_rows = data["s6_nat"]["EDU"]/data["s7_nat"]["EDU"].

    Return: str, atau "" kalau data kosong/tidak sepadan.
    """
    if not pa_rows or not national_rows or len(pa_rows) != len(national_rows):
        return ""
    edu_display = {
        "1. SLTA sederajat & di bawahnya": "SMA",
        "2. Diploma": "Diploma",
        "3. Sarjana": "Sarjana",
        "4. Pasca Sarjana": "Pasca Sarjana",
    }
    selisih = []
    for reg_row, nat_row in zip(pa_rows, national_rows):
        label = edu_display.get(reg_row[0], reg_row[0])
        selisih.append((label, reg_row[3] - nat_row[3]))

    terburuk = min(selisih, key=lambda x: x[1])
    terbaik = max(selisih, key=lambda x: x[1])
    if terburuk[1] >= 0 and terbaik[1] >= 0:
        return (f"Secara Pendidikan, seluruh kategori pendidikan region setara atau lebih baik "
                f"dari proporsi PA Nasional, dengan {terbaik[0]} unggul paling jauh.")

    kalimat = f"Secara Pendidikan, proporsi PA Pendidikan {terburuk[0]} tidak lebih baik dari proporsi PA Nasional."
    if terbaik[0] != terburuk[0] and terbaik[1] > 0:
        kalimat += f" Terhadap nasional, distribusi PA Pendidikan {terbaik[0]} region lebih baik dari angka nasional."
    return kalimat


def get_slide6_highlights(data, region):
    """
    Membuat draf teks highlight/insight otomatis untuk slide 6 — PA Sales
    (3 slot: {SLIDE6_HIGHLIGHT_A/B/C} = LOS/usia/pendidikan), seluruhnya
    dihitung dari `data["s57"][region]` (angka region) dan `data["s6_nat"]`
    (angka nasional Sales), lewat `_pa_los_narrative`/`_pa_age_narrative`/
    `_pa_edu_narrative` (dipakai bersama slide 7, lihat docstring
    masing-masing untuk logikanya).

    Parameter: data (dict load_all()), region (str).

    Return: dict {"A": str, "B": str, "C": str}. String kosong kalau data
    dasar tidak ditemukan. Dipanggil dari _apply_slide6_highlights() di
    pptx_updater.py.
    """
    s57 = data.get("s57", {}).get(region, {})
    nat = data.get("s6_nat", {})
    return {
        "A": _pa_los_narrative(s57.get("pa_sales_los", []), s57.get("count_sales_los", []),
                                nat.get("LOS", []), "SO"),
        "B": _pa_age_narrative(s57.get("pa_sales_age", []), nat.get("AGE", [])),
        "C": _pa_edu_narrative(s57.get("pa_sales_edu", []), nat.get("EDU", [])),
    }


def get_slide7_highlights(data, region):
    """
    Sama seperti get_slide6_highlights(), tapi untuk slide 7 — PA
    Collection (3 slot: {SLIDE7_HIGHLIGHT_A/B/C}), memakai
    `pa_coll_los`/`pa_coll_age`/`pa_coll_edu`/`count_coll_los` dan
    `data["s7_nat"]` (nasional Collection) sebagai gantinya.

    Return: dict {"A": str, "B": str, "C": str}. Dipanggil dari
    _apply_slide7_highlights() di pptx_updater.py.
    """
    s57 = data.get("s57", {}).get(region, {})
    nat = data.get("s7_nat", {})
    return {
        "A": _pa_los_narrative(s57.get("pa_coll_los", []), s57.get("count_coll_los", []),
                                nat.get("LOS", []), "CO"),
        "B": _pa_age_narrative(s57.get("pa_coll_age", []), nat.get("AGE", [])),
        "C": _pa_edu_narrative(s57.get("pa_coll_edu", []), nat.get("EDU", [])),
    }


def get_slide14_highlights(data, region):
    """
    Membuat draf teks highlight/insight otomatis untuk slide 14 (4 slot:
    {SLIDE14_HIGHLIGHT_A..D}), berdasarkan `data["attrition"][region]`
    (hasil `load_attrition()` di data_loader.py, termasuk `reason_out_rows`
    dari `load_slide14_reason_out()`) dan `data["attrition"]["Nasional"]`
    sebagai pembanding.

    Cara kerja tiap slot:
      A — Nama region saja (label/judul kotak, bukan kalimat insight),
          persis seperti baris pertama contoh di Template_Highlight.txt
          ("Jawa Tengah" berdiri sendiri sebelum paragraf insight-nya).
      B — Arah tren YoY (naik/turun) utuk 3 metrik (non regret, regret,
          total), dibandingkan arah tren nasional untuk metrik yang sama —
          disebut "mengikuti trend nasional" kalau SEMUA metrik region
          searah dengan nasional, atau menyebutkan metrik mana yang
          berlawanan arah kalau tidak.
      C — Reason involuntary non-regret dengan porsi TERBESAR (bisa lebih
          dari satu kalau ada yang seri/tied), diambil dari baris-baris
          `reason_out_rows` berkategori "Involuntary" (bukan header).
      D — Reason voluntary dengan porsi TOTAL terbesar ("tertinggi"), dan
          terpisah, reason voluntary dengan porsi REGRET terbesar (dua
          hal ini BISA berbeda, seperti pada contoh: tertinggi karena
          "better job" tapi regret-nya karena "start business").

    Parameter: data (dict load_all()), region (str).

    Return: dict {"A": str, "B": str, "C": str, "D": str}. String kosong
    kalau data dasar tidak ditemukan. Dipanggil dari
    `_apply_slide14_highlights()` di pptx_updater.py.
    """
    attr = data.get("attrition", {}).get(region, {})
    nat = data.get("attrition", {}).get("Nasional", {})
    result = {"A": region, "B": "", "C": "", "D": ""}
    if not attr:
        return result

    # --- B: arah tren YoY vs nasional ---
    metrik_label = {"non_regret": "non regret", "regret": "regret", "total": "total attrition"}
    arah_region, arah_nasional, berlawanan = {}, {}, []
    for key in ("non_regret", "regret", "total"):
        r25, r26 = attr.get(f"yoy_{key}_25"), attr.get(f"yoy_{key}_26")
        n25, n26 = nat.get(f"yoy_{key}_25"), nat.get(f"yoy_{key}_26")
        if None in (r25, r26, n25, n26):
            continue
        arah_region[key] = "turun" if r26 < r25 else ("naik" if r26 > r25 else "tetap")
        arah_nasional[key] = "turun" if n26 < n25 else ("naik" if n26 > n25 else "tetap")
        if arah_region[key] != arah_nasional[key]:
            berlawanan.append(metrik_label[key])

    _ARAH_NOUN = {"turun": "penurunan", "naik": "kenaikan", "tetap": "kestabilan"}
    if arah_region:
        arah_total = arah_region.get("total")
        noun_total = _ARAH_NOUN.get(arah_total)
        if not berlawanan and noun_total:
            result["B"] = (f"{region} mengikuti trend nasional di mana terjadi {noun_total} "
                            f"attrition di semua lini baik di non regret, regret, dan total attrition")
        elif noun_total:
            result["B"] = (f"Secara total attrition, {region} mengalami {noun_total} mengikuti "
                            f"trend nasional, kecuali pada {_join_indonesian(berlawanan)} yang "
                            f"berlawanan arah dengan nasional")

    rows = attr.get("reason_out_rows", [])

    def top_labels(pool, key):
        if not pool:
            return []
        best = max(r[key] for r in pool)
        if best <= 0:
            return []
        return [r["label"] for r in pool if r[key] == best]

    # --- C: reason involuntary non-regret terbesar ---
    inv_idx = next((i for i, r in enumerate(rows) if r["label"] == "Involuntary"), None)
    if inv_idx is not None:
        next_header = next((i for i in range(inv_idx + 1, len(rows)) if rows[i]["is_header"]), len(rows))
        inv_reasons = [r for r in rows[inv_idx + 1:next_header] if not r["is_header"]]
        top_nr = top_labels(inv_reasons, "non_regret")
        if top_nr:
            result["C"] = (f"Secara reason out, involuntary non regret masih karena terminasi "
                            f"{_join_indonesian(top_nr)}")

    # --- D: reason voluntary tertinggi (total) vs regret tertinggi ---
    vol_idx = next((i for i, r in enumerate(rows) if r["label"] == "Voluntary"), None)
    if vol_idx is not None:
        next_header = next((i for i in range(vol_idx + 1, len(rows)) if rows[i]["is_header"]), len(rows))
        vol_reasons = [r for r in rows[vol_idx + 1:next_header] if not r["is_header"]]
        top_total = top_labels(vol_reasons, "total")
        top_regret = top_labels(vol_reasons, "regret")
        if top_total and top_regret:
            if set(top_total) == set(top_regret):
                result["D"] = (f"Reason out voluntary tertinggi maupun voluntary regret "
                                f"sama-sama karena {_join_indonesian(top_total)}")
            else:
                result["D"] = (f"Reason out voluntary tertinggi karena {_join_indonesian(top_total)} "
                                f"tetapi voluntary regret karena {_join_indonesian(top_regret)}")

    return result


def _fraud_branch_base_name(label):
    """Ambil nama lokasi dasar dari label branch/cluster fraud (mis. "Salatiga-Osamaliki" -> "Salatiga", "Tegal " -> "Tegal") dengan memotong bagian setelah "-" dan men-strip spasi, supaya bisa dicocokkan antara daftar worst5 branch dan worst5 cluster yang format namanya sedikit beda."""
    return label.split("-", 1)[0].strip()


def get_slide18_highlights(data, region):
    """
    Membuat draf teks highlight/insight otomatis untuk slide 18 (fraud):
    1 token angka {HIGHLIGHT_FRAUDRATE_YOY} + 4 slot
    {SLIDE18_HIGHLIGHT_A..D}, seluruhnya dari `data["fraud"][region]`
    (hasil `load_slide18_fraud()` di data_loader.py).

    Cara kerja tiap slot:
      FRAUDRATE_YOY — Kalimat ringkasan: arah & besar perubahan YoY total
          potential loss (potloss_25 -> potloss_26), plus 1-2 lokasi
          "kasus berat" (branch dengan potential loss TERBESAR di
          worst5_branch_25 — proxy data untuk "kasus berat", karena tidak
          ada kolom jumlah-kasus per lokasi untuk mendeteksi lokasi dengan
          >1 kasus).
      A — Lokasi (Branch SSD vs Cluster Collection) mana yang jadi
          pendorong utama kenaikan/penurunan, dari `by_location_25/26`
          (dibandingkan lewat SELISIH absolut, bukan %, supaya kontribusi
          nominal besar tidak kalah oleh basis kecil yang %-nya kebetulan
          tinggi).
      B — Apakah kenaikan/penurunan terjadi di kedua sisi Internal maupun
          Eksternal (`by_ie_25/26`), atau cuma salah satu.
      C — 2 jenis fraud (`by_kasus_25/26`) dengan selisih absolut potloss
          YoY TERBESAR (sumber utama kenaikan/penurunan).
      D — Lokasi yang muncul di KEDUA daftar worst5 (`worst5_branch_25`
          dan `worst5_cluster_25`), dicocokkan lewat nama dasarnya
          (`_fraud_branch_base_name`, memotong akhiran "-Cabang").

    Parameter: data (dict load_all()), region (str).

    Return: dict {"FRAUDRATE_YOY": str, "A": str, "B": str, "C": str,
    "D": str}. String kosong kalau data dasar tidak ditemukan. Dipanggil
    dari `_apply_slide18_highlights()` di pptx_updater.py.
    """
    fraud = data.get("fraud", {}).get(region, {})
    result = {"FRAUDRATE_YOY": "", "A": "", "B": "", "C": "", "D": ""}
    if not fraud:
        return result

    p25, p26 = fraud.get("potloss_25"), fraud.get("potloss_26")
    if p25 is not None and p26 is not None:
        arah = "kenaikan" if p26 >= p25 else "penurunan"
        pct = abs(p26 - p25) / p25 * 100 if p25 else 0.0
        worst5 = fraud.get("worst5_branch_25", [])[:2]
        lokasi_berat = _join_indonesian([_fraud_branch_base_name(b[0]) for b in worst5]) if worst5 else ""
        kalimat = f"Fraud Rate {region} mengalami {arah} sebesar {pct:.0f}% dibanding tahun lalu"
        if lokasi_berat:
            kalimat += f", dengan kasus berat di {lokasi_berat}"
        result["FRAUDRATE_YOY"] = kalimat

    # --- A: lokasi pendorong utama (Branch SSD vs Cluster Collection) ---
    loc25, loc26 = fraud.get("by_location_25", {}), fraud.get("by_location_26", {})
    if loc25 and loc26:
        selisih_loc = {k: loc26.get(k, 0) - loc25.get(k, 0) for k in loc25}
        top_loc = max(selisih_loc, key=lambda k: abs(selisih_loc[k]))
        arah_loc = "Kenaikan" if selisih_loc[top_loc] >= 0 else "Penurunan"
        result["A"] = f"{arah_loc} fraud ini paling besar terjadi di {top_loc}"

    # --- B: internal vs eksternal ---
    ie25, ie26 = fraud.get("by_ie_25", {}), fraud.get("by_ie_26", {})
    if ie25 and ie26:
        naik_int = ie26.get("Internal", 0) >= ie25.get("Internal", 0)
        naik_ext = ie26.get("External", 0) >= ie25.get("External", 0)
        if naik_int and naik_ext:
            result["B"] = "Terjadi kenaikan fraud baik oleh Internal maupun Eksternal"
        elif not naik_int and not naik_ext:
            result["B"] = "Terjadi penurunan fraud baik oleh Internal maupun Eksternal"
        elif naik_int:
            result["B"] = "Kenaikan fraud terjadi oleh pihak Internal, sedangkan Eksternal menurun"
        else:
            result["B"] = "Kenaikan fraud terjadi oleh pihak Eksternal, sedangkan Internal menurun"

    # --- C: 2 jenis fraud dengan selisih YoY terbesar ---
    kasus25 = dict(fraud.get("by_kasus_25", []))
    kasus26 = dict(fraud.get("by_kasus_26", []))
    if kasus25 or kasus26:
        semua_jenis = set(kasus25) | set(kasus26)
        selisih_kasus = [(j, kasus26.get(j, 0) - kasus25.get(j, 0)) for j in semua_jenis]
        selisih_kasus.sort(key=lambda x: -x[1])
        top2 = [j for j, d in selisih_kasus[:2] if d > 0]
        if top2:
            result["C"] = f"Sumber kenaikannya dari jenis fraud {_join_indonesian(top2)}"

    # --- D: lokasi yang masuk worst5 branch maupun worst5 cluster ---
    branch_bases = {_fraud_branch_base_name(b[0]) for b in fraud.get("worst5_branch_25", [])}
    cluster_bases = {_fraud_branch_base_name(c[0]) for c in fraud.get("worst5_cluster_25", [])}
    overlap = sorted(branch_bases & cluster_bases)
    if overlap:
        result["D"] = (f"{_join_indonesian(overlap)} menjadi highlight karena masuk fraud worst 5 "
                        f"baik di Branch SSD maupun Cluster Collection")

    return result


def _branch_city_name(label):
    """Ambil nama kota dasar dari label branch/cluster slide 16 (mis. "PURWOKERTO-JEND. SUDIRMAN" -> "purwokerto", "Pekalongan" -> "pekalongan"), huruf kecil semua supaya pencocokan branch (Sales) vs cluster (Collection) case-insensitive dan tidak peduli detail alamat setelah tanda "-"."""
    return label.split("-", 1)[0].strip().lower()


def get_slide16_highlights(data, region):
    """
    Membuat draf teks insight otomatis untuk slide 16 (1 slot:
    {SLIDE16_INSIGHT}, digabung jadi satu paragraf berisi 2 kalimat —
    sama seperti pola slot gabungan di slide 6/7 — karena hanya ada 1
    placeholder yang tersedia di template untuk 2 poin pada contoh di
    Template_Highlight.txt), berdasarkan `data["s16"][region]`.

    Cara kerja tiap kalimat:
      1 — Cari kota yang MASUK top-5 pct_rg (%Regret) di KEDUA daftar
          `branch_sales` (granularitas branch, nama kota diambil dari
          bagian sebelum "-") dan `branch_collection` (granularitas
          cluster/kota) — kota yang beririsan di top-5 kedua sisi
          dianggap "beririsan regretted attrition". Kalau lebih dari
          satu kota beririsan, disebutkan semua.
      2 — Reason voluntary regret dengan jumlah TERBESAR, dicari terpisah
          untuk Sales (`sales_officer`) dan Collection (`collection_officer`)
          dari `data["s16"][region]["reason_out"]` (baris non-header —
          list ini sudah representasi alasan REGRET saja, lihat baris
          "Total Regret" di baris header terakhirnya).

    Parameter: data (dict load_all()), region (str).

    Return: str (bisa "" kalau data kosong). Dipanggil dari
    `_apply_slide16_highlight()` di pptx_updater.py.
    """
    s16 = data.get("s16", {}).get(region, {})
    if not s16:
        return ""

    kalimat = []

    branch_sales = s16.get("branch_sales", [])
    branch_coll = s16.get("branch_collection", [])
    top_sales_city = {_branch_city_name(r["label"]) for r in branch_sales[:5]}
    top_coll_city = {_branch_city_name(r["label"]) for r in branch_coll[:5]}
    overlap = top_sales_city & top_coll_city
    if overlap:
        nama = _join_indonesian([c.title() for c in sorted(overlap)])
        kalimat.append(f"Branch dan Cluster beririsan regretted attrition di {nama}")

    reason_rows = [r for r in s16.get("reason_out", []) if not r["is_header"]]
    if reason_rows:
        top_sales = max(reason_rows, key=lambda r: r["sales_officer"])
        top_coll = max(reason_rows, key=lambda r: r["collection_officer"])
        if top_sales["sales_officer"] > 0 or top_coll["collection_officer"] > 0:
            bagian = []
            if top_sales["sales_officer"] > 0:
                bagian.append(f"Sales karena {top_sales['label']}")
            if top_coll["collection_officer"] > 0:
                bagian.append(f"Collection karena {top_coll['label']}")
            kalimat.append(f"Alasan voluntary regret tertinggi di {_join_indonesian(bagian)}")

    return ". ".join(kalimat) + ("." if kalimat else "")


def get_slide15_highlights(data, region):
    """
    Membuat draf teks highlight/insight otomatis untuk slide 15 (3 slot:
    {SLIDE15_HIGHLIGHT_A/B/C}), berdasarkan `data["s15_func"][region]`
    ("y26"/"y25", hasil `load_slide15_function()` — lihat `_SLIDE15_ROWS`
    untuk urutan barisnya, baris terakhir "Grand Total").

    Cara kerja tiap slot (hanya 2 dari 4 poin di Template_Highlight.txt
    yang benar-benar informatif & bisa dihitung dari data — 1 poin lain di
    contoh murni mengulang kesimpulan Grand Total di level function tanpa
    tambahan info, dan 1 poin lagi murni definisi organisasi ["Support
    terdiri dari fungsi back office dan SH"] yang tidak ada di data mana
    pun, jadi TIDAK dipaksakan jadi slot ke-4/5 — cukup ditampung di 3
    slot yang sudah ada):
      A — Arah tren Grand Total YoY (naik/turun, utuk non regret/regret/
          total sekaligus) + penyebabnya (perubahan jumlah keluar
          `out_total` dan perubahan `avg_hc`), digabung satu paragraf.
      B — Fungsi individual (bukan Grand Total) yang arah tren TOTAL
          attrition-nya BERLAWANAN dari Grand Total (mis. Grand Total
          turun tapi fungsi tsb naik) — kalau tidak ada, dicek juga
          fungsi yang regret-nya naik meski totalnya turun (kasus
          "regret naik walau total turun", meniru pola contoh "Sales
          Support naik total, Collection Support naik regret-nya saja").
      C — SENGAJA DIKOSONGKAN (lihat catatan di atas soal definisi
          organisasi "Support = back office + SH").

    Parameter: data (dict load_all()), region (str).

    Return: dict {"A": str, "B": str, "C": ""}. Dipanggil dari
    `_apply_slide15_highlights()` di pptx_updater.py.
    """
    s15 = data.get("s15_func", {}).get(region, {})
    y26, y25 = s15.get("y26", []), s15.get("y25", [])
    result = {"A": "", "B": "", "C": ""}
    if not y26 or not y25 or len(y26) != len(y25):
        return result

    gt26 = next((r for r in y26 if r["label"] == "Grand Total"), None)
    gt25 = next((r for r in y25 if r["label"] == "Grand Total"), None)
    if gt26 and gt25:
        arah = "penurunan" if gt26["pct_total"] < gt25["pct_total"] else (
            "kenaikan" if gt26["pct_total"] > gt25["pct_total"] else "kestabilan")
        arah_nr = "turun" if gt26["pct_nr"] < gt25["pct_nr"] else ("naik" if gt26["pct_nr"] > gt25["pct_nr"] else "tetap")
        arah_rg = "turun" if gt26["pct_rg"] < gt25["pct_rg"] else ("naik" if gt26["pct_rg"] > gt25["pct_rg"] else "tetap")
        result["A"] = (f"Secara YoY Attrition, ada {arah} attrition baik dari segi non regret "
                        f"({arah_nr}) maupun regret ({arah_rg})")
        sebab = []
        if gt26["out_total"] != gt25["out_total"]:
            arah_keluar = "menurunnya" if gt26["out_total"] < gt25["out_total"] else "meningkatnya"
            sebab.append(f"{arah_keluar} jumlah orang keluar ({gt25['out_total']} -> {gt26['out_total']})")
        if gt26["avg_hc"] != gt25["avg_hc"]:
            arah_hc = "meningkatnya" if gt26["avg_hc"] > gt25["avg_hc"] else "menurunnya"
            sebab.append(f"{arah_hc} headcount aktif sebagai pembagi ({gt25['avg_hc']:.0f} -> {gt26['avg_hc']:.0f})")
        if sebab:
            result["A"] += f", disebabkan karena {' dan '.join(sebab)}"
        result["A"] += "."

    if gt26 and gt25:
        arah_grand = "turun" if gt26["pct_total"] < gt25["pct_total"] else ("naik" if gt26["pct_total"] > gt25["pct_total"] else "tetap")
        anomali = []
        for r26, r25 in zip(y26, y25):
            if r26["label"] == "Grand Total":
                continue
            arah_fn_total = "turun" if r26["pct_total"] < r25["pct_total"] else ("naik" if r26["pct_total"] > r25["pct_total"] else "tetap")
            if arah_fn_total != arah_grand:
                anomali.append(f"{r26['label']} ({arah_fn_total} total attrition)")
                continue
            arah_fn_rg = "turun" if r26["pct_rg"] < r25["pct_rg"] else ("naik" if r26["pct_rg"] > r25["pct_rg"] else "tetap")
            arah_grand_rg = "turun" if gt26["pct_rg"] < gt25["pct_rg"] else ("naik" if gt26["pct_rg"] > gt25["pct_rg"] else "tetap")
            if arah_fn_rg != arah_grand_rg and arah_fn_rg != "tetap":
                anomali.append(f"{r26['label']} ({arah_fn_rg} attrition regret meski total {arah_fn_total})")
        if anomali:
            result["B"] = (f"Berbeda dari tren Grand Total yang {arah_grand}, "
                            f"{_join_indonesian(anomali)} bergerak berlawanan arah")

    return result


def _pct_of_total(entries, idx, series_idx):
    """
    Menghitung persentase satu entri terhadap total seluruh entri pada list
    yang sama — dipakai untuk data NASIONAL slide 5, yang hanya tersedia
    sebagai angka mentah (jumlah orang) per kategori, bukan persentase yang
    sudah dihitung (beda dengan data REGIONAL yang sudah punya
    `*_sales_pct`/`*_coll_pct` dari data_loader.py).

    Parameter:
        entries: list tuple `(label, val_sales, val_coll)`, mis.
                 `data["s5_nat"]["LOS"]`.
        idx: index kategori yang mau dihitung persentasenya (0-based,
             sesuai urutan `entries`).
        series_idx: 1 untuk Field Sales, 2 untuk Field Coll (posisi kolom
                    pada tuple `entries`).

    Return: float persentase (0.0-1.0), atau None jika data tidak cukup
    (list kosong/index di luar jangkauan/total 0). Dipanggil oleh
    `get_slide5_pct_boxes()` di bawah.
    """
    if not entries or idx >= len(entries):
        return None
    total = sum(e[series_idx] for e in entries if e[series_idx] is not None)
    if not total:
        return None
    v = entries[idx][series_idx]
    return (v / total) if v is not None else None


def _bandingkan_pct(reg_pct, nat_pct, toleransi=0.03):
    """
    Bandingkan `reg_pct` vs `nat_pct` (keduanya fraksi 0-1), return frasa
    Indonesia "lebih tinggi dari"/"lebih rendah dari"/"setara dengan"
    (kalau selisihnya di dalam `toleransi`, default 3 poin persentase).
    Dipakai berulang oleh get_slide4_highlights() untuk membandingkan
    berbagai metrik regional vs nasional dengan bahasa yang konsisten.
    """
    selisih = reg_pct - nat_pct
    if abs(selisih) <= toleransi:
        return "setara dengan"
    return "lebih tinggi dari" if selisih > 0 else "lebih rendah dari"


def _fmt_slide4_pct(pct):
    """
    Format persentase (0-100, sudah dikali 100) ala template slide 4: bulat
    ke integer terdekat (round-half-up) untuk nilai >= 1%, TAPI pakai 1
    desimal berkoma Indonesia (mis. "0,2%") untuk nilai kecil di bawah 1%
    yang kalau dibulatkan ke integer akan hilang jadi "0%" — meniru gaya
    tampilan template asli (mis. Pasca Sarjana "0,2%"/"0,3%" di tabel EDU).
    """
    if pct > 0 and round(pct) == 0:
        return f"{pct:.1f}".replace(".", ",") + "%"
    return f"{_round_half_up(pct)}%"


def get_slide4_lea_pct(data, region):
    """
    Menghitung persentase AGE/EDU/LOS (region vs nasional) untuk 3 tabel
    statis di slide 4 (di bawah chart10/11/12 histogram) yang SEBELUMNYA
    TIDAK PERNAH DISENTUH kode apa pun — selalu menampilkan nilai contoh
    statis milik region demo template (Jawa Tengah) untuk SEMUA region.

    PENTING soal urutan kategori LOS: tabel LOS di template menyusun 6
    barisnya dalam urutan MENAIK (a. <1 thn -> f. > 20th) — KEBALIKAN dari
    `s4_lea[region]["LOS"]` yang urutannya MENURUN (LOS_ORDER, f -> a,
    dipakai chart10 histogram bar). Kalau urutan ini tidak dibalik saat
    mengisi tabel, persentase tiap baris akan tertukar dengan baris
    tetangganya (baris "a. <1 thn" akan menerima angka milik "f. > 20th",
    dst.) — makanya list LOS di sini SENGAJA di-reverse sebelum dipakai.

    Sumber data:
      - Region: `data["s4_lea"][region]["AGE"/"EDU"/"LOS"]` (list of
        (label, count), hasil `load_slide4_los_edu_age()`).
      - Nasional: dihitung DI SINI sebagai jumlah count SELURUH 12 region
        di `data["s4_lea"]` (bukan rata-rata seperti `s4_fl`/`s4_soc` —
        `s4_lea` adalah agregasi individual karyawan per region, jadi
        menjumlahkannya menghasilkan TOTAL PERUSAHAAN yang valid, bukan
        estimasi/rata-rata).

    Parameter: data (dict load_all()), region (str).

    Return: dict {
        "AGE": {"reg": [pct×5, urutan AGE_ORDER menaik], "nas": [pct×5]},
        "EDU": {"reg": [pct×4, urutan EDU_ORDER menaik], "nas": [pct×4]},
        "LOS": {"reg": [pct×6, urutan a->f MENAIK], "nas": [pct×6]},
    } — tiap pct adalah angka 0-100 (sudah dikali 100, BUKAN fraksi 0-1),
    siap diformat lewat `_fmt_slide4_pct`. List kosong kalau data tidak
    ditemukan. Dipanggil dari `_apply_slide4_lea_tables()` di
    pptx_updater.py.
    """
    lea = data.get("s4_lea", {})
    reg_lea = lea.get(region, {})
    result = {"AGE": {"reg": [], "nas": []}, "EDU": {"reg": [], "nas": []}, "LOS": {"reg": [], "nas": []}}
    if not reg_lea:
        return result

    def pct_list(rows):
        total = sum(cnt for _, cnt in rows)
        if not total:
            return [0.0] * len(rows)
        return [cnt / total * 100 for _, cnt in rows]

    def national_pct_list(key, label_order):
        agg = {}
        for v in lea.values():
            for lbl, cnt in v.get(key, []):
                agg[lbl] = agg.get(lbl, 0) + cnt
        total = sum(agg.values())
        if not total:
            return [0.0] * len(label_order)
        return [agg.get(lbl, 0) / total * 100 for lbl in label_order]

    age_rows = reg_lea.get("AGE", [])
    result["AGE"]["reg"] = pct_list(age_rows)
    result["AGE"]["nas"] = national_pct_list("AGE", [lbl for lbl, _ in age_rows])

    edu_rows = reg_lea.get("EDU", [])
    result["EDU"]["reg"] = pct_list(edu_rows)
    result["EDU"]["nas"] = national_pct_list("EDU", [lbl for lbl, _ in edu_rows])

    los_rows_asc = list(reversed(reg_lea.get("LOS", [])))  # tabel: a->f menaik, s4_lea: f->a menurun
    result["LOS"]["reg"] = pct_list(los_rows_asc)
    result["LOS"]["nas"] = national_pct_list("LOS", [lbl for lbl, _ in los_rows_asc])

    return result


def get_slide4_highlights(data, region):
    """
    Membuat draf teks highlight/insight otomatis untuk slide 4 (3 slot:
    {SLIDE4_HIGHLIGHT_A/B/C}), berdasarkan `data["s4_npat"]`, `data["s4_wc"]`,
    `data["s4_func"]`, `data["s4_fl"]`, `data["s4_soc"]`, `data["s4_lea"]`
    (masing-masing hasil loader terpisah, lihat data_loader.py).

    10 poin di Template_Highlight.txt (dikurangi 1 poin murni naratif
    bisnis ["Proporsi headcount dipengaruhi oleh migrasi 9 cab HI ..."]
    yang TIDAK ada di data manapun, jadi tidak dipaksakan) digabung jadi
    3 slot (jumlah placeholder yang tersedia di template):
      A — Tren HC Managed vs NPAT/HC 3 tahun terakhir (dari
          `s4_npat[region]["FY23/FY24/FY25"]["NPAT_HC"]`, mendeteksi
          pola turun-lalu-naik atau arah lain apa adanya) + proporsi Non
          Permanent (`s4_wc`, `pct_contract+pct_os`) dibanding nasional
          (`nas_contract+nas_os`).
      B — Fungsi dengan proporsi headcount TERBESAR (`s4_func`, bisa lebih
          dari satu kalau seri) + proporsi Frontliners Staff terhadap
          total HC (`s4_fl`, dibanding total company HC di
          `s4_npat[region]["5M2026"]["HC"]`) dibanding rata-rata nasional
          per region yang sepadan (`nas_sales_fl+nas_coll_fl` sebagai
          proksi, karena tidak ada angka nasional absolut — lihat catatan
          "rata-rata nasional" di docstring `load_slide4_frontliners`).
      C — Span of Control Field Sales & Field Coll (`s4_soc`, region vs
          `s4_soc["NAS"]`, dibandingkan lewat `_bandingkan_pct` dengan
          basis rasio bukan persentase) + profil demografi keseluruhan
          (BUKAN cuma frontliners — dari `s4_lea`, LOS/EDU/AGE per
          karyawan): proporsi masa kerja < 5 tahun (dibanding rata-rata
          nasional, dihitung sebagai TOTAL SEMUA REGION di `s4_lea` —
          ini genuinely representasi total perusahaan, bukan rata-rata
          per region, karena `s4_lea` adalah agregasi individual
          karyawan, bukan sheet ringkasan seperti `s4_fl`), proporsi
          Sarjana & Diploma, dan proporsi usia 26-36 tahun.

    Parameter: data (dict load_all()), region (str).

    Return: dict {"A": str, "B": str, "C": str}. Dipanggil dari
    `_apply_slide4_highlights()` di pptx_updater.py.
    """
    result = {"A": "", "B": "", "C": ""}

    # --- A: tren NPAT/HC + Non Permanent % ---
    npat = data.get("s4_npat", {}).get(region, {})
    bagian_a = []
    tahun_urut = [y for y in ("FY23", "FY24", "FY25") if y in npat]
    if len(tahun_urut) >= 2:
        nilai = [npat[y]["NPAT_HC"] for y in tahun_urut]
        arah = []
        for i in range(1, len(nilai)):
            arah.append("naik" if nilai[i] > nilai[i - 1] else ("turun" if nilai[i] < nilai[i - 1] else "tetap"))
        if len(set(arah)) > 1:
            bagian_a.append(f"Ada tren yang fluktuatif antara HC Managed terhadap NPAT ({' lalu '.join(arah)} "
                             f"dari {tahun_urut[0]} ke {tahun_urut[-1]})")
        else:
            bagian_a.append(f"NPAT terhadap HC Managed konsisten {arah[0]} dari {tahun_urut[0]} ke {tahun_urut[-1]}")

    wc = data.get("s4_wc", {}).get(region, {})
    if wc:
        non_perm = wc.get("pct_contract", 0) + wc.get("pct_os", 0)
        nas_non_perm = wc.get("nas_contract", 0) + wc.get("nas_os", 0)
        komparasi = _bandingkan_pct(non_perm, nas_non_perm)
        bagian_a.append(f"{non_perm:.0%} proporsi Non Permanent (contract & outsource) dari total karyawan, "
                         f"Proporsi Non Permanen Regional {komparasi} Nasional ({non_perm:.0%} vs {nas_non_perm:.0%})")
    if bagian_a:
        result["A"] = ". ".join(bagian_a) + "."

    # --- B: fungsi terbesar + proporsi frontliners ---
    func = dict(data.get("s4_func", {}).get(region, {}))
    func.pop("Grand Total", None)
    func = {k: v for k, v in func.items() if v is not None}
    bagian_b = []
    if func:
        maks = max(func.values())
        top_func = sorted([k for k, v in func.items() if v == maks])
        bagian_b.append(f"Proporsi karyawan terbesar ada di {_join_indonesian(top_func)}")

    fl = data.get("s4_fl", {}).get(region, {})
    total_hc = npat.get("5M2026", {}).get("HC")
    if fl and total_hc:
        frontliners = fl.get("sales_fl", 0) + fl.get("coll_fl", 0)
        pct_fl = frontliners / total_hc
        nas_fl_total = fl.get("nas_sales_fl", 0) + fl.get("nas_sales_non_fl", 0) + fl.get("nas_coll_fl", 0) + fl.get("nas_coll_non_fl", 0)
        nas_fl = fl.get("nas_sales_fl", 0) + fl.get("nas_coll_fl", 0)
        kalimat = (f"Proporsi Frontliners Staff adalah {pct_fl:.0%} dari total HC Managed "
                   f"({frontliners} dari {total_hc})")
        if nas_fl_total:
            pct_nas_fl = nas_fl / nas_fl_total
            komparasi = _bandingkan_pct(pct_fl, pct_nas_fl)
            kalimat += f". Proporsi ini {komparasi} proporsi frontliners Nasional ({pct_fl:.0%} vs {pct_nas_fl:.0%})"
        bagian_b.append(kalimat)
    if bagian_b:
        result["B"] = ". ".join(bagian_b) + "."

    # --- C: Span of Control + demografi keseluruhan ---
    soc = data.get("s4_soc", {}).get(region, {})
    nas_soc = data.get("s4_soc", {}).get("NAS", {})
    bagian_c = []
    if soc and nas_soc:
        def soc_komparasi(reg_key, label):
            reg_v, nat_v = soc.get(reg_key), nas_soc.get(reg_key)
            if reg_v is None or nat_v is None or not nat_v:
                return None
            komp = _bandingkan_pct(reg_v / nat_v - 1, 0, toleransi=0.05)
            return f"{label} {komp} Nasional ({reg_v} vs {nat_v:.2f})"

        sales_parts = [p for p in (soc_komparasi("sales_so_sh", "SO-SH"), soc_komparasi("sales_sh_mgr", "SH-MGR")) if p]
        if sales_parts:
            bagian_c.append(f"Di Field Sales, Span of Control {', '.join(sales_parts)}")
        coll_parts = [p for p in (soc_komparasi("coll_so_sh", "SO-SH"), soc_komparasi("coll_sh_mgr", "SH-MGR")) if p]
        if coll_parts:
            bagian_c.append(f"Di Field Coll, Span of Control {', '.join(coll_parts)}")

    lea = data.get("s4_lea", {})
    reg_lea = lea.get(region, {})
    if reg_lea:
        los_rows = dict(reg_lea.get("LOS", []))
        total_los = sum(los_rows.values())
        if total_los:
            los_muda = los_rows.get("a. <1 thn", 0) + los_rows.get("b. 1<x<5 thn", 0)
            pct_los_muda = los_muda / total_los
            los_kalimat = f"{pct_los_muda:.0%} karyawan memiliki masa kerja < 5 tahun"
            all_regions_los = {}
            for v in lea.values():
                for lbl, cnt in v.get("LOS", []):
                    all_regions_los[lbl] = all_regions_los.get(lbl, 0) + cnt
            total_nas_los = sum(all_regions_los.values())
            if total_nas_los:
                nas_los_muda = all_regions_los.get("a. <1 thn", 0) + all_regions_los.get("b. 1<x<5 thn", 0)
                pct_nas_los_muda = nas_los_muda / total_nas_los
                komp = _bandingkan_pct(pct_los_muda, pct_nas_los_muda)
                los_kalimat += f", proporsinya {komp} nasional ({pct_los_muda:.0%} vs {pct_nas_los_muda:.0%})"
            bagian_c.append(los_kalimat)

        edu_rows = dict(reg_lea.get("EDU", []))
        total_edu = sum(edu_rows.values())
        if total_edu:
            pct_sarjana = edu_rows.get("3. Sarjana", 0) / total_edu
            pct_diploma = edu_rows.get("2. Diploma", 0) / total_edu
            bagian_c.append(f"{pct_sarjana:.0%} karyawan adalah sarjana & {pct_diploma:.0%} karyawan adalah Diploma")

        age_rows = dict(reg_lea.get("AGE", []))
        total_age = sum(age_rows.values())
        if total_age:
            pct_2636 = age_rows.get("b. 26-36 thn", 0) / total_age
            bagian_c.append(f"{pct_2636:.0%} karyawan berusia di 26 - 36 tahun")

    if bagian_c:
        result["C"] = ". ".join(bagian_c) + "."

    return result


def get_slide5_highlights(data, region):
    """
    Membuat draf teks highlight/insight otomatis untuk slide 5, berdasarkan
    `data["s57"][region]` (*_sales_pct/*_coll_pct, hasil
    `load_slide5_7_data()`) dan `data["s5_nat"]` (angka mentah nasional,
    dikonversi ke persentase lewat `_pct_of_total`).

    STRUKTUR TEMPLATE (beda dari slide lain): token `{SLIDE5_HIGHLIGHT A}`
    (dengan SPASI, bukan underscore) muncul 4 KALI di kotak highlight yang
    sama (4 paragraf bullet terpisah yang semuanya masih berlabel "A" —
    tampaknya artefak saat template dibuat, paragraf di-duplikasi tapi
    lupa diganti jadi B/C/D), lalu ADA SATU LAGI token terpisah
    `{SLIDE5_HIGHLIGHT B}` di kotak lain. Karena itu, fungsi ini
    mengembalikan LIST 4 STRING (satu per bullet "A", diisi berurutan
    sesuai urutan tampil di XML oleh `_apply_slide5_highlights`) ditambah
    1 string "B" terpisah — cocok persis dengan 4 poin di
    Template_Highlight.txt (masa kerja Sales, masa kerja Collection, usia,
    pendidikan), tanpa perlu menambah/mengubah struktur token apa pun di
    template.

    4 poin (index list "A"):
      0 — Kategori masa kerja (LOS) DOMINAN untuk Field Sales (dicari
          lewat kategori dengan persentase tertinggi di `los_sales_pct`,
          BUKAN diasumsikan selalu "1 - 5 tahun" seperti pada contoh).
      1 — Sama untuk Field Coll, plus kategori kedua-terbesarnya (meniru
          pola contoh yang menyebut urutan populasi ke-2 untuk Collection).
      2 — Kategori usia DOMINAN untuk Sales Officer dan Collection Officer
          (dari `age_sales_pct`/`age_coll_pct`), dibandingkan ke persentase
          nasional yang sepadan.
      3 — Kategori pendidikan SLTA dibandingkan region vs nasional (untuk
          Sales dan Coll).
    "B" — SENGAJA DIKOSONGKAN: posisi/konteks kotak highlight "B" ini di
        template tidak jelas dari isi XML-nya saja (tidak ada label/judul
        di sekitarnya yang menunjukkan topiknya), jadi tidak dipaksakan
        diisi supaya tidak salah taruh insight yang tidak relevan dengan
        kotak itu.

    Parameter: data (dict load_all()), region (str).

    Return: dict {"A_list": [str, str, str, str], "B": ""}. Dipanggil dari
    `_apply_slide5_highlights()` di pptx_updater.py.
    """
    s57 = data.get("s57", {}).get(region, {})
    nat = data.get("s5_nat", {})
    result = {"A_list": ["", "", "", ""], "B": ""}
    if not s57:
        return result

    def top2(pct_rows):
        ranked = sorted(pct_rows, key=lambda r: -r[1])
        return ranked[0], (ranked[1] if len(ranked) > 1 else None)

    label = lambda lbl: lbl.split(". ", 1)[-1]

    los_sales = s57.get("los_sales_pct", [])
    los_coll = s57.get("los_coll_pct", [])
    if los_sales:
        top_sales, _ = top2(los_sales)
        result["A_list"][0] = (f"Secara masa kerja, Demografi Sales Officer di {region} didominasi "
                                f"oleh masa kerja {label(top_sales[0])} sebesar {top_sales[1]:.0%} "
                                f"dari total populasi SO.")
    if los_coll:
        top_coll, second_coll = top2(los_coll)
        top_sales_lbl = top2(los_sales)[0][0] if los_sales else None
        kalimat = "Sedangkan di Collection, "
        if top_sales_lbl == top_coll[0]:
            kalimat += f"meski sama-sama didominasi masa kerja {label(top_coll[0])} di {top_coll[1]:.0%}"
        else:
            kalimat += f"populasi terbanyak ada di masa kerja {label(top_coll[0])} sebesar {top_coll[1]:.0%}"
        if second_coll:
            kalimat += f", populasi kedua ada di masa kerja {label(second_coll[0])} dengan {second_coll[1]:.0%}"
        result["A_list"][1] = kalimat + "."

    age_sales = s57.get("age_sales_pct", [])
    age_coll = s57.get("age_coll_pct", [])
    nat_age = nat.get("AGE", [])
    if age_sales and age_coll:
        top_age_sales, _ = top2(age_sales)
        top_age_coll, _ = top2(age_coll)
        if top_age_sales[0] == top_age_coll[0]:
            label_usia = label(top_age_sales[0])
            kalimat = (f"Secara usia, Sales Officer dan Collection Officer dengan usia {label_usia} "
                       f"proporsinya paling dominan ({top_age_sales[1]:.0%} untuk Sales, "
                       f"{top_age_coll[1]:.0%} untuk Collection)")
            if nat_age:
                idx_age = [i for i, r in enumerate(age_sales) if r[0] == top_age_sales[0]][0]
                nat_sales_pct = _pct_of_total(nat_age, idx_age, 1)
                nat_coll_pct = _pct_of_total(nat_age, idx_age, 2)
                if nat_sales_pct is not None and nat_coll_pct is not None:
                    kalimat += (f". Secara nasional, proporsi untuk Sales Officer di {nat_sales_pct:.0%} "
                                f"dan Collection Officer di {nat_coll_pct:.0%}")
            result["A_list"][2] = kalimat + "."

    edu_sales = s57.get("edu_sales_pct", [])
    edu_coll = s57.get("edu_coll_pct", [])
    nat_edu = nat.get("EDU", [])
    if edu_sales and edu_coll and nat_edu:
        slta_sales, slta_coll = edu_sales[0][1], edu_coll[0][1]
        nat_slta_sales = _pct_of_total(nat_edu, 0, 1)
        nat_slta_coll = _pct_of_total(nat_edu, 0, 2)
        if nat_slta_sales is not None and nat_slta_coll is not None:
            sales_ok, coll_ok = slta_sales < nat_slta_sales, slta_coll < nat_slta_coll
            if sales_ok and coll_ok:
                result["A_list"][3] = ("Secara Pendidikan, baik proporsi SMA Sales maupun Collection "
                                        "berada di bawah proporsi nasional.")
            elif not sales_ok and not coll_ok:
                result["A_list"][3] = ("Secara Pendidikan, baik proporsi SMA Sales maupun Collection "
                                        "berada di atas proporsi nasional.")
            else:
                lebih = "Sales" if not sales_ok else "Collection"
                result["A_list"][3] = (f"Secara Pendidikan, proporsi SMA {lebih} berada di atas "
                                        f"proporsi nasional sedangkan sisi lainnya di bawah.")

    return result


# Pemetaan nama shape kotak persentase di slide 5 → (kategori ke berapa
# dari daftar terurut, apakah untuk Field Sales atau Field Coll). Urutan
# tiap pasangan SELALU (shape_sales, shape_coll) — hasil verifikasi manual
# terhadap posisi x tiap shape (shape dengan x lebih kecil = Field Sales,
# x lebih besar = Field Coll) dan dicocokkan persis dengan angka
# placeholder template (mis. box "Rectangle 64"=29% cocok dengan
# los_sales_pct kategori "a. <1 thn" Jawa Tengah=29%).
#
# Hanya kategori PERTAMA yang dapat kotak highlight di template: 3 kategori
# LOS (a, b, c), 2 kategori AGE (a, b), 1 kategori EDU (yang pertama) — sisa
# kategori lain tidak dianotasi dengan kotak %, jadi tidak perlu diisi.
_SLIDE5_PCT_BOX_MAP = {
    # (metrik, cakupan): [(shape_sales, shape_coll), ...] berurutan sesuai kategori
    ("LOS", "reg"): [("Rectangle 64", "Rectangle 61"), ("Rectangle 79", "Rectangle 65"), ("Rectangle 87", "Rectangle 86")],
    ("LOS", "nat"): [("Rectangle 81", "Rectangle 80"), ("Rectangle 83", "Rectangle 82"), ("Rectangle 85", "Rectangle 84")],
    ("AGE", "reg"): [("Rectangle 45", "Rectangle 44"), ("Rectangle 47", "Rectangle 46")],
    ("AGE", "nat"): [("Rectangle 30", "Rectangle 29"), ("Rectangle 32", "Rectangle 31")],
    ("EDU", "reg"): [("Rectangle 76", "Rectangle 75")],
    ("EDU", "nat"): [("Rectangle 78", "Rectangle 77")],
}


def get_slide5_pct_boxes(data, region):
    """
    Menghitung isi 24 kotak highlight persentase pada slide 5 (LOS/AGE/EDU
    × Regional/Nasional × Field Sales/Field Coll) — kotak-kotak ini adalah
    shape teks berdiri sendiri (bukan bagian dari chart), dan sebelum
    perbaikan ini TIDAK PERNAH diisi oleh kode (selalu menampilkan angka
    contoh statis dari template).

    Cara kerja: untuk data REGIONAL, ambil langsung dari
    `data["s57"][region]["<metrik>_sales_pct"/"_coll_pct"]` (sudah berupa
    persentase per kategori, lihat `load_slide5_7_data` di data_loader.py).
    Untuk data NASIONAL, `data["s5_nat"]["<metrik>"]` hanya berisi angka
    mentah per kategori, jadi persentasenya dihitung di sini lewat
    `_pct_of_total()` (jumlah kategori dibagi total seluruh kategori pada
    metrik yang sama).

    Hasil dipetakan ke nama shape lewat `_SLIDE5_PCT_BOX_MAP` (lihat
    penjelasan pemetaan di atasnya).

    Parameter: `data` (dict `load_all()`), `region` (str).

    Return: dict `{nama_shape: "XX%"}` — mis. `{"Rectangle 64": "29%",
    ...}`. Dipanggil dari `_apply_slide5_pct_boxes()` di pptx_updater.py,
    yang mencari tiap `<p:sp>` pada slide 5 berdasarkan nama shape ini dan
    mengganti isi teksnya.
    """
    s = data.get("s57", {}).get(region, {})
    nat = data.get("s5_nat", {})

    def fmt(v):
        return f"{v * 100:.0f}%" if v is not None else "0%"

    result = {}
    for metric, key_prefix in (("LOS", "los"), ("AGE", "age"), ("EDU", "edu")):
        # --- Regional: sudah dalam bentuk persentase per label ---
        sales_pct = dict(s.get(f"{key_prefix}_sales_pct", []))
        coll_pct  = dict(s.get(f"{key_prefix}_coll_pct", []))
        labels_reg = [lbl for lbl, _ in s.get(f"{key_prefix}_sales_pct", [])]
        for i, (box_sales, box_coll) in enumerate(_SLIDE5_PCT_BOX_MAP[(metric, "reg")]):
            if i < len(labels_reg):
                lbl = labels_reg[i]
                result[box_sales] = fmt(sales_pct.get(lbl))
                result[box_coll]  = fmt(coll_pct.get(lbl))

        # --- Nasional: masih angka mentah, hitung persentase dari total ---
        entries_nat = nat.get(metric, [])
        for i, (box_sales, box_coll) in enumerate(_SLIDE5_PCT_BOX_MAP[(metric, "nat")]):
            result[box_sales] = fmt(_pct_of_total(entries_nat, i, 1))
            result[box_coll]  = fmt(_pct_of_total(entries_nat, i, 2))

    return result


def get_slide4_total_hc(data, region):
    """
    Menghasilkan string total HC dengan format ribuan gaya Indonesia
    (pemisah '.'), mis. '2.537', diambil dari
    data["s4_npat"][region]["5M2026"]["HC"]. Dipanggil dari pptx_updater.py
    (dua tempat: saat mengisi teks slide 4, dan sekitar baris
    `hc = get_slide4_total_hc(data, region)`) untuk mengisi text box total
    HC.
    """
    npat = data["s4_npat"].get(region, {})
    hc = npat.get("5M2026", {}).get("HC", 0) or 0
    # Format angka gaya Indonesia: pemisah ribuan = '.'
    return f"{hc:,}".replace(",", ".")


def get_slide4_wc_pct(data, region):
    """
    Menghasilkan tuple (pct_permanent_str, pct_non_permanent_str) dalam
    format '%' tanpa desimal, dihitung dari data["s4_wc"][region]
    ["pct_permanent"] (non-permanent = 1.0 - pct_permanent). Dipanggil dari
    pptx_updater.py untuk mengisi text box persentase status kontrak kerja
    di slide 4.
    """
    wc = data["s4_wc"].get(region, {})
    perm = wc.get("pct_permanent", 0) or 0
    non_perm = 1.0 - perm
    return (f"{perm*100:.0f}%", f"{non_perm*100:.0f}%")
