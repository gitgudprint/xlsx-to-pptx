"""
Mendefinisikan array data untuk tiap nomor chart, berdasarkan data yang sudah
di-load per region (lihat data_loader.load_all()).
Mengembalikan (categories, series_list) atau (scatter_data) yang siap
diteruskan ke XML updater (src/xml_updater.py).
"""
from .config import (
    REGIONS, REGION_WILAYAH, REGION_ABBREV, FUNCTION_ORDER,
    LOS_ORDER, EDU_ORDER, AGE_ORDER,
)


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
    reg_vals = [func_d.get(f, 0) or 0 for f in cats]
    nas_vals = [nas_d.get(f, 0) or 0 for f in cats]
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
    reg_v = [fl.get("sales_fl", 0), fl.get("sales_non_fl", 0)]
    nas_v = [fl.get("nas_sales_fl", 0), fl.get("nas_sales_non_fl", 0)]
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
    reg_v = [fl.get("coll_fl", 0), fl.get("coll_non_fl", 0)]
    nas_v = [fl.get("nas_coll_fl", 0), fl.get("nas_coll_non_fl", 0)]
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
    # Slide 5
    18: chart18_edu_fl_reg,
    19: chart19_age_fl_reg,
    20: chart20_los_fl_reg,
    21: chart21_los_fl_nat,
    22: chart22_age_fl_nat,
    23: chart23_edu_fl_nat,
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
