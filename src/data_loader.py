"""
Memuat dan memproses semua file sumber xlsx menjadi struktur data per-region.

Modul ini berisi seluruh fungsi `load_*` yang membaca file-file xlsx (dikonfigurasi
lewat XLSX_FILES di src/config.py, key "a".."g") dan mengubahnya menjadi dict/list
Python yang mudah dipakai. Fungsi terakhir, `load_all()`, memanggil hampir semua
loader di file ini dan menggabungkan hasilnya menjadi satu dict besar (kunci seperti
"slide3", "s4_func", "attrition", "s15_func", "s16", "fraud", dst.) yang lalu dipakai
oleh src/chart_data.py (untuk mengisi chart) dan src/pptx_updater.py (untuk mengisi
tabel/teks di slide PPTX).
"""
import io
import re
import html
import zipfile
import posixpath
import functools
import openpyxl
import pandas as pd
from .config import (
    XLSX_FILES, REGIONS, REGION_WILAYAH, REGION_TRAINING, REGION_FRAUD,
    LOS_ORDER, LOS_ORDER_FRONTLINERS, EDU_ORDER, AGE_ORDER, FUNCTION_ORDER,
)


def _wb(key, read_only=True, data_only=True):
    """
    Buka workbook xlsx yang path-nya terdaftar di XLSX_FILES[key] (key adalah
    huruf "a".."g", lihat src/config.py) memakai openpyxl.

    Parameter:
      key: str, huruf file sumber ("a".."g").
      read_only: bool, mode baca-saja openpyxl (lebih hemat memori untuk file besar).
      data_only: bool, jika True openpyxl mengembalikan NILAI HASIL formula
        (bukan teks rumusnya) — penting karena banyak sheet sumber berisi formula.

    Return: objek Workbook openpyxl yang masih terbuka (pemanggil bertanggung
    jawab menutupnya, lihat `_sheet_rows` di bawah). Dipakai secara internal oleh
    `_sheet_rows`.
    """
    return openpyxl.load_workbook(XLSX_FILES[key], data_only=data_only, read_only=read_only)


def _sheet_rows(key, sheet_name):
    """
    Baca satu sheet penuh dari workbook `key` menjadi list of tuple (satu tuple
    per baris, berisi nilai sel apa adanya — tidak ada parsing/pembersihan).
    Membuka workbook lewat `_wb`, mengambil worksheet `sheet_name`, mengiterasi
    semua barisnya dengan `values_only=True`, lalu menutup workbook sebelum
    mengembalikan hasilnya.

    Parameter:
      key: str, huruf file sumber ("a".."g").
      sheet_name: str, nama sheet persis seperti di file Excel.

    Return: list[tuple], satu tuple per baris (indeks 0 = baris pertama sheet).
    Dipakai oleh hampir semua loader baris-tetap di file ini (mis. load_slide3_yoy,
    load_slide4_*, load_slide10_training, load_slide12_scatter, load_attrition)
    yang lalu mengakses kolom lewat indeks tetap `r[i]`.
    """
    wb = _wb(key)
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows


def _pd(key, sheet_name, header=0):
    """
    Muat satu sheet ke pandas DataFrame lewat pd.read_excel (dipakai saat data
    lebih nyaman diproses dengan operasi DataFrame — filter/groupby/value_counts —
    dibanding list of tuple mentah dari `_sheet_rows`).

    Parameter:
      key: str, huruf file sumber ("a".."g").
      sheet_name: str, nama sheet.
      header: int, indeks baris header (diteruskan langsung ke pandas).

    Return: pandas.DataFrame. Catatan: fungsi ini sendiri tidak dipanggil dari
    tempat lain di file ini saat ini (loader DataFrame yang dipakai adalah
    `_load_db_b` dan `_load_active_frontliners` di bagian bawah, yang memanggil
    pd.read_excel langsung, bukan lewat `_pd`).
    """
    return pd.read_excel(XLSX_FILES[key], sheet_name=sheet_name, header=header, engine="openpyxl")


# ---------------------------------------------------------------------------
# SLIDE 3 – Chart perbandingan YoY (chart 1-6)
# Semua region ditampilkan di setiap chart, diurutkan naik berdasarkan nilai 2026.
# ---------------------------------------------------------------------------
def load_slide3_yoy():
    """
    Baca sheet "Slide 3 - YoY" (asumsi layout kolom TETAP, lihat daftar indeks
    di bawah) dan susun ulang menjadi dict per metrik untuk chart YoY.

    Layout kolom di sheet 'Slide 3 - YoY' (0-indexed):
      0=Org, 1=Region(HC), 2=HC26, 3=HC25
      7=Region(BISNIS), 8=BISNIS26, 9=BISNIS25
      13=Region(SSD), 14=SSD26, 15=SSD25
      19=Region(COLL), 20=COLL26, 21=COLL25
      25=Region(CREDIT), 26=CREDIT26, 27=CREDIT25
      31=Region(LAR), 32=LAR26, 33=LAR25
      Juga: 5=diff% HC, 11=diff% BISNIS, 17=diff% SSD, 23=diff% COLL, 29=diff% CREDIT, 35=diff% LAR

    Parameter: tidak ada (path file diambil dari XLSX_FILES["a"] via _sheet_rows).

    Return: dict dengan key = nama metrik ("HC", "BISNIS", "SSD", "COLL",
    "CREDIT", "LAR") → list of (region, val_2026, val_2025) diurutkan naik
    berdasarkan val_2026, ditambah key khusus "_dif_pct" → dict region →
    {metrik: nilai diff%}. Dipanggil oleh load_all() dan disimpan sebagai
    data["slide3"]; dipakai oleh chart_data.py (chart1_hc..chart6_bisnis, via
    data["slide3"][<metrik>]) dan oleh get_slide3_annotations()/pptx_updater.py
    (via data["slide3"]["_dif_pct"][region]) untuk anotasi teks diff% di slide.
    """
    rows = _sheet_rows("a", "Slide 3 - YoY")

    # Lewati baris header (baris 0-1 adalah tahun/header) dan baris Grand Total
    data_rows = [r for r in rows[2:] if r and r[0] and r[0] != "Grand Total"]

    metrics = {
        "HC":     (1, 2, 3, 5),
        "BISNIS": (7, 8, 9, 11),
        "SSD":    (13, 14, 15, 17),
        "COLL":   (19, 20, 21, 23),
        "CREDIT": (25, 26, 27, 29),
        "LAR":    (31, 32, 33, 35),
    }

    result = {}
    for metric, (ri, v26i, v25i, pct_i) in metrics.items():
        entries = []
        for r in data_rows:
            region = r[ri]
            v26 = r[v26i]
            v25 = r[v25i]
            if region and v26 is not None:
                entries.append((region, v26, v25))
        entries.sort(key=lambda x: x[1])
        result[metric] = entries

    # Ambil juga dif% per region untuk anotasi teks
    dif_pct = {}
    for r in data_rows:
        region = r[1]
        if region:
            dif_pct[region] = {
                "HC":     r[5],
                "BISNIS": r[11],
                "SSD":    r[17],
                "COLL":   r[23],
                "CREDIT": r[29],
                "LAR":    r[35],
            }
    result["_dif_pct"] = dif_pct
    return result


# ---------------------------------------------------------------------------
# SLIDE 4 – Data demografi (chart 7-17)
# ---------------------------------------------------------------------------
def load_slide4_function():
    """
    Baca sheet "Slide 4 - FUNCTION" (asumsi layout baris/kolom tetap: baris 1
    kosong, baris 2 = header, data mulai baris 3) dan susun headcount per
    fungsi untuk tiap region.

    Parameter: tidak ada.

    Return: dict region → {BISNIS, SSD, COLL, CREDIT, LAR, SUPP OPR, Honorer,
    OB Sec} (nama kolom fungsi diambil langsung dari header baris 2), ditambah
    entry "AVG NASIONAL". Dipanggil oleh load_all() dan disimpan sebagai
    data["s4_func"]; dipakai oleh chart7_function() di chart_data.py
    (dibandingkan dengan data["s4_func"]["AVG NASIONAL"]).
    """
    rows = _sheet_rows("a", "Slide 4 - FUNCTION")
    # Baris 1 kosong, baris 2 adalah header (Org, Row Labels, BISNIS, SSD, COLL, CREDIT, LAR, SUPP OPR, Honorer, OB Sec, Grand Total)
    # Baris data dimulai dari baris 3 (index 2)
    header = rows[1]  # index 1 = baris 2
    func_cols = list(header[2:11])  # BISNIS...OB Sec

    result = {}
    for r in rows[2:]:
        label = r[1]
        if label and label != "Grand Total":
            # Buang prefix "Wilayah Area " agar didapat nama region kanonis
            region = label.replace("Wilayah Area ", "")
            result[region] = dict(zip(func_cols, r[2:11]))

    # Baris Grand Total / AVG NASIONAL
    avg_row = rows[-2]  # baris kedua dari bawah adalah AVG NASIONAL
    if avg_row[1] == "AVG NASIONAL":
        result["AVG NASIONAL"] = dict(zip(func_cols, avg_row[2:11]))
    else:
        # Fallback: pakai Grand Total dibagi jumlah region sebagai indikator rata-rata
        gt_row = next((r for r in rows if r[1] == "Grand Total"), None)
        if gt_row:
            n = len([r for r in rows[2:] if r[1] and r[1] != "Grand Total" and r[1] != "AVG NASIONAL"])
            result["AVG NASIONAL"] = {k: (gt_row[2 + i] / n if gt_row[2 + i] else 0)
                                      for i, k in enumerate(func_cols)}
    return result


def load_slide4_npat_hc():
    """
    Baca sheet "Slide 4 - NPAT - HC" (asumsi layout kolom tetap, 2 baris
    header bertingkat lalu data mulai baris 3) dan susun HC/NPAT/NPAT-per-HC
    per region untuk 4 periode.

    Layout header (untuk referensi):
      Baris header 1: (None, '5M2026', 'FY25', 'FY25', 'FY25', 'FY24', 'FY24', 'FY24', 'FY23', 'FY23', 'FY23')
      Baris header 2: ('Row Labels', 'HC Managed', 'HC Managed', 'NPAT', 'NPAT/HC', 'HC Managed', 'NPAT', 'NPAT/HC', ...)

    Parameter: tidak ada.

    Return: dict region → {
        '5M2026': {'HC': int},
        'FY25': {'HC': int, 'NPAT': float, 'NPAT_HC': float},
        'FY24': {'HC': int, 'NPAT': float, 'NPAT_HC': float},
        'FY23': {'HC': int, 'NPAT': float, 'NPAT_HC': float},
    }
    Dipanggil oleh load_all() dan disimpan sebagai data["s4_npat"]; dipakai oleh
    chart8_hc_npat() dan chart15_npat_hc_line() di chart_data.py, serta oleh
    get_slide4_total_hc() (via data["s4_npat"][region]["5M2026"]["HC"]) yang
    dipanggil dari pptx_updater.py.
    """
    rows = _sheet_rows("a", "Slide 4 - NPAT - HC")
    result = {}
    for r in rows[2:]:
        if not r or not r[0] or r[0] == "Grand Total":
            continue
        region = r[0]
        result[region] = {
            "5M2026": {"HC": r[1]},
            "FY25":   {"HC": r[2], "NPAT": r[3], "NPAT_HC": r[4]},
            "FY24":   {"HC": r[5], "NPAT": r[6], "NPAT_HC": r[7]},
            "FY23":   {"HC": r[8], "NPAT": r[9], "NPAT_HC": r[10]},
        }
    return result


def load_slide4_work_contract():
    """
    Baca sheet "Slide 4 - Work Contract" (asumsi layout baris/kolom tetap:
    baris nasional di baris 3, data per region mulai baris 8, persentase di
    kolom O/P/Q) dan susun persentase status kerja (permanen/kontrak/OS) per
    region beserta angka nasionalnya.

    Layout:
      Baris 3 (index 2): total nasional ADMF
      Baris 8+ (index 7+): data per region
      Kolom persentase (0-indexed): O=14 (perm%), P=15 (contract%), Q=16 (OS%)

    Parameter: tidak ada.

    Return: dict region → {pct_permanent, pct_contract, pct_os, nas_permanent,
    nas_contract, nas_os} (3 nilai nas_* sama untuk semua region — angka
    nasional tunggal dari baris 3). Dipanggil oleh load_all() dan disimpan
    sebagai data["s4_wc"]; dipakai oleh chart9_work_contract() di
    chart_data.py dan oleh get_slide4_wc_pct() (dipanggil dari
    pptx_updater.py) via data["s4_wc"][region].
    """
    rows = _sheet_rows("a", "Slide 4 - Work Contract")

    # Baris nasional adalah baris 3 (index 2): Org='ADMF', label='ADMF'
    nat_row = rows[2]
    nas_perm = nat_row[14]
    nas_cont = nat_row[15]
    nas_os = nat_row[16]

    result = {}
    # Baris per region dimulai dari index 7 (baris 8)
    for r in rows[7:]:
        if not r[1] or r[1] in ("Grand Total", "AVG NASIONAL", "Row Labels"):
            continue
        if r[0] not in ("ADMF", "Grand Total") and r[1] == "Grand Total":
            continue
        label = r[1]
        if label and str(label).startswith("Wilayah Area"):
            region = label.replace("Wilayah Area ", "")
            result[region] = {
                "pct_permanent": r[14],
                "pct_contract":  r[15],
                "pct_os":        r[16],
                "nas_permanent": nas_perm,
                "nas_contract":  nas_cont,
                "nas_os":        nas_os,
            }
    return result


def load_slide4_los_edu_age(df_b=None):
    """
    Hitung distribusi Length-of-Service (LOS), pendidikan (EDU), dan usia
    (AGE) per region dari database karyawan file b, dengan mengecualikan
    fungsi HEAD OFFICE. Bukan pembacaan sheet baris-tetap seperti loader
    lain — ini agregasi pandas (value_counts) atas seluruh baris DataFrame,
    difilter per region lewat kolom "Region Function".

    Parameter:
      df_b: pandas.DataFrame atau None. Jika None, dimuat sendiri lewat
        `_load_db_b()`. load_all() memanggil fungsi ini dengan df_b yang
        SUDAH dimuat sekali (dibagi juga ke load_slide4_span_of_control)
        agar file b tidak dibaca dua kali.

    Return: dict region → {
        'LOS': [(label, count), ...],   # urutan label ikut LOS_ORDER
        'EDU': [(label, count), ...],   # urutan label ikut edu_map di bawah
        'AGE': [(label, count), ...],   # urutan label ikut AGE_ORDER
    }
    Dipanggil oleh load_all() dan disimpan sebagai data["s4_lea"]; dipakai
    oleh chart10_los(), chart11_edu(), chart12_age() di chart_data.py via
    data["s4_lea"][region]["LOS"/"EDU"/"AGE"].
    """
    if df_b is None:
        df_b = _load_db_b()

    df = df_b[df_b["FUNCTION NEW 1"] != "HEAD OFFICE"].copy()

    result = {}
    for region in REGIONS:
        wilayah = REGION_WILAYAH[region]
        rdf = df[df["Region Function"] == wilayah]

        # LOS
        los_counts = rdf["Los cat"].value_counts()
        los_data = []
        for label in LOS_ORDER:
            los_data.append((label, int(los_counts.get(label, 0))))

        # EDU (percobaan pertama: cocokkan label parsial — hasilnya dibuang
        # dan dihitung ulang di bawah dengan pemetaan prefix eksplisit)
        edu_counts = rdf["Edu Cat"].value_counts()
        edu_data = []
        for label in EDU_ORDER:
            # Cocokkan label parsial
            matched = sum(v for k, v in edu_counts.items() if str(k).startswith(label[:4]))
            edu_data.append((label, int(matched) if matched else 0))
        # Hitung ulang dengan pemetaan langsung
        edu_data = []
        edu_map = {
            "1. SLTA sederajat & di bawahnya": "1.",
            "2. Diploma":                       "2.",
            "3. Sarjana":                       "3.",
            "4. Pasca Sarjana":                 "4.",
        }
        for full_label, prefix in edu_map.items():
            cnt = sum(v for k, v in edu_counts.items() if str(k).startswith(prefix))
            edu_data.append((full_label, int(cnt)))

        # AGE (usia)
        age_counts = rdf["Age Cat"].value_counts()
        age_data = []
        for label in AGE_ORDER:
            age_data.append((label, int(age_counts.get(label, 0))))

        result[region] = {"LOS": los_data, "EDU": edu_data, "AGE": age_data}

    return result


def load_slide4_span_of_control(df_b=None):
    """
    Hitung rasio span-of-control (jumlah bawahan per atasan) untuk hierarki
    Sales (Sales Officer -> Sales Head -> HOS) dan Collection (Rem Officer ->
    ARH -> CCH/CCS) per region, dari database karyawan file b; nilai
    nasionalnya sendiri dibaca langsung dari sheet precomputed (bukan
    dihitung ulang dari df_b).

    Parameter:
      df_b: pandas.DataFrame atau None. Jika None, dimuat sendiri lewat
        `_load_db_b()`. load_all() memanggil fungsi ini dengan df_b yang
        sama yang dipakai load_slide4_los_edu_age (dimuat sekali saja).

    Return: dict region → {sales_so_sh, sales_sh_mgr, coll_so_sh,
    coll_sh_mgr} (rasio, dibulatkan 2 desimal), ditambah entry "NAS" berisi
    4 nilai nasional yang sama untuk semua region. Dipanggil oleh load_all()
    dan disimpan sebagai data["s4_soc"]; dipakai oleh chart13_soc_sales() dan
    chart14_soc_coll() di chart_data.py via data["s4_soc"][region] dan
    data["s4_soc"]["NAS"].
    """
    if df_b is None:
        df_b = _load_db_b()

    # Nilai nasional dari sheet yang sudah dihitung sebelumnya (precomputed)
    rows = _sheet_rows("b", "SLIDE 4 - SPAN OF CONTROL")
    nas_sales_so_sh  = rows[2][2]  # baris3 kol C
    nas_sales_sh_mgr = rows[3][2]  # baris4 kol C
    nas_coll_so_sh   = rows[2][6]  # baris3 kol G
    nas_coll_sh_mgr  = rows[3][6]  # baris4 kol G

    result = {"NAS": {
        "sales_so_sh":  nas_sales_so_sh,
        "sales_sh_mgr": nas_sales_sh_mgr,
        "coll_so_sh":   nas_coll_so_sh,
        "coll_sh_mgr":  nas_coll_sh_mgr,
    }}

    for region in REGIONS:
        wilayah = REGION_WILAYAH[region]
        rdf = db = df_b[df_b["Region Function"] == wilayah]

        # Hierarki Sales dari kolom JOB CAT
        sales_df = rdf[rdf["FUNCTION NEW 1"] == "SALES"]
        so_cnt  = (sales_df["JOB CAT"] == "Sales Officer").sum()
        sh_cnt  = (sales_df["JOB CAT"] == "Sales Head + CDO").sum()
        hos_cnt = (sales_df["JOB CAT"] == "HOS").sum()

        so_sh  = (so_cnt  / sh_cnt)  if sh_cnt  > 0 else 0
        sh_mgr = (sh_cnt  / hos_cnt) if hos_cnt > 0 else 0

        # Hierarki Collection
        coll_df = rdf[rdf["FUNCTION NEW 1"] == "COLL"]
        rem_cnt = (coll_df["JOB CAT"] == "Rem Off").sum()
        arh_cnt = (coll_df["JOB CAT"] == "ARH").sum()
        ccx_cnt = (coll_df["JOB CAT"].isin(["CCH", "CCS"])).sum()

        coll_so_sh  = (rem_cnt / arh_cnt) if arh_cnt  > 0 else 0
        coll_sh_mgr = (arh_cnt / ccx_cnt) if ccx_cnt  > 0 else 0

        result[region] = {
            "sales_so_sh":  round(so_sh,  2),
            "sales_sh_mgr": round(sh_mgr, 2),
            "coll_so_sh":   round(coll_so_sh,  2),
            "coll_sh_mgr":  round(coll_sh_mgr, 2),
        }

    return result


def load_slide4_frontliners():
    """
    Baca sheet "SLIDE 4 - Proporsi Frontliners" (tabel ringkasan semua region,
    asumsi layout baris/kolom tetap seperti tercatat di komentar-komentar
    eksplorasi di bawah) dan susun jumlah frontliner vs non-frontliner untuk
    Sales dan Collection per region, plus rata-rata nasionalnya.

    Parameter: tidak ada.

    Return: dict region → {
        'sales_fl': int, 'sales_non_fl': int,
        'coll_fl':  int, 'coll_non_fl':  int,
        'nas_sales_fl': float, 'nas_sales_non_fl': float,
        'nas_coll_fl':  float, 'nas_coll_non_fl':  float,
    }
    Dipanggil oleh load_all() dan disimpan sebagai data["s4_fl"]; dipakai
    oleh chart16_fl_sales() dan chart17_fl_coll() di chart_data.py via
    data["s4_fl"][region].
    """
    rows = _sheet_rows("b", "SLIDE 4 - Proporsi Frontliners")

    # Bagian SALES: baris 12-15 (index 11-14)
    # Baris 12 (index 11): header berisi nama region di kolom D-P (index 3-15)
    sales_header_row = rows[11]
    sales_fl_row     = rows[12]
    sales_non_fl_row = rows[13]
    sales_total_row  = rows[14]

    # Indeks kolom 3 = 'Wilayah Area Bali', 4='Jabar', 5='Jabotabek 1', 6='Jabotabek 2',
    # 7='Jawa Tengah', 8='Jawa Timur', 9='Kalimantan', 10='Pasima',
    # 11='Sultan', 12='Sumatera 1', 13='Sumatera 2', 14='Sumatera 3'
    sales_region_cols = {}
    for i, v in enumerate(sales_header_row):
        if v and str(v).startswith("Wilayah Area"):
            r = str(v).replace("Wilayah Area ", "")
            sales_region_cols[r] = i

    # Rata-rata NAS ada di kolom B (index 1) pada baris 13 (FL) dan 14 (non-FL)
    nas_sales_fl     = rows[3][2]   # Baris 4 kol C = NAS fl (rata-rata float)
    nas_sales_non_fl = rows[3][1]   # Baris 4 kol B = ... perlu dicek ulang

    # Dari data sebenarnya:
    # Baris 4 (index 3): ('REGION', 237, 373, ...) → kol1=non_fl, kol2=fl untuk region saat ini
    # Baris 4 kol 1=region non_fl (237 untuk Sumatera 1), kol 2=fl (373)
    # NAS: kol 2 baris 4 di bagian kedua? Sebenarnya dari sheet:
    # rows[12] = ('Frontliners', 373, 526.58, 'Frontliners', 348, 653, 1271, 274, 617, 562, 668, 343, 346, 373, 436, 428, 6319)
    # Jadi indeks kol 1=REG fl (373 untuk Sumatera 1), kol 2=rata-rata NAS fl (526.58)
    # rows[13] = ('Non Frontliners', 237, 318.83, ..., 192, 468, ...)
    # Jadi REG non_fl=kol1=237, rata-rata NAS non_fl=kol2=318.83
    nas_sales_fl     = sales_fl_row[2]     if sales_fl_row[2] else 0
    nas_sales_non_fl = sales_non_fl_row[2] if sales_non_fl_row[2] else 0

    # Bagian COLL: mirip, tapi mulai dari rows[20] (index 20)
    coll_header_row  = rows[20]
    coll_fl_row      = rows[21]
    coll_non_fl_row  = rows[22]

    coll_region_cols = {}
    for i, v in enumerate(coll_header_row):
        if v and str(v).startswith("Wilayah Area"):
            r = str(v).replace("Wilayah Area ", "")
            coll_region_cols[r] = i

    nas_coll_fl     = coll_fl_row[2]     if len(coll_fl_row) > 2 and coll_fl_row[2] else 0
    nas_coll_non_fl = coll_non_fl_row[2] if len(coll_non_fl_row) > 2 and coll_non_fl_row[2] else 0

    result = {}
    for region in REGIONS:
        sc = sales_region_cols.get(region)
        cc = coll_region_cols.get(region)
        result[region] = {
            "sales_fl":      int(sales_fl_row[sc])     if sc and sales_fl_row[sc]     else 0,
            "sales_non_fl":  int(sales_non_fl_row[sc]) if sc and sales_non_fl_row[sc] else 0,
            "coll_fl":       int(coll_fl_row[cc])      if cc and coll_fl_row[cc]      else 0,
            "coll_non_fl":   int(coll_non_fl_row[cc])  if cc and coll_non_fl_row[cc]  else 0,
            "nas_sales_fl":      nas_sales_fl,
            "nas_sales_non_fl":  nas_sales_non_fl,
            "nas_coll_fl":       nas_coll_fl,
            "nas_coll_non_fl":   nas_coll_non_fl,
        }

    return result


# ---------------------------------------------------------------------------
# SLIDE 5-7 – Potret staf frontliner & kinerja PA
# ---------------------------------------------------------------------------
def load_slide5_7_data(df_c=None):
    """
    Hitung, dari database "Active Frontliners" file c, seluruh breakdown
    LOS/usia/pendidikan (count dan persentase) serta breakdown PA (Penilaian
    kinerja: kategori 2/3/>=4) untuk staf Field Sales dan Field Coll, per
    region. Ini agregasi pandas (value_counts/groupby manual) atas DataFrame
    yang sudah difilter per region via kolom "Region Function", bukan
    pembacaan sheet baris-tetap.

    Parameter:
      df_c: pandas.DataFrame atau None. Jika None, dimuat sendiri lewat
        `_load_active_frontliners()`. load_all() memanggil fungsi ini dengan
        df_c yang sudah dimuat sekali di awal.

    Return: dict region → {
        'los_sales': [(label, count), ...],
        'los_coll':  [(label, count), ...],
        'los_sales_pct': [(label, pct), ...],
        'los_coll_pct':  [(label, pct), ...],
        'age_sales': [...], 'age_coll': [...],
        'age_sales_pct': [...], 'age_coll_pct': [...],
        'edu_sales': [...], 'edu_coll': [...],
        'edu_sales_pct': [...], 'edu_coll_pct': [...],
        'pa_sales_los': [(label, pa2, pa3, pa4), ...],
        'pa_sales_age': [...],
        'pa_sales_edu': [...],
        'pa_coll_los':  [...],
        'pa_coll_age':  [...],
        'pa_coll_edu':  [...],
        'count_sales_los': [(label, count), ...],
        'count_sales_age': [...],
        'count_sales_edu': [...],
        'count_coll_los':  [...],
        'count_coll_age':  [...],
        'count_coll_edu':  [...],
    }
    Dipanggil oleh load_all() dan disimpan sebagai data["s57"]; dipakai oleh
    banyak fungsi chart di chart_data.py (chart18-chart20, chart24-chart29,
    chart33-chart38) via data["s57"][region][<key>].
    """
    if df_c is None:
        df_c = _load_active_frontliners()

    result = {}
    for region in REGIONS:
        wilayah = REGION_WILAYAH[region]
        rdf = df_c[df_c["Region Function"] == wilayah]

        sales_df = rdf[rdf["Frontliners Category"] == "Field Sales"]
        coll_df  = rdf[rdf["Frontliners Category"] == "Field Coll"]

        def count_by(df, col, order):
            counts = df[col].value_counts()
            return [(lbl, int(counts.get(lbl, 0))) for lbl in order]

        def pct_by(df, col, order):
            total = len(df)
            if total == 0:
                return [(lbl, 0.0) for lbl in order]
            counts = df[col].value_counts()
            return [(lbl, counts.get(lbl, 0) / total) for lbl in order]

        def pa_breakdown(df, group_col, order):
            """Untuk tiap kategori dalam `order`, hitung % PA (Penilaian) 2, 3, >=4."""
            rows = []
            for lbl in order:
                sub = df[df[group_col] == lbl]
                total = len(sub)
                if total == 0:
                    rows.append((lbl, 0.0, 0.0, 0.0))
                else:
                    pa2 = (sub["cat pa"] == 2).sum() / total
                    pa3 = (sub["cat pa"] == 3).sum() / total
                    pa4 = (sub["cat pa"].isin([">=4", 4])).sum() / total
                    rows.append((lbl, pa2, pa3, pa4))
            return rows

        # Pemetaan label pendidikan antara database dan format chart
        def _edu_count(df, display_order):
            counts = df["Edu Cat"].value_counts()
            result_list = []
            for lbl in display_order:
                cnt = sum(v for k, v in counts.items() if str(k).startswith(lbl[:2]))
                result_list.append((lbl, int(cnt)))
            return result_list

        def _edu_pct(df, display_order):
            total = len(df)
            if total == 0:
                return [(lbl, 0.0) for lbl in display_order]
            counts = df["Edu Cat"].value_counts()
            result_list = []
            for lbl in display_order:
                cnt = sum(v for k, v in counts.items() if str(k).startswith(lbl[:2]))
                result_list.append((lbl, cnt / total))
            return result_list

        def _edu_pa(df, display_order):
            result_list = []
            for lbl in display_order:
                sub = df[[str(r).startswith(lbl[:2]) for r in df["Edu Cat"]]]
                total = len(sub)
                if total == 0:
                    result_list.append((lbl, 0.0, 0.0, 0.0))
                else:
                    pa2 = (sub["cat pa"] == 2).sum() / total
                    pa3 = (sub["cat pa"] == 3).sum() / total
                    pa4 = (sub["cat pa"].isin([">=4", 4])).sum() / total
                    result_list.append((lbl, pa2, pa3, pa4))
            return result_list

        edu_chart_order = ["1. SLTA sederajat & di bawahnya", "2. Diploma", "3. Sarjana", "4. Pasca Sarjana"]

        result[region] = {
            # Slide 5 – Jumlah dan proporsi frontliner
            # CATATAN: LOS di slide 5/6/7 memakai LOS_ORDER_FRONTLINERS
            # (menaik a → f), BUKAN LOS_ORDER (menurun f → a, khusus
            # chart10 slide 4). Template chart 20/24/25/33/34 dan seluruh
            # data nasionalnya memakai urutan menaik, jadi bagian regional
            # harus sama agar tidak tampil terbalik.
            "los_sales":     count_by(sales_df, "Los cat", LOS_ORDER_FRONTLINERS),
            "los_coll":      count_by(coll_df,  "Los cat", LOS_ORDER_FRONTLINERS),
            "los_sales_pct": pct_by(sales_df, "Los cat", LOS_ORDER_FRONTLINERS),
            "los_coll_pct":  pct_by(coll_df,  "Los cat", LOS_ORDER_FRONTLINERS),
            "age_sales":     count_by(sales_df, "Age Cat", AGE_ORDER),
            "age_coll":      count_by(coll_df,  "Age Cat", AGE_ORDER),
            "age_sales_pct": pct_by(sales_df, "Age Cat", AGE_ORDER),
            "age_coll_pct":  pct_by(coll_df,  "Age Cat", AGE_ORDER),
            "edu_sales":     _edu_count(sales_df, edu_chart_order),
            "edu_coll":      _edu_count(coll_df,  edu_chart_order),
            "edu_sales_pct": _edu_pct(sales_df, edu_chart_order),
            "edu_coll_pct":  _edu_pct(coll_df,  edu_chart_order),
            # Slide 6 – PA Sales (LOS juga memakai urutan menaik, lihat catatan di atas)
            "pa_sales_los":  pa_breakdown(sales_df, "Los cat", LOS_ORDER_FRONTLINERS),
            "pa_sales_age":  pa_breakdown(sales_df, "Age Cat", AGE_ORDER),
            "pa_sales_edu":  _edu_pa(sales_df, edu_chart_order),
            "count_sales_los": count_by(sales_df, "Los cat", LOS_ORDER_FRONTLINERS),
            "count_sales_age": count_by(sales_df, "Age Cat", AGE_ORDER),
            "count_sales_edu": _edu_count(sales_df, edu_chart_order),
            # Slide 7 – PA Coll (idem, urutan LOS menaik)
            "pa_coll_los":  pa_breakdown(coll_df, "Los cat", LOS_ORDER_FRONTLINERS),
            "pa_coll_age":  pa_breakdown(coll_df, "Age Cat", AGE_ORDER),
            "pa_coll_edu":  _edu_pa(coll_df, edu_chart_order),
            "count_coll_los": count_by(coll_df, "Los cat", LOS_ORDER_FRONTLINERS),
            "count_coll_age": count_by(coll_df, "Age Cat", AGE_ORDER),
            "count_coll_edu": _edu_count(coll_df, edu_chart_order),
        }

    return result


def load_slide5_national():
    """
    Baca sheet precomputed "SLIDE 5 - Nat Fl (chart bawah)" (asumsi layout
    baris tetap seperti tercatat di komentar bagian bawah fungsi) untuk
    mengambil data LOS/usia/pendidikan NASIONAL (sama untuk semua region,
    tidak difilter per region — file c sudah punya sheet ringkasan nasional
    tersendiri) yang dipakai chart 21-23 pada slide 5.

    Parameter: tidak ada.

    Return: dict {"LOS": [(label, val_sales, val_coll), ...], "AGE": [...],
    "EDU": [...]}. Dipanggil oleh load_all() dan disimpan sebagai
    data["s5_nat"]; dipakai oleh chart21_los_fl_nat(), chart22_age_fl_nat(),
    chart23_edu_fl_nat() di chart_data.py (parameter `region` pada fungsi
    chart tersebut tidak dipakai, karena data ini sama untuk semua region).

    Catatan: `los_data`/`age_data` di bawah dihitung tapi TIDAK dipakai —
    hasil akhir fungsi ini dibangun ulang lewat `result_los`/`result_age`/
    `result_edu` di bagian bawah.
    """
    rows = _sheet_rows("c", "SLIDE 5 - Nat Fl (chart bawah)")
    los, age, edu = _split_national_sections(rows, value_cols=(1, 2))
    return {"LOS": los, "AGE": age, "EDU": edu}


def load_slide6_national():
    """
    Data PA Sales NASIONAL (tetap/sama untuk semua region) dari sheet
    precomputed "SLIDE 6 - PA Sales Nas", lewat `_load_pa_national`.
    Dipanggil oleh load_all() dan disimpan sebagai data["s6_nat"]; dipakai
    oleh chart30_pa_sales_edu_nat(), chart31_pa_sales_age_nat(),
    chart32_pa_sales_los_nat() di chart_data.py.
    """
    return _load_pa_national("c", "SLIDE 6 - PA Sales Nas")


def load_slide7_national():
    """
    Data PA Coll NASIONAL (tetap/sama untuk semua region) dari sheet
    precomputed "SLIDE 7 - PA Coll Nas", lewat `_load_pa_national`.
    Dipanggil oleh load_all() dan disimpan sebagai data["s7_nat"]; dipakai
    oleh chart39_pa_coll_edu_nat(), chart40_pa_coll_age_nat(),
    chart41_pa_coll_los_nat() di chart_data.py.
    """
    return _load_pa_national("c", "SLIDE 7 - PA Coll Nas")


def _split_national_sections(rows, value_cols):
    """
    Memecah sheet nasional (slide 5/6/7) menjadi 3 bagian berurutan: LOS,
    AGE, lalu EDU — TANPA mengandalkan nomor baris yang di-hardcode.

    Cara kerja: sheet-sheet ini berisi 3 blok pivot yang masing-masing
    diawali baris header berlabel "Row Labels" di kolom A dan diakhiri
    baris "Grand Total". Fungsi ini mencari ketiga baris "Row Labels"
    tersebut, lalu mengambil baris-baris data di bawahnya sampai bertemu
    "Grand Total" (atau baris kosong). Urutan blok di semua sheet ini
    selalu LOS → AGE → EDU.

    Alasan tidak memakai rentang baris tetap: versi sebelumnya memakai
    rentang hardcode (mis. `rows[29:34]` untuk EDU) yang meleset satu blok
    ketika layout sheet bergeser — akibatnya kategori "4. Pasca Sarjana"
    terpotong di chart EDU nasional slide 5, dan chart EDU nasional slide
    6/7 hanya menyisakan 1 kategori dari 4.

    Parameter:
        rows: list tuple hasil `_sheet_rows(...)`.
        value_cols: tuple index kolom nilai yang mau diambil per baris,
            mis. (1, 2) untuk sheet slide 5 (Field Sales, Field Coll) atau
            (1, 2, 3) untuk sheet PA (pa2, pa3, pa>=4).

    Return: tuple (los, age, edu); tiap elemen list of tuple
    `(label, *nilai)` sesuai `value_cols`. Dipakai oleh
    `load_slide5_national()` dan `_load_pa_national()`.
    """
    header_idxs = [i for i, r in enumerate(rows)
                   if r and r[0] and str(r[0]).strip() == "Row Labels"]

    sections = []
    for start in header_idxs[:3]:
        block = []
        for r in rows[start + 1:]:
            if not r or not r[0]:
                break
            label = str(r[0]).strip()
            if label == "Grand Total":
                break
            block.append(tuple([label] + [(r[c] if c < len(r) and r[c] is not None else 0)
                                          for c in value_cols]))
        sections.append(block)

    while len(sections) < 3:
        sections.append([])
    return sections[0], sections[1], sections[2]


def _load_pa_national(key, sheet):
    """
    Baca sebuah sheet PA nasional precomputed dan ambil 3 blok pivotnya
    (LOS, AGE, EDU) lewat `_split_national_sections` — pembagian blok
    dicari dari baris header "Row Labels", bukan nomor baris tetap.

    Parameter:
      key: str, huruf file sumber (dipanggil dengan "c" oleh load_slide6_national
        dan load_slide7_national).
      sheet: str, nama sheet ("SLIDE 6 - PA Sales Nas" atau "SLIDE 7 - PA Coll Nas").

    Return: dict {"LOS": [(label, pa2, pa3, pa4), ...], "AGE": [...], "EDU": [...]}.
    """
    rows = _sheet_rows(key, sheet)
    los_pa, age_pa, edu_pa = _split_national_sections(rows, value_cols=(1, 2, 3))
    return {"LOS": los_pa, "AGE": age_pa, "EDU": edu_pa}


# ---------------------------------------------------------------------------
# SLIDE 10 – Tabel training BSC
# ---------------------------------------------------------------------------
def load_slide10_training():
    """
    Baca sheet "Sheet1" file d (asumsi layout 2 tabel berdampingan dengan
    baris/kolom tetap, lihat layout di bawah) dan susun angka realisasi
    (actual) vs rencana (plan) training bulanan, plus % pencapaian YTD, per
    region/label.

    Layout tabel di "Sheet1" (1-indexed):
      Tabel 1 (actual): B4:H19  — B=label region, C:H=Jan..Total (header baris 4, data baris 5-19)
      Tabel 2 (plan):   J4:P19  — J=label region, K:P=Jan..Total
      YTD ACH:          R5:R19

    Parameter: tidak ada.

    Return: dict region_atau_label → {
        'actual':  [Jan, Feb, Mar, Apr, May, Total],
        'plan':    [Jan, Feb, Mar, Apr, May, Total],
        'ytd_ach': value,
    }
    Key selain nama region kanonis (lewat REGION_TRAINING) juga mencakup
    label khusus "HO/Sentralisasi", "National", "Regional". Dipanggil oleh
    load_all() dan disimpan sebagai data["s10"]; dipakai oleh
    _apply_slide10_table() di pptx_updater.py untuk mengisi tabel training
    di slide 10.
    """
    rows = _sheet_rows("d", "Sheet1")
    # 0-indexed: baris data 5-19 (1-indexed) -> rows[4:19]
    SPECIAL_LABELS = ("HO/Sentralisasi", "National", "Regional")
    result = {}
    for r in rows[3:20]:
        if not r or not r[1]:
            continue
        entry = {
            "actual":  [r[2], r[3], r[4], r[5], r[6], r[7]],
            "plan":    [r[10], r[11], r[12], r[13], r[14], r[15]],
            "ytd_ach": r[17],
        }
        if r[1] in SPECIAL_LABELS:
            result[r[1]] = entry
            continue
        for canon, abbrev in REGION_TRAINING.items():
            if r[1] == abbrev:
                result[canon] = entry
    return result


# ---------------------------------------------------------------------------
# SLIDE 12 – Scatter chart (Ach Prody vs ROA)
# ---------------------------------------------------------------------------
def load_slide12_scatter():
    """
    Baca 3 sheet dari file e (asumsi layout kolom tetap tiap sheet) dan
    susun titik-titik scatter (Ach Prody vs ROA) per cabang dan per region
    untuk chart 42-43.

    Parameter: tidak ada.

    Return: dict region → {
        'branches': [(ach_prody, roa, med_ach, med_roa, label_short, label_full), ...],
    }
    ditambah key "ALL_REGIONS": [(ach_prody, roa, thresh_ach, thresh_roa, label), ...].
    Dipanggil oleh load_all() dan disimpan sebagai data["s12"]; dipakai oleh
    chart42_scatter_region() (via data["s12"]["ALL_REGIONS"]) dan
    chart43_scatter_branch() (via data["s12"][region]) di chart_data.py.

    KETERBATASAN: sheet "DATA BUBBLE BRANCH BY REGION" tidak punya kolom
    region, jadi `branch_region_map` yang dibangun dari "Std NPAT ROA" di
    bawah ini SEBENARNYA DIHITUNG TAPI TIDAK DIPAKAI — setiap region saat ini
    diberi data cabang yang SAMA (seluruh cabang, lihat "ALL" di bawah),
    bukan hanya cabang di region tersebut.
    """
    rows_branch = _sheet_rows("e", "DATA BUBBLE BRANCH BY REGION")
    rows_region = _sheet_rows("e", "DATA BUBBLE REGION NASIONAL")
    rows_std    = _sheet_rows("e", "Std NPAT ROA")

    # Bangun kolom REGIONAL untuk tiap cabang dari sheet Std NPAT ROA
    branch_region_map = {}
    for r in rows_std[3:]:  # lewati header
        if r[0] and r[2]:  # KODE CABANG dan REGIONAL
            branch_region_map[str(r[0]).strip()] = str(r[2]).strip()

    # Data scatter per cabang
    region_branches = {r: [] for r in REGIONS}
    for r in rows_branch[1:]:  # lewati header
        if r[0] is None:
            continue
        ach  = r[0]
        roa  = r[1]
        m_ach = r[2]
        m_roa = r[3]
        lbl_short = r[8]
        lbl_full  = r[9]
        # Coba cocokkan cabang ke region dari sheet Std
        # Pendekatan: pakai label kalau memungkinkan
        region_branches.setdefault("ALL", []).append((ach, roa, m_ach, m_roa, lbl_short, lbl_full))

    # Data scatter region (sama untuk semua output PPTX)
    all_regions = []
    for r in rows_region[1:]:
        if r[0] is not None:
            all_regions.append((r[0], r[1], r[2], r[3], r[9]))

    result = {"ALL_REGIONS": all_regions}
    # Data cabang: untuk saat ini pakai SEMUA cabang (scatter seharusnya
    # spesifik per region di template, tapi kita tidak punya kolom region di
    # sheet DATA BUBBLE BRANCH BY REGION) — kita berikan semua cabang dan
    # catat keterbatasan ini (lihat catatan di docstring di atas)
    for region in REGIONS:
        result[region] = region_branches.get("ALL", [])

    return result


# ---------------------------------------------------------------------------
# SLIDE 14-16 – Tabel-tabel attrition
# ---------------------------------------------------------------------------
def load_attrition():
    """
    Baca sheet "SLIDE 14 - Attr YoY May" file f1 (asumsi layout kolom tetap:
    3 blok kolom independen untuk non-regret/regret/total, masing-masing
    dengan kolom "Region"-nya sendiri) untuk angka YoY attrition per region,
    lalu tambahkan tabel "Top 3 Reason Out" per region lewat pemanggilan
    internal ke `load_slide14_reason_out`.

    Parameter: tidak ada.

    Return: dict region_atau_label → {
        'yoy_non_regret_25': float, 'yoy_non_regret_26': float,
        'yoy_regret_25': float,     'yoy_regret_26': float,
        'yoy_total_25': float,      'yoy_total_26': float,
        'reason_out_rows': [...],   # lihat load_slide14_reason_out
    }
    Key selain nama region kanonis juga mencakup label ringkasan "Head
    Office" dan "Nasional" (tidak punya reason_out_rows karena loop kedua
    hanya mengisi region di REGIONS). Dipanggil oleh load_all() dan
    disimpan sebagai data["attrition"]; dipakai oleh _apply_slide14_tables()
    (3 tabel YoY) dan _apply_slide14_reason_table() (via
    data["attrition"][region]["reason_out_rows"]) di pptx_updater.py.
    """
    result = {}

    # Slide 14 – Tabel YoY (region + baris ringkasan "Head Office" + "Nasional").
    # Masing-masing dari 3 kelompok metrik (non-regret / regret / total) punya
    # kolom "Region"-nya sendiri (0, 4, 8) dan diurutkan secara independen, jadi
    # nilai tiap kelompok HARUS dipasangkan dengan kolom region MILIKNYA
    # SENDIRI, bukan kolom bersama — memasangkan semuanya terhadap kolom 0
    # akan diam-diam salah pasang 2 dari 3 kelompok setiap kali urutan sortnya
    # berbeda dari urutan sort kolom 0.
    rows14 = _sheet_rows("f1", "SLIDE 14 - Attr YoY May")
    valid_labels = set(REGIONS) | {"Head Office", "Nasional"}
    for r in rows14[2:]:
        groups = [
            (r[0], {"yoy_non_regret_25": r[1], "yoy_non_regret_26": r[2]}),
            (r[4], {"yoy_regret_25":     r[5], "yoy_regret_26":     r[6]}),
            (r[8], {"yoy_total_25":      r[9], "yoy_total_26":      r[10]}),
        ]
        for region, vals in groups:
            if region and region in valid_labels:
                result.setdefault(region, {}).update(vals)

    # Slide 14 – Tabel Top 3 Reason Out, dihitung per region dari pivot cache
    for region in REGIONS:
        result.setdefault(region, {})["reason_out_rows"] = load_slide14_reason_out(region)

    return result


@functools.lru_cache(maxsize=None)
def _sheet_cache_nums(xlsx_key, sheet_name):
    """
    Temukan nomor pivotCacheDefinition untuk SETIAP PivotTable yang
    ditempatkan pada `sheet_name`, dengan menelusuri rantai relasi OOXML
    secara dinamis — BUKAN dengan angka yang di-hardcode, karena nomor ini
    berubah setiap kali file di-save-ulang oleh Excel.

    Cara kerja teknisnya, langkah demi langkah, murni lewat regex atas isi
    file .xlsx (yang sebenarnya adalah arsip ZIP berisi XML):
      1. Cari elemen <sheet name="..."> di xl/workbook.xml untuk mendapatkan
         atribut r:id (mis. "rId5") milik sheet tersebut.
      2. Cari r:id itu di xl/_rels/workbook.xml.rels untuk mendapatkan
         Target-nya — path fisik file sheet (mis. "worksheets/sheet3.xml").
      3. Buka file relasi milik sheet itu sendiri
         (xl/worksheets/_rels/sheet3.xml.rels) dan ambil semua Target yang
         Type-nya "...pivotTable" — ini adalah path ke part pivotTableN.xml
         yang ditempatkan di sheet tsb.
      4. Untuk tiap part pivotTable itu, baca atribut cacheId="..." di
         dalamnya.
      5. Cari cacheId itu di blok <pivotCaches> pada xl/workbook.xml untuk
         mendapatkan r:id relasinya (elemen <pivotCache cacheId="N"
         r:id="rIdM">).
      6. Cari r:id itu lagi di xl/_rels/workbook.xml.rels untuk mendapatkan
         Target-nya — path ke pivotCacheDefinitionK.xml — lalu ambil angka K.

    Hasil per (xlsx_key, sheet_name) di-cache selamanya lewat
    @functools.lru_cache karena isi file tidak berubah dalam satu proses
    generate-laporan, dan chain regex di atas cukup mahal untuk dipanggil
    berulang-ulang oleh setiap loader region.

    Parameter:
      xlsx_key: str, huruf file sumber ("a".."g", untuk fungsi ini dipanggil
        dengan "f1" dan "g" oleh loader-loader di bawah).
      sheet_name: str, nama sheet yang memuat PivotTable, persis seperti di
        Excel (mis. "SLIDE 14 - Reason Out - Region").

    Return: list[str] berisi nomor cache (sebagai string, mis. ["12"]) untuk
    setiap PivotTable di sheet tsb, urutan sesuai urutan pivotTable part
    ditemukan di file relasi sheet; list kosong jika sheet/relasi tidak
    ditemukan. Dipakai secara internal oleh `_find_cache_with_field`.
    """
    z = zipfile.ZipFile(XLSX_FILES[xlsx_key])
    try:
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

        m = re.search(rf'<sheet name="{re.escape(sheet_name)}"[^>]*r:id="(rId\d+)"', wbxml)
        if not m:
            return []
        sheet_target = re.search(rf'Id="{m.group(1)}"[^>]*Target="([^"]+)"', rels).group(1)
        sheet_path = "xl/" + sheet_target
        sheet_dir = posixpath.dirname(sheet_path)
        sheet_rels_path = f"{sheet_dir}/_rels/{posixpath.basename(sheet_path)}.rels"
        if sheet_rels_path not in z.namelist():
            return []

        srels = z.read(sheet_rels_path).decode("utf-8")
        pt_targets = re.findall(r'Type="[^"]*pivotTable"\s+Target="([^"]+)"', srels)
        caches_block = re.search(r'<pivotCaches>.*?</pivotCaches>', wbxml, re.DOTALL)
        if not caches_block:
            return []

        nums = []
        for t in pt_targets:
            pt_path = posixpath.normpath(posixpath.join(sheet_dir, t))
            cache_id = re.search(r'cacheId="(\d+)"', z.read(pt_path).decode("utf-8")).group(1)
            cache_rid = re.search(rf'<pivotCache cacheId="{cache_id}" r:id="(rId\d+)"',
                                   caches_block.group(0)).group(1)
            cache_target = re.search(rf'Id="{cache_rid}"[^>]*Target="([^"]+)"', rels).group(1)
            nums.append(re.search(r'pivotCacheDefinition(\d+)\.xml', cache_target).group(1))
        return nums
    finally:
        z.close()


@functools.lru_cache(maxsize=None)
def _pivot_cache_field_names(xlsx_key, cache_num):
    """
    Ambil daftar nama field (kolom) sebuah pivot cache, langsung dari
    atribut name="..." pada tiap elemen <cacheField> di
    xl/pivotCache/pivotCacheDefinition{cache_num}.xml — ini adalah nama
    kolom ASLI dari data sumber pivot (bukan nama field hasil rename di
    tampilan pivot table).

    Parameter:
      xlsx_key: str, huruf file sumber.
      cache_num: str, nomor cache seperti dikembalikan oleh `_sheet_cache_nums`.

    Return: list[str], nama field sesuai urutan kemunculan di XML (urutan
    ini penting karena `_parse_pivot_cache` memetakan record berdasarkan
    posisi indeks, bukan nama). Hasilnya di-cache lewat lru_cache. Dipakai
    secara internal oleh `_find_cache_with_field`.
    """
    z = zipfile.ZipFile(XLSX_FILES[xlsx_key])
    try:
        def_xml = z.read(f"xl/pivotCache/pivotCacheDefinition{cache_num}.xml").decode("utf-8")
    finally:
        z.close()
    return re.findall(r'<cacheField name="([^"]*)"', def_xml)


@functools.lru_cache(maxsize=None)
def _find_cache_with_field(xlsx_key, sheet_name, field_name_substr):
    """
    Dari semua PivotTable cache pada `sheet_name` (lewat `_sheet_cache_nums`),
    pilih yang mana yang punya sebuah field bernama mengandung
    `field_name_substr` (dicocokkan tanpa peduli huruf besar/kecil). Ini
    diperlukan karena satu sheet Excel kadang menampung LEBIH DARI SATU
    PivotTable (mis. satu untuk data exit/keluar, satu untuk snapshot
    headcount) yang cache-nya harus dibedakan oleh isi field-nya, bukan oleh
    urutan/posisi.

    Parameter:
      xlsx_key: str, huruf file sumber (dipanggil dengan "f1" oleh loader
        slide 14-16, dan "g" oleh load_slide18_fraud).
      sheet_name: str, nama sheet yang memuat satu atau lebih PivotTable.
      field_name_substr: str, potongan nama field yang jadi ciri cache yang
        dicari, mis. "reason cat" (cache data exit) atau "bulan lapor"
        (cache snapshot headcount) atau "potential losses" (cache fraud).

    Return: str nomor cache (lihat `_sheet_cache_nums`) dari cache pertama
    yang cocok, atau None jika tidak ada cache yang punya field tersebut.
    Hasilnya di-cache lewat lru_cache. Dipakai oleh semua loader pivot-cache
    di bawah (load_slide14_reason_out, _read_slide15_function_table,
    _slide16_branch_rows, _compute_slide16_cluster_table,
    load_slide16_reason_out, load_slide18_fraud) untuk menemukan cache yang
    tepat sebelum memanggil `_parse_pivot_cache`.
    """
    needle = field_name_substr.lower()
    for num in _sheet_cache_nums(xlsx_key, sheet_name):
        if any(needle in f.lower() for f in _pivot_cache_field_names(xlsx_key, num)):
            return num
    return None


@functools.lru_cache(maxsize=None)
def _parse_pivot_cache(xlsx_key, cache_num):
    """
    Parse cache sebuah PivotTable Excel (pasangan file
    pivotCacheDefinitionN.xml + pivotCacheRecordsN.xml) menjadi list of dict
    {nama_field: nilai} — yaitu BARIS-BARIS SUMBER MENTAH DAN TIDAK
    TERFILTER di balik pivot tersebut, tidak peduli filter/slice apa pun
    yang sedang aktif di tampilan pivot table yang tersimpan. Ini yang
    membuat teknik ini berguna: kita bisa mengakses seluruh dataset asli
    (mis. semua region, semua bulan) yang dipakai untuk membangun pivot,
    lalu memfilter/mengagregasi ulang dengan logika kita sendiri di Python,
    tanpa perlu Excel/COM untuk "membuka kembali" pivot-nya.

    Cara kerja teknisnya:
      - pivotCacheDefinitionN.xml mendefinisikan skema field: nama tiap
        <cacheField>, dan untuk field yang nilainya berulang (mis. field
        kategorikal seperti region/bulan), sebuah daftar <sharedItems> berisi
        nilai unik yang mungkin muncul — mirip kamus dictionary-encoding.
      - pivotCacheRecordsN.xml berisi baris-baris data aktual sebagai
        elemen <r>, di mana tiap child menunjuk field ke-i (berdasarkan
        POSISI, sejajar dengan urutan <cacheField> di definition). Child
        <x v="idx"/> berarti "ambil item ke-idx dari <sharedItems> field
        ini"; child <s>/<n>/<d>/<b> berarti nilai literal langsung (string/
        angka/tanggal/boolean); <m/> berarti nilai kosong (missing).
      - Beberapa field campur tipe pada <sharedItems>-nya (mis. field
        tanggal yang menyimpan baik nilai <d> tanggal ASLI maupun, di
        <fieldGroup> terpisah, label bucket berupa string) — hanya blok
        <sharedItems> yang benar-benar ditunjuk oleh indeks record, jadi
        parser ini HANYA membaca dari situ, bukan dari <fieldGroup>.

    Parameter:
      xlsx_key: str, huruf file sumber.
      cache_num: str, nomor cache seperti dikembalikan `_sheet_cache_nums`/
        `_find_cache_with_field`.

    Return: list[dict], satu dict per baris sumber pivot, key = nama field
    (dari `_pivot_cache_field_names`), value = string hasil html.unescape
    atau None untuk sel kosong. Hasilnya di-cache lewat lru_cache (dataset
    ini bisa besar dan dipakai berulang oleh banyak region). Dipakai oleh
    semua loader di bagian SLIDE 14-16 dan 18 setelah mereka menemukan
    nomor cache yang tepat lewat `_find_cache_with_field`.
    """
    z = zipfile.ZipFile(XLSX_FILES[xlsx_key])
    try:
        def_xml = z.read(f"xl/pivotCache/pivotCacheDefinition{cache_num}.xml").decode("utf-8")
        rec_xml = z.read(f"xl/pivotCache/pivotCacheRecords{cache_num}.xml").decode("utf-8")
    finally:
        z.close()

    field_names = re.findall(r'<cacheField name="([^"]*)"', def_xml)
    field_bodies = re.findall(r'<cacheField name="[^"]*"[^>]*>(.*?)</cacheField>', def_xml, re.DOTALL)
    shared = []
    for body in field_bodies:
        m = re.search(r'<sharedItems\b[^>]*>(.*?)</sharedItems>', body, re.DOTALL)
        items = ([html.unescape(v) for v in re.findall(r'<(?:s|n|d|b) v="([^"]*)"', m.group(1))]
                 if m else [])
        shared.append(items)

    child_re = re.compile(r'<(s|n|d|x|b|e|m)(?:\s+v="([^"]*)")?\s*/>')
    records = re.findall(r'<r>(.*?)</r>', rec_xml, re.DOTALL)

    rows = []
    for r in records:
        row = {}
        for i, (tag, val) in enumerate(child_re.findall(r)):
            name = field_names[i] if i < len(field_names) else f"_col{i}"
            if tag == "x":
                idx = int(val)
                row[name] = shared[i][idx] if i < len(shared) and idx < len(shared[i]) else None
            elif tag == "m":
                row[name] = None
            else:
                row[name] = html.unescape(val)
        rows.append(row)
    return rows


def load_slide14_reason_out(region):
    """
    Hitung tabel "Top 3 Reason Out Attrition" untuk `region`, dari
    baris-baris exit record mentah di balik pivot "SLIDE 14 - Reason Out -
    Region" file f1 (cache-nya menampung SEMUA 12 region sekaligus — fungsi
    ini mengambil satu cache lewat `_find_cache_with_field` +
    `_parse_pivot_cache`, lalu memfilter sendiri baris milik `region` di
    Python). Untuk tiap 2 kategori (Involuntary/Voluntary): baris total
    kategori itu, plus 3 alasan (reason) individual dengan jumlah terbanyak;
    ditutup dengan satu baris Grand Total keseluruhan.

    Parameter:
      region: str, nama region kanonis (key di REGIONS/REGION_WILAYAH).
        Dipanggil dengan tiap region dari load_attrition() (loop `for region
        in REGIONS`).

    Return: list[dict] {label, non_regret, regret, total, is_header} — satu
    dict per baris tabel, dalam urutan tampil (header Involuntary, top-3
    alasannya, header Voluntary, top-3 alasannya, Grand Total). Dipanggil
    oleh load_attrition() dan disimpan sebagai
    data["attrition"][region]["reason_out_rows"]; dipakai oleh
    _apply_slide14_reason_table() di pptx_updater.py untuk mengisi tabel ke-4
    di slide 14.

    CATATAN: persentase adalah porsi tiap alasan dari TOTAL EXIT REGION ITU
    SENDIRI (Grand Total berjumlah 100%), BUKAN attrition rate perusahaan
    (exit / headcount) — basis persentase asli pivot Excel-nya tidak bisa
    direkonstruksi hanya dari cache saja.
    """
    cache_num = _find_cache_with_field("f1", "SLIDE 14 - Reason Out - Region", "reason cat")
    if cache_num is None:
        return []
    raw = _parse_pivot_cache("f1", cache_num)
    wilayah = REGION_WILAYAH[region]
    recs = [r for r in raw if r.get("Region Business") == wilayah]
    total = len(recs)
    if total == 0:
        return []

    def pct_row(label, sub_recs, is_header=False):
        nr = sum(1 for r in sub_recs if r.get("Regret / Non Regret New") == "non regret")
        rg = sum(1 for r in sub_recs if r.get("Regret / Non Regret New") == "Regret")
        return {
            "label": label,
            "non_regret": nr / total,
            "regret": rg / total,
            "total": (nr + rg) / total,
            "is_header": is_header,
        }

    rows = []
    for cat in ("Involuntary", "Voluntary"):
        cat_recs = [r for r in recs if r.get("Reason Cat") == cat]
        rows.append(pct_row(cat, cat_recs, is_header=True))
        by_reason = {}
        for r in cat_recs:
            by_reason.setdefault(r.get("Reason"), []).append(r)
        top3 = sorted(by_reason.items(), key=lambda kv: -len(kv[1]))[:3]
        for reason, sub_recs in top3:
            rows.append(pct_row(reason, sub_recs))

    rows.append(pct_row("Grand Total", recs, is_header=True))
    return rows


# Daftar baris tabel "Attrition Report by Function" beserta cara
# pencocokannya: "broad" mencocokkan field Function 2 (kategori luas seperti
# "Sales"/"CREDIT"/dst.), "detail" mencocokkan field Function 1 (jabatan
# spesifik seperti "Sales Officer"), "grand_total" menjumlahkan semua baris
# broad sekaligus (lihat _SLIDE15_GRAND_TOTAL_FUNCTIONS).
_SLIDE15_ROWS = [
    ("Sales",               "broad", "Sales"),
    ("Sales Officer",       "detail", "Sales Officer"),
    ("Sales Support",       "detail", "Sales Support"),
    ("CREDIT",              "broad", "CREDIT"),
    ("COLLECTION",          "broad", "COLLECTION"),
    ("Collection Officer",  "detail", "Collection Officer"),
    ("Collection Support",  "detail", "Collection Support"),
    ("OPERATION",           "broad", "OPERATION"),
    ("Grand Total",         "grand_total", None),
]
_SLIDE15_GRAND_TOTAL_FUNCTIONS = {"Sales", "CREDIT", "COLLECTION", "OPERATION"}


def _read_slide15_function_table(xlsx_key, sheet_name, region):
    """
    Hitung baris-baris "Attrition Report by Function" untuk `region`, dari
    2 PivotTable cache yang ada di sheet yang sama: satu cache exit-record
    (berisi klasifikasi Reason Cat / Regret, dipakai untuk kolom Out
    NonRegret/Regret/Total) dan satu cache snapshot karyawan aktif (log
    per-karyawan per-bulan-lapor, dipakai untuk kolom Active Jan/May
    headcount). Kedua cache ditemukan lewat `_find_cache_with_field` (dicari
    di sheet yang sama, dibedakan lewat field cirinya masing-masing:
    "reason cat" untuk cache exit, "bulan lapor" untuk cache headcount) lalu
    dibaca lewat `_parse_pivot_cache`. Nama field sedikit berbeda kapitalisasi
    antar file (mis. "Function 1" vs "function 1", "Bulan Lapor" vs "Bulan
    lapor"), jadi pencarian key dilakukan tanpa peduli huruf besar/kecil
    lewat helper `field()` di bawah.

    Rata-rata headcount (avg_hc) dihitung dari rata-rata jumlah karyawan
    aktif pada bulan snapshot paling awal dan paling akhir yang ada di
    cache headcount (month_early/month_late — bukan berarti selalu Januari
    dan Mei, hanya bulan pertama dan terakhir yang benar-benar tercatat).

    Parameter:
      xlsx_key: str, huruf file sumber ("f1" untuk data y26, "f2" untuk y25
        — lihat load_slide15_function).
      sheet_name: str, nama sheet berisi kedua PivotTable ("SLIDE 15 Atas -
        Region Function" atau "SLIDE 15 Bawah - region functio").
      region: str, nama region kanonis.

    Return: list[dict] {label, avg_hc, out_nr, out_rg, out_total, pct_nr,
    pct_rg, pct_total} — satu dict per baris di _SLIDE15_ROWS, urutan
    dipertahankan sesuai list itu. List kosong jika salah satu cache tidak
    ditemukan. Dipanggil oleh load_slide15_function() (2 kali, untuk y26 dan
    y25) yang disimpan sebagai data["s15_func"][region]["y26"/"y25"], lalu
    dipakai oleh _apply_slide15_tables() dan _apply_slide15_indicators() di
    pptx_updater.py, serta oleh _apply_slide16_branch_table() (untuk baris
    grand-total pembanding) via data["s15_func"][region]["y26"].
    """
    exit_cache = _find_cache_with_field(xlsx_key, sheet_name, "reason cat")
    hc_cache = _find_cache_with_field(xlsx_key, sheet_name, "bulan lapor")
    if exit_cache is None or hc_cache is None:
        return []

    exit_raw = _parse_pivot_cache(xlsx_key, exit_cache)
    hc_raw = _parse_pivot_cache(xlsx_key, hc_cache)

    def field(rows, *candidates):
        """Cari `candidates` mana yang benar-benar menjadi key di baris-baris ini (tanpa peduli huruf besar/kecil)."""
        if not rows:
            return candidates[0]
        keys = {k.lower(): k for k in rows[0].keys()}
        for c in candidates:
            if c.lower() in keys:
                return keys[c.lower()]
        return candidates[0]

    exit_region_f = field(exit_raw, "Region Business")
    exit_f1 = field(exit_raw, "Function 1", "function 1")
    exit_f2 = field(exit_raw, "Function 2", "function 2")
    exit_regret_f = field(exit_raw, "Regret / Non Regret New")

    hc_region_f = field(hc_raw, "Region Business (FIX)", "Region Business")
    hc_f1 = field(hc_raw, "FUNCTION 1", "function 1")
    hc_f2 = field(hc_raw, "FUNCTION 2", "function 2")
    hc_month_f = field(hc_raw, "Bulan Lapor", "Bulan lapor")

    wilayah = REGION_WILAYAH[region]
    exit_recs = [r for r in exit_raw if r.get(exit_region_f) == wilayah]
    hc_recs = [r for r in hc_raw if r.get(hc_region_f) == wilayah]

    months = sorted({r.get(hc_month_f) for r in hc_recs if r.get(hc_month_f)})
    month_early = months[0] if months else None
    month_late = months[-1] if len(months) > 1 else month_early

    def match_fn(rows, f1_field, f2_field, kind, value):
        if kind == "broad":
            return [r for r in rows if r.get(f2_field) == value]
        if kind == "detail":
            return [r for r in rows if r.get(f1_field) == value]
        return [r for r in rows if r.get(f2_field) in _SLIDE15_GRAND_TOTAL_FUNCTIONS]

    result = []
    for label, kind, value in _SLIDE15_ROWS:
        er = match_fn(exit_recs, exit_f1, exit_f2, kind, value)
        hr = match_fn(hc_recs, hc_f1, hc_f2, kind, value)

        out_nr = sum(1 for r in er if r.get(exit_regret_f) in ("non regret", "Non Regret"))
        out_rg = sum(1 for r in er if r.get(exit_regret_f) in ("Regret", "regret"))
        active_jan = sum(1 for r in hr if r.get(hc_month_f) == month_early)
        active_may = sum(1 for r in hr if r.get(hc_month_f) == month_late)
        avg_hc = (active_jan + active_may) / 2

        result.append({
            "label":     label,
            "avg_hc":    avg_hc,
            "out_nr":    out_nr,
            "out_rg":    out_rg,
            "out_total": out_nr + out_rg,
            "pct_nr":    (out_nr / avg_hc) if avg_hc else 0,
            "pct_rg":    (out_rg / avg_hc) if avg_hc else 0,
            "pct_total": ((out_nr + out_rg) / avg_hc) if avg_hc else 0,
        })
    return result


def load_slide15_function(region):
    """
    Gabungkan hasil `_read_slide15_function_table` untuk 2 file/sheet
    berbeda (tahun berjalan dan tahun lalu) menjadi satu dict untuk
    `region`.

    Parameter:
      region: str, nama region kanonis. Dipanggil dari load_all() untuk
        setiap region di REGIONS (comprehension
        `{region: load_slide15_function(region) for region in REGIONS}`).

    Return: {'y26': [...], 'y25': [...]} — masing-masing list of dict sesuai
    bentuk return `_read_slide15_function_table`. y26 dari pivot cache sheet
    "SLIDE 15 Atas - Region Function" file f1, y25 dari sheet "SLIDE 15
    Bawah - region functio" file f2. Disimpan oleh load_all() sebagai
    data["s15_func"][region]; lihat daftar konsumen di docstring
    `_read_slide15_function_table`.
    """
    return {
        "y26": _read_slide15_function_table("f1", "SLIDE 15 Atas - Region Function", region),
        "y25": _read_slide15_function_table("f2", "SLIDE 15 Bawah - region functio", region),
    }


# ---------------------------------------------------------------------------
# SLIDE 16 – Field Regretted Attrition (tabel branch/cluster + reason-out)
# ---------------------------------------------------------------------------
def _slide16_branch_rows(sheet_name, group_field):
    """
    Cari dan baca kedua pivot cache (exit-record dan snapshot headcount)
    milik `sheet_name`, lalu tentukan nama field aktual (case-insensitive)
    yang dipakai untuk region, Function 1, grup (`group_field` — mis.
    "DEPT / BRANCH"), status regret, dan bulan lapor. Fungsi persiapan
    (tidak melakukan agregasi sendiri) yang dipakai bersama oleh
    `_compute_slide16_table`, sebelum agregasi per-region/per-fungsi
    dilakukan di sana.

    Parameter:
      sheet_name: str, nama sheet berisi kedua PivotTable (mis. "SLIDE 16 -
        field by branch SSD").
      group_field: str, nama field pengelompokan yang dicari di cache exit
        (dipakai apa adanya) dan di cache headcount (dicoba versi
        UPPERCASE-nya dulu, baru versi asli — lihat `hc_group_f` di bawah).

    Return: tuple (exit_raw, hc_raw, exit_region_f, exit_f1, exit_group_f,
    exit_regret_f, hc_region_f, hc_f1, hc_group_f, hc_month_f) — dua list of
    dict (baris pivot cache mentah) diikuti nama-nama field yang sudah
    diresolusi; atau list kosong `[]` jika salah satu cache tidak
    ditemukan (dicek pemanggil lewat `if not fields`). Dipanggil hanya oleh
    `_compute_slide16_table`.

    (Meski nama fungsi menyebut "branch rows", ia HANYA melakukan resolusi
    cache & nama field — perhitungan baris per-branch/cluster yang
    sebenarnya, diurutkan menurun berdasarkan % Regret, ada di
    `_compute_slide16_table`.)
    """
    exit_cache = _find_cache_with_field("f1", sheet_name, "reason cat")
    hc_cache = _find_cache_with_field("f1", sheet_name, "bulan lapor")
    if exit_cache is None or hc_cache is None:
        return []
    exit_raw = _parse_pivot_cache("f1", exit_cache)
    hc_raw = _parse_pivot_cache("f1", hc_cache)

    def field(rows, *candidates):
        if not rows:
            return candidates[0]
        keys = {k.lower(): k for k in rows[0].keys()}
        for c in candidates:
            if c.lower() in keys:
                return keys[c.lower()]
        return candidates[0]

    exit_region_f = field(exit_raw, "Region Business")
    exit_f1 = field(exit_raw, "Function 1", "function 1")
    exit_group_f = field(exit_raw, group_field)
    exit_regret_f = field(exit_raw, "Regret / Non Regret New")

    hc_region_f = field(hc_raw, "Region Business (FIX)", "Region Business")
    hc_f1 = field(hc_raw, "FUNCTION 1", "function 1")
    hc_group_f = field(hc_raw, group_field.upper(), group_field)
    hc_month_f = field(hc_raw, "Bulan Lapor", "Bulan lapor")

    return exit_raw, hc_raw, exit_region_f, exit_f1, exit_group_f, exit_regret_f, \
        hc_region_f, hc_f1, hc_group_f, hc_month_f


def _compute_slide16_table(sheet_name, region, function1_value, group_field):
    """
    Hitung breakdown per-branch (atau per-cluster/grup lain, tergantung
    `group_field`) untuk satu fungsi Field tertentu (mis. "Sales Officer"),
    dari exit-record dan snapshot-headcount cache yang sudah diresolusi
    lewat `_slide16_branch_rows`. Sama seperti `_read_slide15_function_table`
    tapi grouping-nya per nilai `group_field` (grup dinamis, hasil `set` dari
    data) alih-alih daftar baris tetap `_SLIDE15_ROWS`.

    Parameter:
      sheet_name: str, nama sheet PivotTable (mis. "SLIDE 16 - field by
        branch SSD").
      region: str, nama region kanonis.
      function1_value: str, nilai field "Function 1" untuk memfilter baris
        (mis. "Sales Officer").
      group_field: str, nama field pengelompokan (mis. "DEPT / BRANCH").

    Return: list[dict] {label, avg_hc, out_nr, out_rg, out_total, pct_nr,
    pct_rg, pct_total} — satu dict per nilai grup yang muncul di salah satu
    cache, diurutkan menurun berdasarkan out_rg (jumlah exit regret) dengan
    label sebagai tie-breaker sekunder (union `set` python urutannya tidak
    stabil antar run, jadi perlu secondary sort key ini). List kosong jika
    `_slide16_branch_rows` gagal menemukan salah satu cache. Dipanggil oleh
    load_slide16_branch_sales(); lihat docstring fungsi itu untuk aliran
    datanya ke pptx_updater.py.
    """
    fields = _slide16_branch_rows(sheet_name, group_field)
    if not fields:
        return []
    (exit_raw, hc_raw, exit_region_f, exit_f1, exit_group_f, exit_regret_f,
     hc_region_f, hc_f1, hc_group_f, hc_month_f) = fields

    wilayah = REGION_WILAYAH[region]
    exit_recs = [r for r in exit_raw if r.get(exit_region_f) == wilayah and r.get(exit_f1) == function1_value]
    hc_recs = [r for r in hc_raw if r.get(hc_region_f) == wilayah and r.get(hc_f1) == function1_value]

    months = sorted({r.get(hc_month_f) for r in hc_recs if r.get(hc_month_f)})
    month_early = months[0] if months else None
    month_late = months[-1] if len(months) > 1 else month_early

    groups = {r.get(exit_group_f) for r in exit_recs if r.get(exit_group_f)}
    groups |= {r.get(hc_group_f) for r in hc_recs if r.get(hc_group_f)}

    rows = []
    for g in groups:
        er = [r for r in exit_recs if r.get(exit_group_f) == g]
        hr = [r for r in hc_recs if r.get(hc_group_f) == g]
        out_nr = sum(1 for r in er if r.get(exit_regret_f) in ("non regret", "Non Regret"))
        out_rg = sum(1 for r in er if r.get(exit_regret_f) in ("Regret", "regret"))
        active_jan = sum(1 for r in hr if r.get(hc_month_f) == month_early)
        active_may = sum(1 for r in hr if r.get(hc_month_f) == month_late)
        avg_hc = (active_jan + active_may) / 2
        rows.append({
            "label":     g,
            "avg_hc":    avg_hc,
            "out_nr":    out_nr,
            "out_rg":    out_rg,
            "out_total": out_nr + out_rg,
            "pct_nr":    (out_nr / avg_hc) if avg_hc else 0,
            "pct_rg":    (out_rg / avg_hc) if avg_hc else 0,
            "pct_total": ((out_nr + out_rg) / avg_hc) if avg_hc else 0,
        })

    # Sort key sekunder (label) membuat hasil seri deterministik — `groups`
    # di atas dibangun dari union `set`, yang urutan iterasinya tidak stabil
    # antar run.
    rows.sort(key=lambda r: (-r["out_rg"], r["label"]))
    return rows


def load_slide16_branch_sales(region):
    """
    Breakdown per-branch untuk fungsi Field Sales (Sales Officer),
    dikelompokkan lewat field "DEPT / BRANCH", diurutkan menurun berdasarkan
    jumlah exit regret (lihat `_compute_slide16_table`).

    Parameter:
      region: str, nama region kanonis. Dipanggil dari load_all() lewat
        comprehension `s16 = {region: {"branch_sales": load_slide16_branch_sales(region), ...}}`.

    Return: list[dict], lihat bentuk return `_compute_slide16_table`.
    Disimpan oleh load_all() sebagai data["s16"][region]["branch_sales"];
    dipakai oleh _apply_slide16_branch_table() di pptx_updater.py (dipanggil
    dengan data_key="branch_sales") untuk mengisi salah satu tabel per-branch
    di slide 16.
    """
    return _compute_slide16_table("SLIDE 16 - field by branch SSD", region, "Sales Officer", "DEPT / BRANCH")


def _compute_slide16_cluster_table(sheet_name, region, function1_value):
    """
    Breakdown per-cluster untuk sebuah fungsi Field yang exit record-nya
    TIDAK punya field cluster secara langsung (mis. Collection Officer —
    cache exit-record tidak punya kolom branch/cluster yang terisi untuk
    peran ini). Berbeda dari `_compute_slide16_table` (yang bisa langsung
    mengelompokkan lewat field cluster milik exit record itu sendiri), di
    sini cluster harus DIREKONSTRUKSI dengan sebuah cross-cache join:

      1. Dari cache snapshot headcount, bangun peta NIP -> "COVER AREA"
         (nama field cluster) untuk bulan snapshot paling akhir (May) dan
         paling awal (Jan) yang ditemukan.
      2. Gabungkan kedua peta itu dengan Jan diprioritaskan/menimpa May
         (`{**nip_cluster_may, **nip_cluster_jan}` — dict yang disebut
         terakhir menang) — karena karyawan yang keluar (exit) di antara Jan
         dan May lebih mungkin masih tercatat di snapshot Januari
         dibanding snapshot Mei (saat itu ia mungkin sudah tidak ada lagi
         di data headcount).
      3. Untuk tiap baris exit record, ambil NIP-nya dan "join" ke peta itu
         (`nip_cluster.get(r.get(exit_nip_f))`) untuk tahu cluster mana yang
         harus dikreditkan atas exit tersebut — TANPA field cluster sama
         sekali di sisi cache exit record.

    Ini adalah join best-effort: sejumlah kecil NIP (mis. staf
    vendor/outsourced) sama sekali tidak muncul di log headcount, sehingga
    tidak bisa diatribusikan ke cluster mana pun — baris-baris exit mereka
    otomatis TIDAK IKUT di baris per-cluster manapun di sini, meski tetap
    terhitung di baris Grand Total yang independen (dihitung terpisah oleh
    load_slide15_function/_read_slide15_function_table, tidak terdampak
    celah ini).

    Parameter:
      sheet_name: str, nama sheet PivotTable (dipanggil dengan "SLIDE 16 -
        field by cluster COL" oleh load_slide16_branch_collection).
      region: str, nama region kanonis.
      function1_value: str, nilai field "Function 1" untuk memfilter baris
        (dipanggil dengan "Collection Officer").

    Return: list[dict] {label, avg_hc, out_nr, out_rg, out_total, pct_nr,
    pct_rg, pct_total} — satu dict per nama cluster yang muncul di peta
    NIP->cluster, diurutkan menurun berdasarkan out_rg dengan label sebagai
    tie-breaker. List kosong jika salah satu cache tidak ditemukan.
    Dipanggil oleh load_slide16_branch_collection().
    """
    exit_cache = _find_cache_with_field("f1", sheet_name, "reason cat")
    hc_cache = _find_cache_with_field("f1", sheet_name, "bulan lapor")
    if exit_cache is None or hc_cache is None:
        return []
    exit_raw = _parse_pivot_cache("f1", exit_cache)
    hc_raw = _parse_pivot_cache("f1", hc_cache)

    def field(rows, *candidates):
        if not rows:
            return candidates[0]
        keys = {k.lower(): k for k in rows[0].keys()}
        for c in candidates:
            if c.lower() in keys:
                return keys[c.lower()]
        return candidates[0]

    exit_region_f = field(exit_raw, "Region Business")
    exit_f1 = field(exit_raw, "Function 1", "function 1")
    exit_regret_f = field(exit_raw, "Regret / Non Regret New")
    exit_nip_f = field(exit_raw, "NIP")

    hc_region_f = field(hc_raw, "Region Business (FIX)", "Region Business")
    hc_f1 = field(hc_raw, "FUNCTION 1", "function 1")
    hc_month_f = field(hc_raw, "Bulan Lapor", "Bulan lapor")
    hc_cluster_f = field(hc_raw, "COVER AREA")
    hc_nip_f = field(hc_raw, "NIP")

    wilayah = REGION_WILAYAH[region]
    exit_recs = [r for r in exit_raw if r.get(exit_region_f) == wilayah and r.get(exit_f1) == function1_value]
    hc_recs = [r for r in hc_raw if r.get(hc_region_f) == wilayah and r.get(hc_f1) == function1_value]

    months = sorted({r.get(hc_month_f) for r in hc_recs if r.get(hc_month_f)})
    month_early = months[0] if months else None
    month_late = months[-1] if len(months) > 1 else month_early

    nip_cluster_may = {r.get(hc_nip_f): r.get(hc_cluster_f) for r in hc_recs if r.get(hc_month_f) == month_late}
    nip_cluster_jan = {r.get(hc_nip_f): r.get(hc_cluster_f) for r in hc_recs if r.get(hc_month_f) == month_early}
    nip_cluster = {**nip_cluster_may, **nip_cluster_jan}   # Jan diprioritaskan dibanding May

    groups = {c for c in nip_cluster.values() if c}

    rows = []
    for g in groups:
        er = [r for r in exit_recs if nip_cluster.get(r.get(exit_nip_f)) == g]
        hr = [r for r in hc_recs if r.get(hc_cluster_f) == g]
        out_nr = sum(1 for r in er if r.get(exit_regret_f) in ("non regret", "Non Regret"))
        out_rg = sum(1 for r in er if r.get(exit_regret_f) in ("Regret", "regret"))
        active_jan = sum(1 for r in hr if r.get(hc_month_f) == month_early)
        active_may = sum(1 for r in hr if r.get(hc_month_f) == month_late)
        avg_hc = (active_jan + active_may) / 2
        rows.append({
            "label":     g,
            "avg_hc":    avg_hc,
            "out_nr":    out_nr,
            "out_rg":    out_rg,
            "out_total": out_nr + out_rg,
            "pct_nr":    (out_nr / avg_hc) if avg_hc else 0,
            "pct_rg":    (out_rg / avg_hc) if avg_hc else 0,
            "pct_total": ((out_nr + out_rg) / avg_hc) if avg_hc else 0,
        })

    rows.sort(key=lambda r: (-r["out_rg"], r["label"]))
    return rows


def load_slide16_branch_collection(region):
    """
    Breakdown per-cluster untuk fungsi Field Collection (Collection
    Officer), direkonstruksi lewat join NIP->COVER AREA (lihat
    `_compute_slide16_cluster_table` untuk mekanismenya), diurutkan menurun
    berdasarkan jumlah exit regret.

    Parameter:
      region: str, nama region kanonis. Dipanggil dari load_all() lewat
        comprehension `s16 = {region: {"branch_collection":
        load_slide16_branch_collection(region), ...}}`.

    Return: list[dict], lihat bentuk return `_compute_slide16_cluster_table`.
    Disimpan oleh load_all() sebagai data["s16"][region]["branch_collection"];
    dipakai oleh _apply_slide16_branch_table() di pptx_updater.py (dipanggil
    dengan data_key="branch_collection") untuk mengisi tabel per-cluster di
    slide 16.
    """
    return _compute_slide16_cluster_table("SLIDE 16 - field by cluster COL", region, "Collection Officer")


def load_slide16_reason_out(region):
    """
    Hitung tabel 5 alasan keluar (exit) sukarela ("regretted": Better
    Job&Benefit, Start Business, Retirement, Family Reason, Back To School)
    untuk Sales Officer vs Collection Officer, plus baris Total Regret. Dari
    pivot cache sheet "SLIDE 16 reason out field" file f1, difilter ke
    Regret/Non Regret = "Regret" dan region ini saja.

    Parameter:
      region: str, nama region kanonis. Dipanggil dari load_all() lewat
        comprehension `s16 = {region: {"reason_out":
        load_slide16_reason_out(region), ...}}`.

    Return: list[dict] {label, sales_officer, collection_officer, total,
    is_header} — satu dict per alasan (5 baris) diikuti satu baris "Total
    Regret". Disimpan oleh load_all() sebagai data["s16"][region]["reason_out"];
    dipakai oleh _apply_slide16_reason_table() di pptx_updater.py untuk
    mengisi tabel reason-out ke-3 di slide 16.
    """
    cache_num = _find_cache_with_field("f1", "SLIDE 16 reason out field", "reason cat")
    if cache_num is None:
        return []
    raw = _parse_pivot_cache("f1", cache_num)
    wilayah = REGION_WILAYAH[region]
    recs = [r for r in raw
            if r.get("Region Business") == wilayah
            and r.get("Regret / Non Regret New") == "Regret"
            and r.get("Function 1") in ("Sales Officer", "Collection Officer")]

    reasons = ["Better Job&Benefit", "Start Business", "Retirement", "Family Reason", "Back To School"]
    rows = []
    for reason in reasons:
        so = sum(1 for r in recs if r.get("Function 1") == "Sales Officer" and r.get("Reason") == reason)
        co = sum(1 for r in recs if r.get("Function 1") == "Collection Officer" and r.get("Reason") == reason)
        rows.append({"label": reason, "sales_officer": so, "collection_officer": co, "total": so + co, "is_header": False})

    total_so = sum(r["sales_officer"] for r in rows)
    total_co = sum(r["collection_officer"] for r in rows)
    rows.append({"label": "Total Regret", "sales_officer": total_so, "collection_officer": total_co,
                 "total": total_so + total_co, "is_header": True})
    return rows

# ---------------------------------------------------------------------------
# SLIDE 18 – Chart fraud rate
# ---------------------------------------------------------------------------
def _pivot_field(rows, *candidates):
    """
    Sama seperti helper `field()` lokal yang dipakai berulang di fungsi
    slide 15/16 (cari `candidates` mana yang benar-benar jadi key di
    baris-baris pivot cache ini, tanpa peduli huruf besar/kecil) — versi ini
    dijadikan fungsi modul-level agar bisa dipakai langsung oleh
    load_slide18_fraud() tanpa didefinisikan ulang secara lokal.

    Parameter:
      rows: list[dict], baris-baris hasil `_parse_pivot_cache`.
      *candidates: str, nama-nama field kandidat, urutan prioritas.

    Return: str, nama key asli (sesuai kapitalisasi sebenarnya di data) dari
    kandidat pertama yang cocok; atau `candidates[0]` apa adanya jika `rows`
    kosong atau tak ada kandidat yang cocok. Dipakai oleh load_slide18_fraud().
    """
    if not rows:
        return candidates[0]
    keys = {k.lower(): k for k in rows[0].keys()}
    for c in candidates:
        if c.lower() in keys:
            return keys[c.lower()]
    return candidates[0]

def load_slide18_fraud():
    """
    Hitung ringkasan kasus fraud per region dari SATU pivot cache di file g
    (sheet "SLIDE 18 Fraud Rate", cache ditemukan lewat field ciri
    "potential losses") yang menampung baris-baris kasus fraud untuk SEMUA
    region dan KEDUA tahun (dibedakan lewat field "Tahun") sekaligus —
    berbeda dari loader slide 14-16 yang butuh join 2 cache, di sini semua
    yang dibutuhkan sudah ada dalam satu cache, difilter per region dan per
    tahun langsung di Python.

    Parameter: tidak ada (loop `for region in REGIONS` dilakukan di dalam
    fungsi ini sendiri, tidak menerima parameter `region` seperti loader
    slide 14-16).

    Return: dict region → {
        'kasus_25': int, 'potloss_25': float,
        'kasus_26': int, 'potloss_26': float,
        'by_location_25': {'Branch SSD': float, 'Cluster Collection': float},
        'by_location_26': {'Branch SSD': float, 'Cluster Collection': float},
        'by_ie_25': {'Internal': float, 'External': float},
        'by_ie_26': {'Internal': float, 'External': float},
        'by_kasus_25': [(kasus, potloss), ...],  # diurutkan menurun, terbesar dulu
        'by_kasus_26': [(kasus, potloss), ...],
        'worst5_branch_25': [(branch_name, potloss), ...],   # top 5, 2025
        'worst5_cluster_25': [(cluster_name, potloss), ...], # top 5, 2025
    }
    Dipanggil oleh load_all() dan disimpan sebagai data["fraud"]; dipakai
    oleh chart44_fraud_summary(), chart45_fraud_by_location(),
    chart46_fraud_by_ie() di chart_data.py, serta oleh
    _apply_slide18_summary() dan _apply_slide18_worst5_table() di
    pptx_updater.py, semuanya via data["fraud"][region].
    """
    cache_num = _find_cache_with_field("g", "SLIDE 18 Fraud Rate", "potential losses")
    raw = _parse_pivot_cache("g", cache_num) if cache_num is not None else []

    region_f    = _pivot_field(raw, "Region")
    year_f      = _pivot_field(raw, "Tahun")
    month_f     = _pivot_field(raw, "Bulan Laporan")
    loc_type_f  = _pivot_field(raw, "Branch / Cluster")
    loc_name_f  = _pivot_field(raw, "Lokasi SLIDE")
    ie_f        = _pivot_field(raw, "Internal / Eksternal / Kolusi")
    kasus_cat_f = _pivot_field(raw, "Kategori Kasus")
    potloss_f   = _pivot_field(raw, "Potential Losses (Gross = O/S)")

    def potloss(r):
        try:
            return float(r.get(potloss_f) or 0)
        except (TypeError, ValueError):
            return 0.0

    def is_ytd_may(r):
        # Nilai "Bulan Laporan" berbentuk seperti "4. April" — slide ini
        # menampilkan angka YTD-May, jadi hanya bulan 1-5 yang dihitung
        # (baris 2025 mencakup satu tahun penuh; baris 2026 secara alami
        # hanya Jan-Mei karena itulah "sekarang").
        try:
            return int((r.get(month_f) or "").split(".")[0]) <= 5
        except ValueError:
            return False

    def summarize(recs):
        by_loc = {lt: sum(potloss(r) for r in recs if r.get(loc_type_f) == lt)
                  for lt in ("Branch SSD", "Cluster Collection")}
        by_ie = {"Internal": 0.0, "External": 0.0}
        for r in recs:
            by_ie["Internal" if r.get(ie_f) == "Internal" else "External"] += potloss(r)
        by_kasus = {}
        for r in recs:
            cat = r.get(kasus_cat_f)
            by_kasus[cat] = by_kasus.get(cat, 0.0) + potloss(r)
        by_kasus_sorted = sorted(by_kasus.items(), key=lambda kv: -kv[1])
        return len(recs), sum(potloss(r) for r in recs), by_loc, by_ie, by_kasus_sorted

    def worst5(recs, loc_type):
        groups = {}
        for r in recs:
            if r.get(loc_type_f) != loc_type or not r.get(loc_name_f):
                continue
            name = r.get(loc_name_f)
            groups[name] = groups.get(name, 0.0) + potloss(r)
        return sorted(groups.items(), key=lambda kv: -kv[1])[:5]

    result = {}
    for region in REGIONS:
        fraud_region = REGION_FRAUD.get(region)
        recs = [r for r in raw if r.get(region_f) == fraud_region and is_ytd_may(r)]
        recs_25 = [r for r in recs if r.get(year_f) == "2025"]
        recs_26 = [r for r in recs if r.get(year_f) == "2026"]

        kasus_25, potloss_25, by_loc_25, by_ie_25, by_kasus_25 = summarize(recs_25)
        kasus_26, potloss_26, by_loc_26, by_ie_26, by_kasus_26 = summarize(recs_26)

        result[region] = {
            "kasus_25":   kasus_25,   "potloss_25": potloss_25,
            "kasus_26":   kasus_26,   "potloss_26": potloss_26,
            "by_location_25": by_loc_25,
            "by_location_26": by_loc_26,
            "by_ie_25":   by_ie_25,
            "by_ie_26":   by_ie_26,
            "by_kasus_25": by_kasus_25,
            "by_kasus_26": by_kasus_26,
            "worst5_branch_25":  worst5(recs_25, "Branch SSD"),
            "worst5_cluster_25": worst5(recs_25, "Cluster Collection"),
        }

    return result


# ---------------------------------------------------------------------------
# Helper internal
# ---------------------------------------------------------------------------
def _load_db_b():
    """
    Muat sheet "DATABASE" file b (database karyawan penuh) sebagai pandas
    DataFrame lewat pd.read_excel.

    Parameter: tidak ada.

    Return: pandas.DataFrame. Dipanggil oleh load_all() (hasilnya dibagi ke
    load_slide4_los_edu_age dan load_slide4_span_of_control lewat parameter
    df_b agar file hanya dibaca sekali), dan juga dipanggil sendiri secara
    internal oleh load_slide4_los_edu_age/load_slide4_span_of_control jika
    dipanggil berdiri sendiri tanpa df_b.
    """
    print("  Loading employee database (file b)...")
    df = pd.read_excel(XLSX_FILES["b"], sheet_name="DATABASE",
                       header=0, engine="openpyxl")
    return df


def _load_active_frontliners():
    """
    Muat sheet "Active Frontliners" file c sebagai pandas DataFrame lewat
    pd.read_excel.

    Parameter: tidak ada.

    Return: pandas.DataFrame. Dipanggil oleh load_all() (hasilnya dibagi ke
    load_slide5_7_data lewat parameter df_c), dan juga dipanggil sendiri
    secara internal oleh load_slide5_7_data jika dipanggil berdiri sendiri
    tanpa df_c.
    """
    print("  Loading Active Frontliners database (file c)...")
    df = pd.read_excel(XLSX_FILES["c"], sheet_name="Active Frontliners",
                       header=0, engine="openpyxl")
    return df


def load_all(verbose=True):
    """
    Muat SELURUH sumber data (memanggil hampir semua fungsi load_* di file
    ini) dan gabungkan hasilnya menjadi satu dict besar per-kategori/slide.
    Ini adalah titik masuk (entry point) utama modul ini.

    Parameter:
      verbose: bool, jika True cetak progres tiap tahap loading ke stdout
        (berguna karena beberapa file besar dan loading bisa memakan
        waktu). Dipanggil dari main.py sebagai `load_all(verbose=True)`.

    Return: dict dengan key: "slide3", "s4_func", "s4_npat", "s4_wc",
    "s4_fl", "s4_lea", "s4_soc", "s57", "s5_nat", "s6_nat", "s7_nat", "s10",
    "s12", "attrition", "s15_func", "s16", "fraud" — masing-masing berisi
    hasil pemanggilan fungsi load_* yang bersangkutan (lihat docstring
    masing-masing fungsi untuk bentuk detail tiap value dan konsumennya di
    chart_data.py/pptx_updater.py). Dipanggil SATU KALI SAJA oleh main.py;
    hasilnya (variabel `data`) diteruskan ke seluruh proses generate PPTX
    (chart_data.get_chart_data(), fungsi-fungsi _apply_slideNN_* di
    pptx_updater.py) untuk semua region, tanpa perlu membaca ulang file
    xlsx untuk tiap region.
    """
    if verbose:
        print("Loading slide 3 YoY data...")
    slide3 = load_slide3_yoy()

    if verbose:
        print("Loading slide 4 function data...")
    s4_func = load_slide4_function()

    if verbose:
        print("Loading slide 4 NPAT/HC data...")
    s4_npat = load_slide4_npat_hc()

    if verbose:
        print("Loading slide 4 work contract data...")
    s4_wc = load_slide4_work_contract()

    if verbose:
        print("Loading slide 4 frontliners data...")
    s4_fl = load_slide4_frontliners()

    if verbose:
        print("Loading employee databases (may take a moment)...")
    df_b = _load_db_b()
    df_c = _load_active_frontliners()

    if verbose:
        print("Computing LOS/EDU/AGE distributions...")
    s4_lea = load_slide4_los_edu_age(df_b)

    if verbose:
        print("Computing span of control...")
    s4_soc = load_slide4_span_of_control(df_b)

    if verbose:
        print("Computing slides 5-7 frontliner data...")
    s57 = load_slide5_7_data(df_c)

    if verbose:
        print("Loading national reference data for slides 5-7...")
    s5_nat = load_slide5_national()
    s6_nat = load_slide6_national()
    s7_nat = load_slide7_national()

    if verbose:
        print("Loading training data (slide 10)...")
    s10 = load_slide10_training()

    if verbose:
        print("Loading scatter data (slide 12)...")
    s12 = load_slide12_scatter()

    if verbose:
        print("Loading attrition data (slides 14-16)...")
    attrition = load_attrition()

    if verbose:
        print("Loading attrition by function data (slide 15)...")
    s15_func = {region: load_slide15_function(region) for region in REGIONS}

    if verbose:
        print("Loading field attrition data (slide 16)...")
    s16 = {region: {
        "branch_sales":      load_slide16_branch_sales(region),
        "branch_collection": load_slide16_branch_collection(region),
        "reason_out":        load_slide16_reason_out(region),
    } for region in REGIONS}

    if verbose:
        print("Loading fraud data (slide 18)...")
    fraud = load_slide18_fraud()

    return {
        "slide3":       slide3,
        "s4_func":      s4_func,
        "s4_npat":      s4_npat,
        "s4_wc":        s4_wc,
        "s4_fl":        s4_fl,
        "s4_lea":       s4_lea,
        "s4_soc":       s4_soc,
        "s57":          s57,
        "s5_nat":       s5_nat,
        "s6_nat":       s6_nat,
        "s7_nat":       s7_nat,
        "s10":          s10,
        "s12":          s12,
        "attrition":    attrition,
        "s15_func":     s15_func,
        "s16":          s16,
        "fraud":        fraud,
    }
