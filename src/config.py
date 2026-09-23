"""
Definisi region dan pemetaan nama region di semua file xlsx.

File ini tidak berisi fungsi — hanya konstanta (path, daftar, dan dict
pemetaan) yang di-import oleh modul lain:
  - `src/data_loader.py` memakai XLSX_FILES, REGIONS, REGION_WILAYAH,
    REGION_TRAINING, REGION_FRAUD, LOS_ORDER, EDU_ORDER, AGE_ORDER,
    FUNCTION_ORDER untuk tahu file mana yang harus dibaca dan bagaimana
    nama region di tiap file harus dicocokkan/diurutkan.
  - `src/pptx_updater.py` memakai TEMPLATE_PATH, OUTPUT_DIR,
    CHART_EMBED_MAP, TEMPLATE_REGION, TEMPLATE_TOTAL_HC, REGION_ABBREV,
    REGION_TRAINING, REGIONS untuk tahu placeholder apa yang harus diganti
    di template.pptx dan file Excel tersemat (embedded) mana yang terkait
    dengan chart nomor berapa.
"""
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# Hanya folder "Input/" yang boleh dibaca sebagai sumber data — folder
# "input/" (huruf kecil) dan "Filtered_Input/" berisi file sumber
# lain/belum difilter dan TIDAK BOLEH dibaca atau dibagikan.
INPUT_DIR = os.path.join(BASE_DIR, "Input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")  # dipakai main.py untuk menentukan lokasi hasil PPTX

TEMPLATE_PATH = os.path.join(INPUT_DIR, "template.pptx")  # dipakai pptx_updater.py (PptxEditor) sebagai berkas dasar setiap PPTX

# Path lengkap ke setiap file xlsx sumber, dikunci dengan huruf singkat
# (a..g). Dipakai oleh hampir semua fungsi loader di data_loader.py lewat
# XLSX_FILES[<huruf>] untuk membuka file yang benar.
XLSX_FILES = {
    "a": os.path.join(INPUT_DIR, "a. SLIDE 03 - 04 - HC Managed by Region - YoY Mei-26_update (NPAT, WC, Function).xlsx"),
    "b": os.path.join(INPUT_DIR, "b. SLIDE 02 - HC Managed FTE ASA Mei-26 (Frontliners s.d Age Cat).xlsx"),
    "c": os.path.join(INPUT_DIR, "c. SLIDE 05 - 07 - HC Managed Frontliners Mei-26.xlsx"),
    "d": os.path.join(INPUT_DIR, "d. SLIDE 10 - 2026 BSC Corporate - Business-Related Training(Mei)-Sent.xlsx"),
    "e": os.path.join(INPUT_DIR, "e. SLIDE 12 - Raw Data Productivity SSD 2026 - Mei 2026 (Aktual)_NPAT&ROA 5M26.xlsx"),
    "f1": os.path.join(INPUT_DIR, "f1. SLIDE 14 - 16 Attrition Details YTD May-26.xlsx"),
    "f2": os.path.join(INPUT_DIR, "f2. SLIDE 15 Attrition Details YTD May-25.xlsx"),
    "g": os.path.join(INPUT_DIR, "g. SLIDE 18 -  Fraud Rate YTD Mei 2026.xlsx"),
}

# Daftar nama region baku (sesuai penulisan di data "Slide 3 - YoY").
# Ini adalah urutan/daftar region — main.py melakukan
# loop atas REGIONS untuk membuat satu PPTX per region, dan hampir semua
# fungsi loader di data_loader.py melakukan loop yang sama untuk mengisi
# data tiap region.
REGIONS = [
    "Jabotabek 2",
    "Bali",
    "Pasima",
    "Sultan",
    "Sumatera 1",
    "Sumatera 2",
    "Sumatera 3",
    "Jawa Timur",
    "Kalimantan",
    "Jawa Tengah",
    "Jabar",
    "Jabotabek 1",
]

# Pemetaan nama baku → format "Wilayah Area X", format yang dipakai di
# sheet DATABASE dan sheet function pada file xlsx. Dipakai secara luas di
# data_loader.py setiap kali perlu memfilter baris berdasarkan region
# (mis. `df[df["Region Business"] == REGION_WILAYAH[region]]`).
REGION_WILAYAH = {r: f"Wilayah Area {r}" for r in REGIONS}

# Pemetaan nama baku → nama singkatan yang dipakai di file training (file d)
REGION_TRAINING = {
    "Jabotabek 2": "Jabodetabek-2",
    "Bali": "BNT",
    "Pasima": "Pasima",
    "Sultan": "Sultan",
    "Sumatera 1": "Sumbagut",
    "Sumatera 2": "Sumbagsel",
    "Sumatera 3": "Sumbagteng",
    "Jawa Timur": "Jatim",
    "Kalimantan": "Kalimantan",
    "Jawa Tengah": "Jateng",
    "Jabar": "Jabar",
    "Jabotabek 1": "Jabodetabek-1",
}

# Pemetaan nama baku → format region pada file fraud, yaitu "N. NamaRegion"
# — sesuai persis dengan nilai literal field "Region" pada pivot cache data
# fraud di file g (penomoran/penamaannya sendiri, berbeda dari
# REGION_TRAINING/REGION_ABBREV). Dipakai oleh load_slide18_fraud() di
# data_loader.py untuk memfilter data fraud per region.
REGION_FRAUD = {
    "Jabotabek 1": "1. Jabodetabekser 1",
    "Jabotabek 2": "2. Jabodetabekser 2",
    "Jabar":       "3. Jawa Barat",
    "Jawa Tengah": "4. Jawa Tengah",
    "Jawa Timur":  "5. Jawa Timur",
    "Bali":        "6. BNT",
    "Sumatera 1":  "7. Sumbagut",
    "Sumatera 3":  "8. Sumbagteng",
    "Sumatera 2":  "9. Sumbagsel",
    "Kalimantan":  "10. Kalimantan",
    "Sultan":      "11. Sultan",
    "Pasima":      "12. Pasima",
}

# Nama singkatan untuk teks pada slide (contoh: "Top 3 Reason Out – Jateng").
# Dipakai oleh pptx_updater.py saat mengganti placeholder teks region di
# slide dengan versi singkatnya.
REGION_ABBREV = {
    "Jabotabek 2": "Jabo 2",
    "Bali": "Bali",
    "Pasima": "Pasima",
    "Sultan": "Sultan",
    "Sumatera 1": "Sumut",
    "Sumatera 2": "Sumsel",
    "Sumatera 3": "Sumteng",
    "Jawa Timur": "Jatim",
    "Kalimantan": "Kal",
    "Jawa Tengah": "Jateng",
    "Jabar": "Jabar",
    "Jabotabek 1": "Jabo 1",
}

# Token placeholder pada template (persis seperti tertulis apa adanya di
# dalam text run template.pptx).
# CATATAN: pada template.pptx yang sekarang, pendekatan data contoh lama
# ("Jawa Tengah", "Jateng", "2.537") sudah diganti dengan token placeholder
# generik. Nama region sekarang diwakili oleh kata literal "REGION" di
# mana-mana (judul, header, dst.), bukan lagi pasangan nama-lengkap/
# singkatan terpisah — jadi TEMPLATE_ABBREV sudah tidak dipakai lagi untuk
# penggantian nama region. Nilainya tetap disimpan di sini untuk jaga-jaga
# kalau revisi template di masa depan memakai token singkatan lagi.
TEMPLATE_REGION = "REGION"  # dicari & diganti oleh pptx_updater.py di hampir semua slide
TEMPLATE_ABBREV = "Jateng"  # saat ini tidak dipakai untuk penggantian — tidak ada token yang cocok di template

# Nilai placeholder anotasi "Total HC Region" pada slide 4, tertanam
# sebagai teks di template (dulu "2.537" pada template data-contoh lama;
# sekarang "1.097" pada template placeholder). Dicari & diganti oleh
# fungsi terkait slide 4 di pptx_updater.py.
TEMPLATE_TOTAL_HC = "1.097"

# Urutan kategori LOS (masa kerja) versi MENURUN (f → a), khusus untuk
# chart10 di slide 4 — template chart10 memang memakai urutan ini.
# Dipakai oleh load_slide4_los_edu_age() di data_loader.py.
LOS_ORDER = ["f. > 20th", "e. 15<x<20 thn", "d. 10<x<15 thn", "c. 5<x<10 thn", "b. 1<x<5 thn", "a. <1 thn"]

# Urutan kategori LOS versi MENAIK (a → f), untuk semua chart frontliner di
# slide 5, 6, dan 7. Template chart 20/21/24/25/32/33/34/41 semuanya memakai
# urutan menaik, dan data nasionalnya (dibaca apa adanya dari sheet) juga
# menaik — jadi bagian regional wajib memakai urutan yang sama, kalau tidak
# grafik regional akan tampil terbalik dibanding grafik nasional di
# sebelahnya. Dipakai oleh load_slide5_7_data() di data_loader.py.
LOS_ORDER_FRONTLINERS = ["a. <1 thn", "b. 1<x<5 thn", "c. 5<x<10 thn", "d. 10<x<15 thn", "e. 15<x<20 thn", "f. > 20th"]

# Urutan kategori pendidikan (EDU)
EDU_ORDER = ["1. SLTA sederajat & di bawahnya", "2. Diploma", "3. Sarjana", "4. Pasca Sarjana"]
EDU_ORDER_SHORT = ["1. SMA ke bawah", "2. Diploma", "3. Sarjana", "4. Pascasarjana"]

# Urutan kategori usia (AGE)
AGE_ORDER = ["a. <26 thn", "b. 26-36 thn", "c. 37-46 thn", "d. 47-50 thn", "e. > 50 th"]

# Urutan fungsi/departemen pada chart 7
FUNCTION_ORDER = ["BISNIS", "SSD", "COLL", "CREDIT", "LAR", "SUPP OPR", "Honorer", "OB Sec"]

# Pemetaan nomor chart → nama file Excel tersemat (embedded) di dalamnya
# (hasil analisis struktur template.pptx). Dipakai oleh
# `_update_chart()`/`_discover_charts()` di pptx_updater.py untuk
# menemukan file workbook tersemat yang harus disinkronkan setiap kali
# cache data sebuah chart diperbarui.
CHART_EMBED_MAP = {
    1: "Microsoft_Excel_Worksheet.xlsx",
    2: "Microsoft_Excel_Worksheet1.xlsx",
    3: "Microsoft_Excel_Worksheet2.xlsx",
    4: "Microsoft_Excel_Worksheet3.xlsx",
    5: "Microsoft_Excel_Worksheet4.xlsx",
    6: "Microsoft_Excel_Worksheet5.xlsx",
    7: "Microsoft_Excel_Worksheet6.xlsx",
    8: "Microsoft_Excel_Worksheet7.xlsx",
    9: "Microsoft_Excel_Worksheet8.xlsx",
    10: "Microsoft_Excel_Worksheet9.xlsx",
    11: "Microsoft_Excel_Worksheet10.xlsx",
    12: "Microsoft_Excel_Worksheet11.xlsx",
    13: "Microsoft_Excel_Worksheet12.xlsx",
    14: "Microsoft_Excel_Worksheet13.xlsx",
    15: "Microsoft_Excel_Worksheet14.xlsx",
    16: "Microsoft_Excel_Worksheet15.xlsx",
    17: "Microsoft_Excel_Worksheet16.xlsx",
    18: "Microsoft_Excel_Worksheet17.xlsx",
    19: "Microsoft_Excel_Worksheet18.xlsx",
    20: "Microsoft_Excel_Worksheet19.xlsx",
    21: "Microsoft_Excel_Worksheet20.xlsx",
    22: "Microsoft_Excel_Worksheet21.xlsx",
    23: "Microsoft_Excel_Worksheet22.xlsx",
    24: "Microsoft_Excel_Worksheet23.xlsx",
    25: "Microsoft_Excel_Worksheet24.xlsx",
    26: "Microsoft_Excel_Worksheet25.xlsx",
    27: "Microsoft_Excel_Worksheet26.xlsx",
    28: "Microsoft_Excel_Worksheet27.xlsx",
    29: "Microsoft_Excel_Worksheet28.xlsx",
    30: "Microsoft_Excel_Worksheet29.xlsx",
    31: "Microsoft_Excel_Worksheet30.xlsx",
    32: "Microsoft_Excel_Worksheet31.xlsx",
    33: "Microsoft_Excel_Worksheet32.xlsx",
    34: "Microsoft_Excel_Worksheet33.xlsx",
    35: "Microsoft_Excel_Worksheet34.xlsx",
    36: "Microsoft_Excel_Worksheet35.xlsx",
    37: "Microsoft_Excel_Worksheet36.xlsx",
    38: "Microsoft_Excel_Worksheet37.xlsx",
    39: "Microsoft_Excel_Worksheet38.xlsx",
    40: "Microsoft_Excel_Worksheet39.xlsx",
    41: "Microsoft_Excel_Worksheet40.xlsx",
    42: "Microsoft_Excel_Worksheet41.xlsx",
    43: "Microsoft_Excel_Worksheet42.xlsx",
    44: "Microsoft_Excel_Worksheet43.xlsx",
    45: "Microsoft_Excel_Worksheet44.xlsx",
    46: "Microsoft_Excel_Worksheet45.xlsx",
}
