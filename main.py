"""
Titik masuk (entry point): membuat satu file PPTX HR Dashboard untuk setiap region.
Cara pakai:
    python main.py                     # semua 12 region
    python main.py "Jawa Tengah"       # hanya satu region tertentu
    python main.py --dry-run           # hanya memuat data, tanpa membuat PPTX
"""
import sys
import os
import time

# Perbaikan encoding console di Windows (agar karakter non-ASCII, mis. simbol
# "–", tidak menyebabkan error saat dicetak ke terminal)
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

# Supaya import "from src...." bisa ditemukan meskipun script dijalankan dari
# direktori lain
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.config import REGIONS, OUTPUT_DIR
from src.data_loader import load_all
from src.pptx_updater import generate_pptx_for_region


def main():
    """
    Alur kerja utama aplikasi. Langkah-langkahnya:
      1. Baca argumen command-line (filter region tertentu, dan/atau flag
         --dry-run untuk hanya menguji proses pemuatan data).
      2. Panggil `load_all()` dari `src/data_loader.py` SATU KALI SAJA — hasilnya
         (dict `data`) dipakai bersama oleh semua region, supaya setiap file
         xlsx sumber hanya perlu dibaca sekali walau menghasilkan 12 PPTX.
      3. Untuk setiap region pada `REGIONS` (dari `src/config.py`) yang cocok
         dengan filter, panggil `generate_pptx_for_region()` dari
         `src/pptx_updater.py`, dengan mengoper `region` dan `data` yang sama.
      4. Cetak ringkasan berapa region yang berhasil/gagal di akhir.

    Fungsi ini tidak menerima parameter maupun mengembalikan nilai — ia hanya
    dipanggil sekali di baris paling bawah file ini (blok `if __name__ ==
    "__main__"`), sebagai titik mulai keseluruhan program.
    """
    args = sys.argv[1:]
    dry_run = "--dry-run" in args
    region_filter = [a for a in args if not a.startswith("--")]

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ------------------------------------------------------------------
    # 1. Muat semua data xlsx sekali saja (dipakai bersama oleh semua region)
    # ------------------------------------------------------------------
    t0 = time.time()
    print("=" * 60)
    print("Loading all data sources...")
    print("=" * 60)

    try:
        # load_all() (didefinisikan di src/data_loader.py) membaca SELURUH
        # file xlsx sumber dan mengembalikan satu dict besar `data` berisi
        # semua data yang sudah diproses per region (mis. data["slide3"],
        # data["attrition"], dst). Dict inilah yang nanti dioper ke
        # generate_pptx_for_region() di bawah, untuk SETIAP region.
        data = load_all(verbose=True)
    except Exception as e:
        print(f"\n[ERROR] Failed to load data: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)

    print(f"Data loaded in {time.time()-t0:.1f}s\n")

    if dry_run:
        print("Dry-run mode: skipping PPTX generation.")
        return

    # ------------------------------------------------------------------
    # 2. Buat satu file PPTX untuk setiap region
    # ------------------------------------------------------------------
    targets = [r for r in REGIONS if (not region_filter or r in region_filter)]
    if not targets:
        print(f"No matching regions for filter: {region_filter}")
        print(f"Available: {REGIONS}")
        sys.exit(1)

    print("=" * 60)
    print(f"Generating {len(targets)} PPTX file(s)...")
    print("=" * 60)

    results = []
    for i, region in enumerate(targets, 1):
        print(f"\n[{i}/{len(targets)}] {region}")
        t1 = time.time()
        try:
            # generate_pptx_for_region() (didefinisikan di
            # src/pptx_updater.py) menerima nama `region` dan dict `data`
            # (hasil load_all() di atas), lalu mengembalikan path file PPTX
            # yang baru dibuat untuk region tersebut. Kegagalan pada satu
            # region ditangkap di sini agar tidak menghentikan region lain.
            out = generate_pptx_for_region(region, data, verbose=True)
        except Exception as e:
            import traceback
            print(f"  [ERROR] {e}")
            traceback.print_exc()
            results.append((region, None, str(e)))
            continue
        results.append((region, out, None))
        print(f"  Done in {time.time()-t1:.1f}s -> {out}")

    # ------------------------------------------------------------------
    # 3. Ringkasan hasil
    # ------------------------------------------------------------------
    print("\n" + "=" * 60)
    ok  = [r for r in results if r[2] is None]
    err = [r for r in results if r[2] is not None]
    print(f"Complete: {len(ok)} succeeded, {len(err)} failed.")
    if err:
        for region, _, msg in err:
            print(f"  FAILED: {region} -> {msg}")
    print(f"Total time: {time.time()-t0:.1f}s")
    print("=" * 60)


if __name__ == "__main__":
    main()
