"""
Titik masuk (entry point): membuat satu file PPTX HR Dashboard untuk setiap region.
Cara pakai:
    python main.py                          # semua 12 region, semua slide (baku)
    python main.py "Jawa Tengah"             # hanya satu region tertentu, semua slide
    python main.py --slides 3,14,16          # semua region, tapi hanya slide 3/14/16
    python main.py "Bali" --slides 3,14,16   # kombinasi: satu region, slide tertentu saja
    python main.py --dry-run                 # hanya memuat data, tanpa membuat PPTX
    python main.py --perf-report out.md      # ubah path laporan performa (baku: performance_report.md)
    python main.py --manual-baseline-min 60  # ubah asumsi waktu manual per region (baku: 45 menit)
    python main.py --no-perf-report          # lewati penulisan laporan performa

CATATAN soal --slides: slide di luar daftar TETAP ADA di file PPTX yang
dihasilkan (bukan dihapus), tapi dibiarkan dalam kondisi ASLI TEMPLATE
(belum diisi data region apa pun) — karena tiap panggilan selalu membangun
dari template dari awal, bukan menambal file hasil generate sebelumnya.
Cocok untuk lihat cepat satu/dua slide tanpa menunggu ke-19 slide diproses.

CATATAN soal laporan performa: lihat `src/perf.py` untuk detail metrik apa
saja yang dikumpulkan (runtime, volume data, error rate, dst.) dan metrik
mana yang secara eksplisit ditandai "N/A" karena memang tidak relevan untuk
arsitektur pipeline ini (tidak ada panggilan LLM/API eksternal).
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
from src.perf import metrics as perf_metrics, write_report as write_perf_report


def main():
    """
    Alur kerja utama aplikasi. Langkah-langkahnya:
      1. Baca argumen command-line: filter region tertentu, flag --dry-run
         (hanya menguji proses pemuatan data), dan/atau --slides N,M,... —
         daftar nomor slide 1-indexed yang mau diproses saja (lihat
         `generate_pptx_for_region(..., slides=...)` di
         `src/pptx_updater.py` untuk penjelasan lengkap apa yang terjadi
         pada slide di luar daftar itu).
      2. Panggil `load_all()` dari `src/data_loader.py` SATU KALI SAJA — hasilnya
         (dict `data`) dipakai bersama oleh semua region, supaya setiap file
         xlsx sumber hanya perlu dibaca sekali walau menghasilkan 12 PPTX.
      3. Untuk setiap region pada `REGIONS` (dari `src/config.py`) yang cocok
         dengan filter, panggil `generate_pptx_for_region()` dari
         `src/pptx_updater.py`, dengan mengoper `region`, `data`, dan
         `slides_filter` yang sama.
      4. Cetak ringkasan berapa region yang berhasil/gagal di akhir.
      5. Tulis laporan performa (`src/perf.py`) ke `--perf-report` (baku
         "performance_report.md"), kecuali `--no-perf-report` dipakai.

    Fungsi ini tidak menerima parameter maupun mengembalikan nilai — ia hanya
    dipanggil sekali di baris paling bawah file ini (blok `if __name__ ==
    "__main__"`), sebagai titik mulai keseluruhan program.
    """
    args = sys.argv[1:]
    dry_run = "--dry-run" in args
    no_perf_report = "--no-perf-report" in args
    args = [a for a in args if a not in ("--dry-run", "--no-perf-report")]

    def _pop_flag_value(flag, cast, default):
        """Ambil nilai `--flag VALUE` dari `args` (buang kedua token dari `args` supaya tidak salah kena parsing filter region di bawah), atau `default` kalau flag tidak dipakai."""
        if flag not in args:
            return default
        idx = args.index(flag)
        if idx + 1 >= len(args):
            print(f"[ERROR] {flag} butuh nilai")
            sys.exit(1)
        raw = args[idx + 1]
        del args[idx:idx + 2]
        try:
            return cast(raw)
        except ValueError:
            print(f"[ERROR] {flag} tidak valid, dapat: {raw!r}")
            sys.exit(1)

    # --slides N,M,... — daftar nomor slide 1-indexed yang mau diproses saja
    # (lihat generate_pptx_for_region(..., slides=...) di src/pptx_updater.py).
    slides_raw = _pop_flag_value("--slides", str, None)
    slides_filter = [int(x.strip()) for x in slides_raw.split(",") if x.strip()] if slides_raw else None

    perf_report_path = _pop_flag_value("--perf-report", str, "performance_report.md")
    manual_baseline_min = _pop_flag_value("--manual-baseline-min", float, 45.0)

    region_filter = [a for a in args if not a.startswith("--")]

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # perf_metrics (src/perf.py) adalah singleton module-level yang dibaca-
    # tulis dari data_loader.py dan pptx_updater.py juga — reset() di sini
    # supaya run ini tidak ikut membawa angka dari pemanggilan main()
    # sebelumnya (mis. dari test/REPL yang memanggil main() berulang).
    perf_metrics.reset()
    perf_metrics.wall_start = time.perf_counter()

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
        with perf_metrics.phase("data_loading"):
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
    if slides_filter:
        print(f"Slides filter: {slides_filter}")
    print("=" * 60)

    results = []
    for i, region in enumerate(targets, 1):
        print(f"\n[{i}/{len(targets)}] {region}")
        t1 = time.time()
        try:
            # generate_pptx_for_region() (didefinisikan di
            # src/pptx_updater.py) menerima nama `region`, dict `data`
            # (hasil load_all() di atas), dan `slides_filter` (None kalau
            # --slides tidak dipakai, berarti semua slide seperti biasa),
            # lalu mengembalikan path file PPTX yang baru dibuat untuk
            # region tersebut. Kegagalan pada satu region ditangkap di
            # sini agar tidak menghentikan region lain.
            out = generate_pptx_for_region(region, data, slides=slides_filter, verbose=True)
        except Exception as e:
            import traceback
            print(f"  [ERROR] {e}")
            traceback.print_exc()
            results.append((region, None, str(e)))
            perf_metrics.record_region(False, type(e).__name__)
            continue
        results.append((region, out, None))
        perf_metrics.record_region(True)
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

    # ------------------------------------------------------------------
    # 4. Tulis laporan performa (src/perf.py) — mengumpulkan metrik yang
    #    sudah dicatat sepanjang run ini (fase waktu, volume data,
    #    tingkat keberhasilan chart/slide/region, puncak RAM) lewat
    #    singleton `perf_metrics`, lalu menuliskannya sebagai Markdown.
    # ------------------------------------------------------------------
    perf_metrics.wall_end = time.perf_counter()
    if not no_perf_report:
        write_perf_report(perf_report_path, manual_baseline_minutes=manual_baseline_min)


if __name__ == "__main__":
    main()
