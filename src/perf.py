"""
Instrumentasi performa: mengumpulkan metrik selama satu kali proses
`main.py` (pemuatan data + generate PPTX untuk semua region yang diminta),
lalu menuliskannya sebagai laporan (`performance_report.md`) yang
strukturnya mengikuti PERFORMANCE_REP_NEEDS di root project.

DESAIN — kenapa satu objek global (`metrics`), bukan parameter yang
dioper ke mana-mana: banyak titik ukur (baca sheet Excel di
data_loader.py, update chart/slide per region di pptx_updater.py) ada
jauh di dalam fungsi yang sudah ada dan sudah teruji — mengubah semua
signature fungsi itu untuk menerima objek "collector" akan jadi
perubahan besar dan berisiko untuk sekadar menambah pencatatan metrik.
Pola singleton module-level (mirip modul `logging` bawaan Python) jauh
lebih kecil dampaknya: titik ukur cukup `from .perf import metrics` lalu
panggil satu method, tanpa mengubah signature fungsi apa pun.

CATATAN JUJUR soal cakupan: pipeline ini TIDAK memanggil API/LLM eksternal
sama sekali — teks highlight/insight (`get_slideN_highlights()` di
chart_data.py) dihasilkan dari heuristik lokal langsung dari data yang
sama yang dipakai mengisi chart/tabel, bukan lewat model bahasa. Karena
itu beberapa metrik yang diminta (biaya token, latensi API, deteksi
halusinasi) secara struktural TIDAK RELEVAN untuk arsitektur ini — laporan
di bawah menyatakan ini secara eksplisit alih-alih mengarang angka.
"""
import os
import sys
import time
import platform
from collections import defaultdict


def _peak_rss_mb():
    """
    Ambil puncak RSS (resident set size — memori fisik yang benar-benar
    dipakai proses ini, bukan cuma dialokasikan) sejak proses dimulai,
    dalam MB, lewat `resource.getrusage` (modul stdlib, tidak butuh
    dependency tambahan seperti psutil).

    KENAPA ADA PENYESUAIAN PLATFORM: `ru_maxrss` satuannya BEDA antar OS —
    KILOBYTE di Linux, tapi BYTE di macOS/BSD (perbedaan lama di glibc vs
    BSD libc yang tidak pernah diseragamkan) — tanpa penyesuaian ini,
    angka RAM yang dilaporkan di macOS akan salah 1000x lipat lebih besar
    dari yang sebenarnya.

    Return: float MB, atau None kalau modul `resource` tidak tersedia
    (mis. di Windows — modul ini Unix-only di pustaka standar).
    """
    try:
        import resource
    except ImportError:
        return None
    raw = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    if sys.platform == "darwin":
        return raw / (1024 * 1024)   # byte -> MB
    return raw / 1024                # KB -> MB


class PerfMetrics:
    """
    Penampung semua metrik satu kali run `main.py`. Satu instance global
    (`metrics`, didefinisikan di bawah kelas ini) dipakai bersama oleh
    `main.py`, `data_loader.py`, dan `pptx_updater.py` — masing-masing
    modul mengimpornya dan memanggil method pencatatan yang relevan di
    titik yang sudah ada (lihat docstring modul ini soal alasan desainnya).

    `reset()` dipanggil oleh `main.py` di awal `main()` supaya run
    berikutnya (mis. dari REPL/test) tidak ikut membawa angka run
    sebelumnya.
    """

    def __init__(self):
        self.reset()

    def reset(self):
        """Kosongkan semua penghitung — dipanggil di awal tiap run `main()`."""
        self.wall_start = None
        self.wall_end = None
        self.phase_seconds = {}       # nama fase -> total detik
        self._phase_stack = []        # (nama, waktu_mulai) — mendukung nested phase

        self.workbooks_configured = 0     # len(XLSX_FILES), diisi data_loader
        self.sheet_reads = 0               # jumlah pemanggilan _sheet_rows/_pd/_load_db_b/dst
        self.rows_read = 0
        self.cells_read = 0

        self.chart_attempts = 0
        self.chart_successes = 0
        self.chart_failure_reasons = defaultdict(int)
        self.chart_total_seconds = 0.0      # sum of per-chart elapsed time (all attempts, all regions)
        self.chart_seconds_by_num = defaultdict(float)   # chart_num -> total seconds across all regions
        self.chart_count_by_num = defaultdict(int)       # chart_num -> attempts across all regions
        self.chart_max_seconds = 0.0
        self.chart_max_seconds_num = None

        self.slide_attempts = 0
        self.slide_successes = 0
        self.slide_failure_reasons = defaultdict(int)

        self.regions_total = 0
        self.regions_succeeded = 0
        self.regions_failed = 0
        self.region_failure_reasons = defaultdict(int)

    # --- Fase waktu -----------------------------------------------------
    def phase(self, name):
        """Context manager: `with metrics.phase("data_loading"): ...` — akumulatif kalau nama fase yang sama dipanggil berulang (mis. satu fase per region)."""
        return _PhaseTimer(self, name)

    def _start_phase(self, name):
        self._phase_stack.append((name, time.perf_counter()))

    def _end_phase(self, name):
        # cari dari belakang (LIFO) entri dengan nama ini — mendukung
        # panggilan phase() yang sama berulang kali (akumulasi), bukan
        # cuma nested sekali.
        for i in range(len(self._phase_stack) - 1, -1, -1):
            if self._phase_stack[i][0] == name:
                _, start = self._phase_stack.pop(i)
                self.phase_seconds[name] = self.phase_seconds.get(name, 0.0) + (time.perf_counter() - start)
                return
        raise RuntimeError(f"_end_phase({name!r}) tanpa _start_phase yang cocok")

    # --- Ingestion volume -------------------------------------------------
    def record_sheet_read(self, n_rows, n_cols_estimate=0):
        """
        Catat satu sheet berhasil dibaca (dipanggil dari `_sheet_rows`,
        `_pd`, `_load_db_b`, `_load_active_frontliners` di data_loader.py).

        Parameter: `n_rows` (jumlah baris), `n_cols_estimate` (perkiraan
        jumlah kolom — dari `len(row)` baris pertama untuk `_sheet_rows`,
        atau `df.shape[1]` untuk pembacaan pandas) dipakai menghitung
        estimasi jumlah sel (`n_rows * n_cols_estimate`), BUKAN dihitung
        sel-per-sel (terlalu lambat untuk sheet besar, dan tidak perlu
        presisi untuk keperluan laporan volume).
        """
        self.sheet_reads += 1
        self.rows_read += n_rows
        self.cells_read += n_rows * n_cols_estimate

    # --- Chart mapping ------------------------------------------------
    def record_chart(self, success, reason=None, elapsed=None, chart_num=None):
        """
        Dipanggil dari loop update chart di `pptx_updater.py` — satu panggilan
        per chart per region.

        `elapsed` (detik, dari `time.perf_counter()` di sekitar `_update_chart`)
        dan `chart_num` opsional supaya kode lama yang belum diupdate untuk
        mengoper argumen ini tetap jalan — kalau `elapsed` diisi, waktunya
        diakumulasi baik secara total maupun per nomor chart (chart yang sama
        muncul di 12 region, jadi angka per-chart adalah TOTAL/rata-rata
        lintas region, bukan satu region saja).
        """
        self.chart_attempts += 1
        if success:
            self.chart_successes += 1
        else:
            self.chart_failure_reasons[reason or "unknown"] += 1
        if elapsed is not None:
            self.chart_total_seconds += elapsed
            if chart_num is not None:
                self.chart_seconds_by_num[chart_num] += elapsed
                self.chart_count_by_num[chart_num] += 1
            if elapsed > self.chart_max_seconds:
                self.chart_max_seconds = elapsed
                self.chart_max_seconds_num = chart_num

    # --- Slide/table mapping -------------------------------------------
    def record_slide(self, success, reason=None):
        """Dipanggil dari `_safe_apply_slide` di `pptx_updater.py` — satu panggilan per blok slide per region."""
        self.slide_attempts += 1
        if success:
            self.slide_successes += 1
        else:
            self.slide_failure_reasons[reason or "unknown"] += 1

    # --- Region-level (dipanggil dari main.py) --------------------------
    def record_region(self, success, reason=None):
        self.regions_total += 1
        if success:
            self.regions_succeeded += 1
        else:
            self.regions_failed += 1
            self.region_failure_reasons[reason or "unknown"] += 1

    # --- Turunan (dihitung, bukan dicatat langsung) ----------------------
    @property
    def total_seconds(self):
        if self.wall_start is None or self.wall_end is None:
            return 0.0
        return self.wall_end - self.wall_start

    @property
    def chart_success_rate(self):
        return (self.chart_successes / self.chart_attempts) if self.chart_attempts else None

    @property
    def slide_success_rate(self):
        return (self.slide_successes / self.slide_attempts) if self.slide_attempts else None

    @property
    def charts_per_second(self):
        secs = self.phase_seconds.get("rendering", 0.0)
        return (self.chart_attempts / secs) if secs else None

    @property
    def avg_chart_seconds(self):
        return (self.chart_total_seconds / self.chart_attempts) if self.chart_attempts else None

    @property
    def slides_per_minute(self):
        secs = self.total_seconds
        total_slide_ops = self.slide_attempts  # tiap blok slide per region
        return (total_slide_ops / secs * 60) if secs else None


class _PhaseTimer:
    """Context manager kecil untuk `PerfMetrics.phase()` — lihat penjelasan di sana."""
    def __init__(self, metrics_obj, name):
        self._m = metrics_obj
        self._name = name

    def __enter__(self):
        self._m._start_phase(self._name)
        return self

    def __exit__(self, exc_type, exc, tb):
        self._m._end_phase(self._name)
        return False


# Singleton yang dipakai bersama lintas modul (lihat docstring modul ini
# soal alasan desainnya).
metrics = PerfMetrics()


def _fmt(v, suffix="", none_text="N/A", digits=1):
    if v is None:
        return none_text
    return f"{v:,.{digits}f}{suffix}"


def write_report(path, manual_baseline_minutes=45, post_run_intervention_note=None):
    """
    Menulis laporan performa Markdown ke `path`, strukturnya mengikuti 4
    bagian di PERFORMANCE_REP_NEEDS (root project) — dipanggil dari
    `main.py` di akhir `main()`, setelah semua region selesai diproses
    (baik yang berhasil maupun gagal).

    Parameter:
      path: str, path file output (mis. "performance_report.md").
      manual_baseline_minutes: float, ASUMSI waktu penyusunan deck HR
        Dashboard SECARA MANUAL per region (dalam menit) yang dipakai
        sebagai pembanding di bagian "Operational ROI" — nilai ini TIDAK
        BISA diukur otomatis oleh kode ini (itu proses manusia, bukan
        proses program), jadi diteruskan sebagai parameter dan ditandai
        jelas sebagai asumsi input di laporan, bukan hasil pengukuran.
        Default 45 menit adalah estimasi awal — sebaiknya disesuaikan
        lewat `--manual-baseline-min` di CLI `main.py` kalau organisasi
        Anda punya angka baseline yang lebih akurat.
      post_run_intervention_note: str atau None — "Post-Run Intervention
        Rate" TIDAK BISA diukur otomatis sama sekali (perlu manusia
        me-review tiap PPTX hasil generate dan menghitung berapa yang
        butuh perbaikan manual setelahnya) — kalau None, laporan
        mencantumkan placeholder yang menjelaskan ini perlu diisi manual
        setelah proses review.

    Return: None. Efeknya menulis file ke `path` (menimpa kalau sudah
    ada) dan mencetak baris konfirmasi ke stdout.
    """
    m = metrics
    lines = []
    lines.append("# Performance Report")
    lines.append("")
    lines.append(f"- Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"- Platform: {platform.system()} {platform.release()} ({platform.machine()}), Python {platform.python_version()}")
    lines.append(f"- Regions processed: {m.regions_total} (requested)")
    lines.append("")

    # -----------------------------------------------------------------
    lines.append("## 1. Runtime & Throughput Metrics")
    lines.append("")
    lines.append(f"- **Total End-to-End Latency**: {_fmt(m.total_seconds, 's', digits=1)}")
    lines.append("- **Phase-by-Phase Breakdown**:")
    data_load_s = m.phase_seconds.get("data_loading")
    render_s = m.phase_seconds.get("rendering")
    lines.append(f"  - File I/O + data processing (`load_all()` — reading all multi-sheet .xlsx sources): {_fmt(data_load_s, 's')}")
    lines.append(f"  - Presentation rendering (chart/table/highlight writes + PPTX save, all regions): {_fmt(render_s, 's')}")
    lines.append("  - Text generation/summarization (dedicated phase): **N/A** — this pipeline does not call an LLM or "
                  "external summarization API. Highlight/insight text (`get_slideN_highlights()` in `chart_data.py`) is "
                  "produced by local deterministic string-formatting heuristics computed directly from the same data used "
                  "to fill charts/tables. That work happens inline during the rendering phase above and is not separately "
                  "timed because it is not a distinct, isolable cost (typically sub-millisecond per call, dominated by the "
                  "XML string operations it's embedded in).")
    lines.append(f"- **Throughput Velocity**:")
    lines.append(f"  - Charts populated / second (during rendering phase): {_fmt(m.charts_per_second, '/s', digits=2)}")
    lines.append(f"  - Slide sections populated / minute (across all regions): {_fmt(m.slides_per_minute, '/min', digits=1)}")
    if m.regions_total and m.total_seconds:
        lines.append(f"  - Regions (full 19-slide decks) / minute: {_fmt(m.regions_total / m.total_seconds * 60, '/min', digits=2)}")
    lines.append("- **Per-Chart Timing** (time spent inside `_update_chart`, wall-clock per call, summed across all regions):")
    lines.append(f"  - Total time across all chart-update calls: {_fmt(m.chart_total_seconds, 's', digits=2)}")
    lines.append(f"  - Average time per chart-update call: {_fmt(m.avg_chart_seconds, 's', digits=3)}")
    if m.chart_max_seconds_num is not None:
        lines.append(f"  - Slowest single chart-update call: chart{m.chart_max_seconds_num} took {_fmt(m.chart_max_seconds, 's', digits=3)}")
    if m.chart_seconds_by_num:
        lines.append("  - Slowest charts by total time (summed across all regions, since the same chart number runs once per region):")
        ranked = sorted(m.chart_seconds_by_num.items(), key=lambda kv: -kv[1])[:10]
        for chart_num, total_s in ranked:
            n = m.chart_count_by_num[chart_num]
            lines.append(f"    - chart{chart_num}: {total_s:.2f}s total across {n} call(s) ({total_s / n:.3f}s avg)")
    lines.append("")

    # -----------------------------------------------------------------
    lines.append("## 2. Data Integrity & Processing Scope")
    lines.append("")
    lines.append(f"- **Ingestion Volume**:")
    lines.append(f"  - Excel workbooks configured as sources: {m.workbooks_configured}")
    lines.append(f"  - Sheet reads performed: {m.sheet_reads:,} (a workbook may be reread per lookup — see `_sheet_rows` in `data_loader.py`; this is call count, not distinct-sheet count)")
    lines.append(f"  - Rows read (sum across all sheet reads): {m.rows_read:,}")
    lines.append(f"  - Cells read (estimated, rows x column-count of each read): {m.cells_read:,}")
    lines.append(f"  - Chart series mapping attempts (charts x regions): {m.chart_attempts:,}")
    lines.append("- **Extraction & Mapping Accuracy**:")
    lines.append(f"  - Chart series successfully mapped: {m.chart_successes:,} / {m.chart_attempts:,} "
                  f"({_fmt((m.chart_success_rate or 0) * 100, '%', digits=1)})")
    if m.chart_failure_reasons:
        lines.append("  - Chart failure reasons (from the existing try/except in the chart-update loop):")
        for reason, count in sorted(m.chart_failure_reasons.items(), key=lambda kv: -kv[1]):
            lines.append(f"    - {count}x: {reason}")
    lines.append(f"  - Slide/table sections successfully populated: {m.slide_successes:,} / {m.slide_attempts:,} "
                  f"({_fmt((m.slide_success_rate or 0) * 100, '%', digits=1)})")
    if m.slide_failure_reasons:
        lines.append("  - Slide failure reasons (from `_safe_apply_slide`'s `[WARN]` log):")
        for reason, count in sorted(m.slide_failure_reasons.items(), key=lambda kv: -kv[1]):
            lines.append(f"    - {count}x: {reason}")
    lines.append("- **Highlight Faithfulness**: **N/A in the LLM-hallucination sense** — there is no language model in "
                  "this pipeline that could invent a number. Every highlight/insight figure is produced by directly "
                  "formatting values already computed from the source Excel data (the same values written into the "
                  "adjacent chart/table on the same slide), so by construction the figures cannot diverge from the "
                  "underlying data. The actual risk class in this codebase is *derivation* bugs (wrong denominator, "
                  "wrong column, sign errors — several of which were found and fixed during development, e.g. the "
                  "slide 14 percentage-basis and slide 15 indicator-orientation fixes), not hallucination. This is "
                  "verified by code review and targeted spot-checks against the source template's own reference "
                  "numbers, not by an automated faithfulness score.")
    lines.append("")

    # -----------------------------------------------------------------
    lines.append("## 3. System Resource Utilization")
    lines.append("")
    peak_rss = _peak_rss_mb()
    lines.append(f"- **Peak RAM Consumption**: {_fmt(peak_rss, ' MB', digits=1)} "
                  f"(peak resident set size for this process for the entire run, via `resource.getrusage` — "
                  f"includes the full employee databases loaded via pandas in `_load_db_b`/`_load_active_frontliners`, "
                  f"which are the largest in-memory structures in this pipeline)")
    lines.append("- **External Dependency Overhead**: **N/A** — this pipeline makes no network/API calls of any kind "
                  "(no LLM endpoint, no external service). All processing is local file I/O (openpyxl/pandas) and "
                  "in-memory string/XML manipulation. There is no API latency, retry count, rate-limit, or token cost "
                  "to report.")
    lines.append("")

    # -----------------------------------------------------------------
    lines.append("## 4. Operational ROI & Error Rates")
    lines.append("")
    if m.regions_total and m.total_seconds:
        auto_min_per_region = (m.total_seconds / m.regions_total) / 60
    else:
        auto_min_per_region = None
    lines.append(f"- **Time Savings Equivalent** (manual baseline is an ASSUMPTION, not measured — see "
                  f"`--manual-baseline-min` in `main.py`):")
    lines.append(f"  - Automated: {_fmt(auto_min_per_region, ' min/region', digits=2)} "
                  f"({_fmt(m.total_seconds / m.regions_total if m.regions_total else None, 's/region', digits=1)})")
    lines.append(f"  - Manual baseline (assumed): {manual_baseline_minutes:.0f} min/region")
    if auto_min_per_region:
        speedup = manual_baseline_minutes / auto_min_per_region
        lines.append(f"  - Speedup: ~{speedup:,.0f}x faster than the assumed manual baseline")
    lines.append(f"- **Exception & Failure Frequency**:")
    lines.append(f"  - Regions failed outright: {m.regions_failed} / {m.regions_total}")
    if m.region_failure_reasons:
        for reason, count in sorted(m.region_failure_reasons.items(), key=lambda kv: -kv[1]):
            lines.append(f"    - {count}x: {reason}")
    lines.append(f"  - Chart-level failures (region completed, one chart skipped): {sum(m.chart_failure_reasons.values())}")
    lines.append(f"  - Slide-level failures (region completed, one slide section skipped): {sum(m.slide_failure_reasons.values())}")
    lines.append("- **Post-Run Intervention Rate**: **not automatable** — this requires a human to open each generated "
                  "deck and judge whether anything needs manual formatting/repositioning/text correction. " +
                  (post_run_intervention_note or "No review has been logged for this run. To track this over time, "
                   "record it after reviewing the output decks and pass it via `write_report(..., "
                   "post_run_intervention_note=...)`."))
    lines.append("")

    with open(path, "w") as f:
        f.write("\n".join(lines) + "\n")
    print(f"Performance report written to: {path}")
