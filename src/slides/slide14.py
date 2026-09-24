"""
Slide 14 — Attrition YoY (3 tabel ranking NR/RG/Total + highlight top-3
per kolom) dan tabel "Top 3 Reason Out Attrition" per region, plus
highlight/insight otomatisnya.
"""
import re

from ..config import REGION_ABBREV, REGIONS
from ..xml_updater import update_table_rows
from ..pptx_common import _rebuild_row_cells, _apply_highlight_tokens, _TCPR_TAIL_OPTIONAL_RE
from ..chart_data import get_slide14_highlights

_SLIDE14_METRIC_KEYS = [
    ("yoy_non_regret_25", "yoy_non_regret_26"),  # tabel 0: NR
    ("yoy_regret_25",     "yoy_regret_26"),      # tabel 1: RG
    ("yoy_total_25",      "yoy_total_26"),       # tabel 2: Total
]


def _fmt_pct_id(v):
    """
    Format persentase gaya Indonesia 2 desimal (koma sebagai pemisah
    desimal), mis. "12,34%". Menerima nilai fraksi (0.1234) ATAU sudah
    dalam skala persen, dibedakan lewat ambang `abs(v) <= 1.5`;
    None → "0,00%".
    """
    if v is None:
        return "0,00%"
    try:
        pct = float(v) * 100 if abs(float(v)) <= 1.5 else float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{pct:.2f}%".replace(".", ",")


def _slide14_label_lookup():
    """
    Membangun peta label→region-kanonis untuk slide 14: tiap region
    (`REGIONS`, dari `config.py`) memetakan ke dirinya sendiri, ditambah
    setiap singkatan region (`REGION_ABBREV`) memetakan balik ke nama
    kanonisnya, plus dua label khusus non-region ("Head Office", "Nasional")
    yang memetakan ke dirinya sendiri. Dipanggil sekali saat modul di-import
    untuk mengisi konstanta `_SLIDE14_LABEL_LOOKUP` di bawah, dipakai oleh
    `_resolve_slide14_table_rows` untuk mengenali label baris tabel apa pun
    bentuknya (nama penuh atau singkatan).
    """
    lookup = {r: r for r in REGIONS}
    lookup.update({abbrev: canon for canon, abbrev in REGION_ABBREV.items()})
    lookup["Head Office"] = "Head Office"
    lookup["Nasional"] = "Nasional"
    return lookup


_SLIDE14_LABEL_LOOKUP = _slide14_label_lookup()


def _resolve_slide14_table_rows(table_xml):
    """
    Mengenali region mana yang menjadi pemilik tiap baris pada satu tabel
    ranking di slide 14 (satu dari 3 tabel: NR/RG/Total).

    MASALAH yang dipecahkan: ketiga tabel di template masing-masing punya
    urutan region sendiri-sendiri yang independen (sudah terurut berdasarkan
    nilainya masing-masing — lihat `load_attrition` di `data_loader.py`),
    dan satu baris labelnya bisa jadi MASIH berupa placeholder literal
    template "{REGION}"/"REGION" (bukan nama/singkatan region yang
    sebenarnya). Karena setiap tabel punya urutan berbeda, tidak bisa
    ditebak begitu saja placeholder itu "pasti region X" hanya dari posisi
    barisnya.

    CARA MENYELESAIKAN — resolusi lewat ELIMINASI: setiap baris yang
    labelnya SUDAH berupa nama/singkatan region yang dikenali
    (`_SLIDE14_LABEL_LOOKUP`) dicatat sebagai `seen`. Setelah semua baris
    diperiksa, region mana pun di `REGIONS` yang TIDAK PERNAH muncul di
    `seen` untuk tabel ini pastilah region yang bersembunyi di balik
    placeholder "REGION" tersebut. Ini hanya bisa disimpulkan dengan pasti
    kalau tepat ada SATU baris placeholder dan SATU region yang hilang
    (`len(placeholder_idx) == 1 and len(missing) == 1`); kalau situasinya
    ambigu (lebih dari satu placeholder atau lebih dari satu kandidat),
    baris itu sengaja dibiarkan tidak teridentifikasi (`None`) daripada
    menebak secara serampangan (menebak salah akan menimpa/menduplikasi
    baris region yang salah).

    Parameter:
        table_xml: string XML satu `<a:tbl>...</a:tbl>` slide 14 (salah
                   satu dari 3 tabel NR/RG/Total, atau — dari pemanggil
                   `_apply_slide14_rect_highlights` — potongan
                   `<p:graphicFrame>` yang membungkus tabel tsb, karena
                   regex `<a:tr\\b.*?</a:tr>` di dalam fungsi ini tetap
                   cocok pada kedua bentuk itu).

    Return: tuple `(row_blocks, resolved)` — `row_blocks` adalah list
    string XML tiap `<a:tr>` (dipakai ulang oleh pemanggil untuk mencari
    posisi baris berdasarkan indeks yang sama), dan `resolved[i]` adalah
    key region kanonis untuk baris ke-i (atau `None` untuk baris yang
    bukan baris data, mis. header, atau yang tidak bisa diselesaikan).
    Dipanggil dari `_apply_slide14_tables` (untuk tahu baris mana perlu
    diisi nilai apa) dan `_apply_slide14_rect_highlights` (untuk tahu
    posisi vertikal baris `region` demi memindahkan kotak highlight).
    """
    row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    resolved = []
    seen = set()
    placeholder_idx = []

    for idx, rb in enumerate(row_blocks):
        texts = re.findall(r'<a:t>([^<]*)</a:t>', rb)
        if len(texts) < 3:
            resolved.append(None)
            continue
        n = len(texts) - 2   # 2 sel terakhir selalu nilai YTD25/YTD26
        label = ''.join(texts[:n])
        canon = _SLIDE14_LABEL_LOOKUP.get(label)
        if canon is not None:
            seen.add(canon)
            resolved.append(canon)
        elif label == "REGION":
            resolved.append("__PLACEHOLDER__")
            placeholder_idx.append(idx)
        else:
            resolved.append(None)

    missing = [r for r in REGIONS if r not in seen]
    if len(placeholder_idx) == 1 and len(missing) == 1:
        resolved[placeholder_idx[0]] = missing[0]
    else:
        for idx in placeholder_idx:
            resolved[idx] = None   # ambigu — biarkan tak terselesaikan daripada menebak

    return row_blocks, resolved


def _apply_slide14_tables(slide_xml, data):
    """
    Mengisi 3 tabel YoY attrition NR / RG / Total di slide 14 — yaitu 3
    blok `<a:tbl>` PERTAMA pada slide (blok ke-4 adalah tabel "Top 3 Reason
    Out", ditangani terpisah oleh `_apply_slide14_reason_table` dan tidak
    disentuh di sini).

    Cara kerja: XML slide dipecah jadi potongan-potongan lewat
    `re.split(r'(<a:tbl>.*?</a:tbl>)', ...)` (grup tertangkap disertakan,
    jadi setiap tabel jadi elemen terpisah dalam `parts`, diselingi teks
    non-tabel). Untuk 3 tabel pertama, `_resolve_slide14_table_rows(part)`
    menentukan region kanonis tiap barisnya, lalu tabel itu diisi lewat
    `update_table_rows()` (`xml_updater.py`) dengan:
      - `row_key_fn`: memakai `counter` closure untuk mengambil elemen
        `resolved[idx]` yang cocok secara BERURUTAN dengan baris yang
        sedang diproses `update_table_rows` (karena `update_table_rows`
        memproses baris satu per satu dengan `re.sub`, urutan pemanggilan
        `row_key_fn` PASTI sama dengan urutan baris fisik di tabel —
        sehingga `counter` yang naik 1 tiap panggilan tetap sinkron dengan
        indeks `resolved`).
      - `row_values_fn`: mengambil baris `attrition[canon]` (dari
        `data["attrition"]`, hasil `load_attrition()` di
        `data_loader.py`), memformat nilai YTD25/YTD26 dengan `_fmt_pct_id`
        memakai key metrik yang sesuai tabel ini (`_SLIDE14_METRIC_KEYS[table_idx]`
        — NR untuk tabel 0, RG untuk tabel 1, Total untuk tabel 2). Kalau
        label baris masih literal placeholder "REGION" (kasus yang
        diresolusi lewat eliminasi tadi), label itu DIGANTI dengan singkatan
        region-nya (`REGION_ABBREV`) — kalau bukan placeholder, label yang
        sudah ada (nama/singkatan region lain) dibiarkan (`None` di posisi
        itu berarti "tidak diubah").

    Parameter: `slide_xml` (bytes XML slide 14, sebelum diisi), `data`
    (dict `load_all()`, dipakai key `"attrition"`).

    Return: bytes XML slide 14 dengan 3 tabel YoY sudah terisi. Dipanggil
    dari `generate_pptx_for_region` (langkah 3), sebelum
    `_apply_slide14_rect_highlights` dan `_apply_slide14_reason_table`
    dijalankan berantai di atas hasilnya.
    """
    attrition = data.get("attrition", {})
    text = slide_xml.decode('utf-8')
    parts = re.split(r'(<a:tbl>.*?</a:tbl>)', text, flags=re.DOTALL)

    table_idx = 0
    for i, part in enumerate(parts):
        if not part.startswith('<a:tbl>'):
            continue
        if table_idx < 3:
            _, resolved = _resolve_slide14_table_rows(part)
            v25_key, v26_key = _SLIDE14_METRIC_KEYS[table_idx]
            counter = [0]

            def row_key_fn(cur_texts, resolved=resolved, counter=counter):
                idx = counter[0]
                counter[0] += 1
                return resolved[idx] if idx < len(resolved) else None

            def row_values_fn(canon, cur_texts, v25_key=v25_key, v26_key=v26_key):
                row = attrition.get(canon)
                if row is None:
                    return None
                n = len(cur_texts) - 2
                label = ''.join(cur_texts[:n])
                v25 = _fmt_pct_id(row.get(v25_key))
                v26 = _fmt_pct_id(row.get(v26_key))
                if label == "REGION":
                    override = REGION_ABBREV.get(canon, canon)
                    labels = [override] + [None] * (n - 1)
                else:
                    labels = [None] * n
                return labels + [v25, v26]

            updated = update_table_rows(part.encode('utf-8'), row_key_fn, row_values_fn)
            parts[i] = updated.decode('utf-8')
        table_idx += 1

    return ''.join(parts).encode('utf-8')


# Fill "pink" (sebenarnya tint pucat dari warna tema accent2, BUKAN kode
# RGB harfiah) yang dipakai untuk highlight top-3 di 3 tabel NR/RG/Total
# slide 14 — persis meniru fill yang dipakai contoh nyata di
# "template_bg.pptx" (baris top-3 Jawa Tengah kolom YTD May26 pada tabel
# NR sudah memakai fill ini di file itu). Dengan `schemeClr` (bukan
# `srgbClr` hardcode), warnanya otomatis ikut tema kalau tema diganti,
# sama seperti cara templatenya sendiri mendefinisikan fill ini.
_SLIDE14_PINK_FILL = '<a:solidFill><a:schemeClr val="accent2"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>'


def _apply_slide14_top3_highlights(slide_xml, data):
    """
    Menghitung dan menerapkan fill pink (`_SLIDE14_PINK_FILL`, tint pucat
    accent2 — persis fill yang sudah dipakai di contoh nyata
    "template_bg.pptx" untuk baris top-3) pada 3 nilai TERBESAR per kolom
    (YTD May25 dan YTD May26, masing-masing dihitung TERPISAH) di tiap
    satu dari 3 tabel ranking NR/RG/Total slide 14 — dipanggil SETELAH
    `_apply_slide14_tables` mengisi teks tabelnya.

    Cara kerja: sama seperti `_apply_slide14_tables`, XML dipecah lewat
    `re.split` supaya 3 tabel pertama bisa diproses satu-satu, lalu tiap
    tabel diproses lewat `_resolve_slide14_table_rows` untuk tahu region
    kanonis pemilik tiap baris. Baris "Head Office"/"Nasional" (bukan
    salah satu dari 12 `REGIONS`) SENGAJA DIKECUALIKAN dari perankingan —
    "3 terbesar" di sini berarti 3 REGION terbesar, bukan ikut
    dibandingkan dengan baris agregat perusahaan/nasional.

    Nilai mentah (float, BUKAN string hasil format) diambil langsung dari
    `data["attrition"][region][v25_key]`/`[v26_key]` (kunci metrik sesuai
    `_SLIDE14_METRIC_KEYS[table_idx]`, sama seperti yang dipakai
    `_apply_slide14_tables` untuk mengisi teksnya) supaya ranking akurat
    tanpa terpengaruh pembulatan tampilan.

    Sel yang di-highlight: kolom YTD25 (index sel ke-1) untuk anggota
    top-3 `y25`, kolom YTD26 (index sel ke-2) untuk anggota top-3 `y26` —
    KEDUA kolom dihitung independen (satu region bisa masuk top-3 salah
    satu kolom saja, top-3 keduanya, atau tidak masuk keduanya). SEMUA
    baris data (baris 1 dst., HANYA baris header index 0 yang dilewati)
    diproses, TERMASUK baris "Head Office"/"Nasional" yang tidak ikut
    diranking (bukan salah satu dari 12 `REGIONS`) — baris-baris ini
    PASTI di-set ke fill putih polos (`schemeClr val="bg1"`, warna default
    baris data di template ini) karena tidak pernah bisa masuk himpunan
    top3. Ini PENTING: kalau baris seperti itu justru DILEWATI (tidak
    diproses sama sekali), fill highlight pink STATIS yang sudah ada di
    template dari contoh demo lamanya (mis. baris "Head Office" pada
    tabel RG) akan tetap nyangkut apa adanya, membuat tampilannya seolah
    ada 4 baris top-3 padahal seharusnya cuma 3.

    Parameter: `slide_xml` (bytes XML slide 14, hasil `_apply_slide14_tables`),
    `data` (dict `load_all()`, key `"attrition"`).

    Return: bytes XML slide 14 dengan highlight top-3 sudah diterapkan.
    Dipanggil dari `generate_pptx_for_region` (langkah 3), tepat setelah
    `_apply_slide14_tables`.
    """
    attrition = data.get("attrition", {})
    text = slide_xml.decode('utf-8')
    parts = re.split(r'(<a:tbl>.*?</a:tbl>)', text, flags=re.DOTALL)

    table_idx = 0
    for i, part in enumerate(parts):
        if not part.startswith('<a:tbl>'):
            continue
        if table_idx >= 3:
            table_idx += 1
            continue
        row_blocks, resolved = _resolve_slide14_table_rows(part)
        v25_key, v26_key = _SLIDE14_METRIC_KEYS[table_idx]

        candidates = []
        for idx, canon in enumerate(resolved):
            if canon not in REGIONS:
                continue
            row = attrition.get(canon)
            if row is None:
                continue
            candidates.append((idx, row.get(v25_key), row.get(v26_key)))

        top3_y25 = {idx for idx, _, _ in
                    sorted((c for c in candidates if c[1] is not None), key=lambda c: -c[1])[:3]}
        top3_y26 = {idx for idx, _, _ in
                    sorted((c for c in candidates if c[2] is not None), key=lambda c: -c[2])[:3]}

        new_part = part
        for idx, row_xml in enumerate(row_blocks):
            if idx == 0:
                continue   # baris header, tidak pernah diisi/di-highlight

            def transform(ci, tc, idx=idx):
                if ci == 1:
                    nyala = idx in top3_y25
                elif ci == 2:
                    nyala = idx in top3_y26
                else:
                    return tc
                fill = _SLIDE14_PINK_FILL if nyala else '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>'
                return _TCPR_TAIL_OPTIONAL_RE.sub(rf'\1{fill}\2', tc, count=1)

            new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
            if num_tcs == 3:
                new_part = new_part.replace(row_xml, new_row_xml, 1)

        parts[i] = new_part
        table_idx += 1

    return ''.join(parts).encode('utf-8')


def _apply_slide14_rect_highlights(slide_xml, region):
    """
    Memindahkan masing-masing dari 3 kotak highlight merah slide 14 ke
    baris yang cocok dengan `region` di TABEL yang sesuai — dengan teknik
    posisi-baris yang sama seperti slide 10 (tinggi kumulatif baris), tapi
    di sini ada langkah tambahan: mencocokkan tiap kotak highlight dengan
    TABEL mana yang menjadi induknya, karena ada 3 tabel berdampingan.

    Cara kerja:
      1. Untuk tiap `<p:graphicFrame>` yang memuat `<a:tbl>` (3 tabel
         pertama saja — pengecekan `table_idx >= 3` melewati tabel reason-
         out ke-4), catat offset-x tabel itu (`off_m`, posisi horizontal
         tabel di slide), region tiap barisnya (lewat
         `_resolve_slide14_table_rows(gf)`), dan tinggi kumulatif tiap
         baris (`cum`, dihitung sama seperti slide 10). Semua ini disimpan
         di `table_cols` sebagai `(off_x, keys, cum)`.
      2. Untuk tiap shape `<p:sp>` yang mengandung warna merah "FF0000"
         (kotak highlight), tabel induknya ditentukan dengan mencari entri
         `table_cols` yang offset-x-nya PALING DEKAT dengan offset-x kotak
         itu sendiri (`min(..., key=lambda t: abs(t[0] - rect_x))`) — ini
         cara mengaitkan kotak dengan tabelnya secara struktural/posisional,
         bukan dengan asumsi urutan tabel = urutan kotak.
      3. Setelah tabel induknya ketemu, delta posisi y dihitung sama
         seperti slide 10: selisih tinggi kumulatif antara baris `region`
         dan baris "Jawa Tengah" (baris acuan template) di tabel itu.

    Parameter: `slide_xml` (bytes XML slide 14, hasil dari
    `_apply_slide14_tables` — sehingga baris-baris sudah berisi data
    terbaru saat posisi barisnya dibaca ulang di sini), `region`.

    Return: bytes XML slide 14 dengan posisi ketiga kotak highlight sudah
    diperbarui. Dipanggil dari `generate_pptx_for_region` (langkah 3),
    hasilnya diteruskan ke `_apply_slide14_reason_table`.
    """
    text = slide_xml.decode('utf-8')

    table_cols = []  # (off_x, keys, tinggi_kumulatif_awal_tiap_baris)
    gfs = re.findall(r'<p:graphicFrame>.*?</p:graphicFrame>', text, re.DOTALL)
    table_idx = 0
    for gf in gfs:
        if '<a:tbl>' not in gf:
            continue
        if table_idx >= 3:
            table_idx += 1
            continue
        off_m = re.search(r'<a:off x="(\d+)" y="\d+"/>', gf)
        if not off_m:
            table_idx += 1
            continue
        _, keys = _resolve_slide14_table_rows(gf)
        heights = [int(h) for h in re.findall(r'<a:tr h="(\d+)"', gf)]
        cum = [0]
        for h in heights:
            cum.append(cum[-1] + h)
        table_cols.append((int(off_m.group(1)), keys, cum))
        table_idx += 1

    if not table_cols:
        return slide_xml

    sp_blocks = re.findall(r'<p:sp\b.*?</p:sp>', text, re.DOTALL)
    for sp in sp_blocks:
        if 'FF0000' not in sp:
            continue
        off_m = re.search(r'<a:off x="(\d+)" y="(\d+)"', sp)
        if not off_m:
            continue
        rect_x, old_y = int(off_m.group(1)), int(off_m.group(2))
        _, keys, cum = min(table_cols, key=lambda t: abs(t[0] - rect_x))
        try:
            target_idx = keys.index(region)
        except ValueError:
            continue
        anchor_idx = keys.index("Jawa Tengah") if "Jawa Tengah" in keys else target_idx
        delta = cum[target_idx] - cum[anchor_idx]
        if delta == 0:
            continue
        new_y = old_y + delta
        new_sp = sp.replace(f'y="{old_y}"', f'y="{new_y}"', 1)
        text = text.replace(sp, new_sp, 1)

    return text.encode('utf-8')


_TCPR_TAIL_RE = re.compile(r'(</a:lnB>)(?:<a:noFill/>|<a:solidFill>.*?</a:solidFill>)(</a:tcPr>)', re.DOTALL)


def _fmt_reason_pct(v):
    """Format persentase 2 desimal gaya Indonesia dari nilai fraksi (0.1234 → "12,34%")."""
    return f"{v * 100:.2f}%".replace(".", ",")


def _apply_slide14_reason_table(slide_xml, data, region):
    """
    Mengisi tabel "Top 3 Reason Out Attrition" di slide 14 (blok `<a:tbl>`
    KE-4, setelah 3 tabel NR/RG/Total yang ditangani `_apply_slide14_tables`)
    dengan nilai `region` yang sedang diproses — diambil dari
    `data["attrition"][region]["reason_out_rows"]`, hasil komputasi
    `load_slide14_reason_out(region)` di `data_loader.py` (jadi PER REGION,
    bukan nilai yang sama untuk semua region).

    Selain mengisi nilai, fungsi ini juga menghitung highlight (fill kuning
    "FFFF00") pada baris alasan (dari 6 baris individual, bukan baris
    header/agregat — `leaf_idxs`) yang punya nilai TERTINGGI di
    masing-masing kolom nilai — meniru highlight contoh statis yang sudah
    ada di template (warna sama), tapi di sini dihitung secara dinamis per
    region lewat `max_idx`. Kolom non_regret dan regret dicari lintas
    SEMUA baris alasan (Involuntary + Voluntary sekaligus), TAPI kolom
    total (GRAND TOTAL) khusus HANYA mencari di antara baris alasan
    VOLUNTARY saja (`voluntary_leaf_idxs`, baris setelah header
    "Voluntary") — highlight kolom itu menandai alasan voluntary
    ter-signifikan, bukan alasan terbesar apa pun lintas kategori.

    Cara kerja detail: tabel ke-4 diambil lewat `tbl_matches[3]`. Validasi
    struktural `len(row_blocks) != len(reason_rows) + 1` memastikan jumlah
    baris tabel di template cocok persis dengan jumlah baris data
    (+1 header) sebelum mengubah apa pun — kalau tidak cocok, tabel
    dibiarkan apa adanya. Untuk tiap baris, `_rebuild_row_cells` dipakai
    (bukan search-and-replace biasa) untuk mengisi ulang label + 3 nilai
    kolom, dan sekaligus menyisipkan/menghapus fill kuning pada `<a:tcPr>`
    sel yang bersangkutan lewat `_TCPR_TAIL_RE` (regex yang menyasar bagian
    ekor `<a:tcPr>`, tepat setelah `</a:lnB>`, tempat elemen fill berada).

    Parameter: `slide_xml` (bytes XML slide 14, hasil rantai
    `_apply_slide14_tables` → `_apply_slide14_rect_highlights`), `data`
    (dict `load_all()`), `region`.

    Return: bytes XML slide 14 final. Dipanggil dari
    `generate_pptx_for_region` (langkah 3), hasilnya langsung disimpan
    lewat `editor.update(slide14_path, slide14_xml)`.
    """
    reason_rows = data.get("attrition", {}).get(region, {}).get("reason_out_rows", [])
    if not reason_rows:
        return slide_xml

    text = slide_xml.decode('utf-8')
    tbl_matches = list(re.finditer(r'<a:tbl>.*?</a:tbl>', text, re.DOTALL))
    if len(tbl_matches) < 4:
        return slide_xml
    table_xml = tbl_matches[3].group(0)

    row_blocks = re.findall(r'<a:tr\b.*?</a:tr>', table_xml, re.DOTALL)
    if len(row_blocks) != len(reason_rows) + 1:   # +1 untuk baris header
        return slide_xml

    leaf_idxs = [i for i, r in enumerate(reason_rows) if not r["is_header"]]
    # Kolom "total" (GRAND TOTAL) HANYA di-highlight pada baris alasan
    # VOLUNTARY dengan nilai terbesar (bukan alasan terbesar lintas
    # Involuntary+Voluntary seperti kolom non_regret/regret) — dicari lewat
    # baris header "Voluntary" di `reason_rows`, semua leaf SETELAHNYA
    # (sampai baris Grand Total, yang bukan leaf) adalah alasan voluntary.
    vol_header_idx = next((i for i, r in enumerate(reason_rows) if r["label"] == "Voluntary"), None)
    voluntary_leaf_idxs = [i for i in leaf_idxs if vol_header_idx is not None and i > vol_header_idx]
    max_idx = {col: max(leaf_idxs, key=lambda i: reason_rows[i][col])
               for col in ("non_regret", "regret")}
    max_idx["total"] = max(voluntary_leaf_idxs, key=lambda i: reason_rows[i]["total"]) if voluntary_leaf_idxs else None
    value_cols = [None, "non_regret", "regret", "total"]

    new_table_xml = table_xml
    for ridx, data_row in enumerate(reason_rows):
        row_xml = row_blocks[ridx + 1]

        new_values = [data_row["label"],
                      _fmt_reason_pct(data_row["non_regret"]),
                      _fmt_reason_pct(data_row["regret"]),
                      _fmt_reason_pct(data_row["total"])]

        def transform(ci, tc, new_values=new_values, data_row=data_row, ridx=ridx):
            new_tc = re.sub(r'(<a:t>)[^<]*(</a:t>)',
                             lambda m, v=new_values[ci]: m.group(1) + v.replace('&', '&amp;') + m.group(2),
                             tc, count=1)
            col = value_cols[ci]
            if col is not None and not data_row["is_header"]:
                fill = '<a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>' if max_idx[col] == ridx else '<a:noFill/>'
                new_tc = _TCPR_TAIL_RE.sub(rf'\1{fill}\2', new_tc, count=1)
            return new_tc

        new_row_xml, num_tcs = _rebuild_row_cells(row_xml, transform)
        if num_tcs != 4:
            continue
        new_table_xml = new_table_xml.replace(row_xml, new_row_xml, 1)

    text = text.replace(table_xml, new_table_xml, 1)
    return text.encode('utf-8')


def _apply_slide14_highlights(slide_xml, data, region):
    """
    Mengisi 4 placeholder highlight slide 14 ({SLIDE14_HIGHLIGHT_A..D} —
    nama region, tren YoY vs nasional, top reason involuntary non-regret,
    top reason voluntary/voluntary-regret) dengan teks draf hasil
    `get_slide14_highlights()` di chart_data.py, lewat
    `_apply_highlight_tokens`.

    Parameter: `slide_xml` (bytes XML slide 14 -- dipanggil setelah
    `_apply_slide14_reason_table`, karena slot C/D di sini dihitung dari
    `reason_out_rows` yang sama, meski secara teknis independen dari isi
    tabel XML-nya), `data` (dict load_all()), `region`.

    Return: bytes XML slide 14 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 3).
    """
    return _apply_highlight_tokens(slide_xml, "SLIDE14_HIGHLIGHT", get_slide14_highlights(data, region))


# ---------------------------------------------------------------------------
# Slide 15 – Attrition Report by Function (2 tabel + indikator naik/turun/tetap)
# ---------------------------------------------------------------------------
