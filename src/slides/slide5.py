"""
Slide 5 — Demografi Frontliners (LOS/usia/pendidikan, Field Sales & Field
Coll): 24 kotak highlight persentase dan highlight/insight otomatisnya.
"""
import re

from ..chart_data import get_slide5_pct_boxes, get_slide5_highlights

def _apply_slide5_pct_boxes(slide_xml, data, region):
    """
    Mengisi 24 kotak highlight persentase pada slide 5 (LOS/AGE/EDU ×
    Regional/Nasional × Field Sales/Field Coll) dengan angka sungguhan,
    hasil `get_slide5_pct_boxes()` di chart_data.py — sebelumnya kotak-
    kotak ini tidak pernah disentuh sama sekali oleh kode manapun, jadi
    selalu menampilkan angka contoh statis dari template.

    Cara kerja: tiap shape dicari lewat nama persisnya (mis. "Rectangle
    64"), BUKAN lewat posisi/isi teks — karena nama shape adalah satu-
    satunya penanda yang stabil (isi teksnya sendiri, mis. "29%", bisa
    kebetulan sama antar kotak, dan posisi x/y tidak cocok untuk regex
    sederhana). Beberapa kotak menyimpan angkanya dalam DUA run terpisah
    (mis. `<a:t>64</a:t>` lalu `<a:t>%</a:t>`, bukan satu run "64%") — jadi
    nilai baru dituliskan ke run `<a:t>` PERTAMA, dan seluruh run
    berikutnya dalam shape yang sama dikosongkan (bukan dibiarkan), supaya
    tidak muncul sisa teks ganda seperti "64%%".

    Parameter: `slide_xml` (bytes XML slide 5), `data` (dict `load_all()`),
    `region` (str).

    Return: bytes XML slide 5 dengan seluruh kotak %-nya sudah diisi.
    Dipanggil dari `generate_pptx_for_region` pada langkah slide 5.
    """
    boxes = get_slide5_pct_boxes(data, region)
    text = slide_xml.decode('utf-8')
    sp_blocks = re.findall(r'<p:sp\b.*?</p:sp>', text, re.DOTALL)
    for sp in sp_blocks:
        m = re.search(r'name="([^"]*)"', sp)
        if not m or m.group(1) not in boxes:
            continue
        value = boxes[m.group(1)]
        counter = [0]

        def repl(mm, value=value, counter=counter):
            i = counter[0]
            counter[0] += 1
            return mm.group(1) + (value if i == 0 else '') + mm.group(2)

        new_sp = re.sub(r'(<a:t>)[^<]*(</a:t>)', repl, sp)
        if new_sp != sp:
            text = text.replace(sp, new_sp, 1)
    return text.encode('utf-8')


def _apply_slide5_highlights(slide_xml, data, region):
    """
    Mengisi placeholder highlight slide 5 dengan teks draf hasil
    `get_slide5_highlights()` di chart_data.py.

    STRUKTUR TOKEN KHUSUS (lihat catatan lengkap di
    `get_slide5_highlights`): token `{SLIDE5_HIGHLIGHT A}` (SPASI, bukan
    underscore) muncul 4 KALI berurutan di XML untuk 4 bullet berbeda,
    jadi diisi satu-per-satu SESUAI URUTAN KEMUNCULAN dari
    `highlights["A_list"]` (elemen ke-0 mengisi kemunculan pertama, dst —
    `text.replace(..., 1)` dipanggil berulang di dalam loop supaya tiap
    panggilan hanya "memakan" satu kemunculan paling awal yang tersisa).
    Token `{SLIDE5_HIGHLIGHT B}` muncul cuma sekali di kotak lain, diisi
    terpisah dari `highlights["B"]` (saat ini selalu "" — lihat alasannya
    di `get_slide5_highlights`).

    Parameter: `slide_xml` (bytes XML slide 5 — dipanggil setelah
    `_apply_slide5_pct_boxes`), `data` (dict load_all()), `region`.

    Return: bytes XML slide 5 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 3b).
    """
    def escape(s):
        return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    highlights = get_slide5_highlights(data, region)
    text = slide_xml.decode('utf-8')

    token_a = '<a:t>{SLIDE5_HIGHLIGHT A}</a:t>'
    a_values = list(highlights["A_list"])
    counter = [0]

    def repl_a(_m):
        i = counter[0]
        counter[0] += 1
        if i < len(a_values) and a_values[i]:
            return f'<a:t>{escape(a_values[i])}</a:t>'
        return _m.group(0)

    text = re.sub(re.escape(token_a), repl_a, text)

    if highlights["B"]:
        text = text.replace('<a:t>{SLIDE5_HIGHLIGHT B}</a:t>', f'<a:t>{escape(highlights["B"])}</a:t>', 1)

    return text.encode('utf-8')


# ---------------------------------------------------------------------------
# Anotasi teks per-slide
# ---------------------------------------------------------------------------
