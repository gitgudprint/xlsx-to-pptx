"""
Slide 6 & 7 — PA (Penilaian kinerja) Sales dan Collection. Kedua slide ini
sederhana (cuma highlight/insight otomatis, tidak ada tabel/shape yang
digeser-geser) jadi digabung dalam satu modul kecil.
"""
from ..pptx_common import _apply_highlight_tokens
from ..chart_data import get_slide6_highlights, get_slide7_highlights

def _apply_slide6_highlights(slide_xml, data, region):
    """
    Mengisi 3 placeholder highlight slide 6 ({SLIDE6_HIGHLIGHT_A/B/C} — PA
    Sales: LOS/usia/pendidikan) dengan teks draf hasil
    `get_slide6_highlights()` di chart_data.py, lewat `_apply_highlight_tokens`.

    Parameter: `slide_xml` (bytes XML slide 6), `data` (dict load_all()),
    `region`.

    Return: bytes XML slide 6 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 9b).
    """
    return _apply_highlight_tokens(slide_xml, "SLIDE6_HIGHLIGHT", get_slide6_highlights(data, region))


def _apply_slide7_highlights(slide_xml, data, region):
    """
    Sama seperti `_apply_slide6_highlights`, untuk slide 7 (PA Collection),
    memakai `get_slide7_highlights()`.

    Return: bytes XML slide 7 dengan token sudah diganti. Dipanggil dari
    `generate_pptx_for_region` (langkah 9b).
    """
    return _apply_highlight_tokens(slide_xml, "SLIDE7_HIGHLIGHT", get_slide7_highlights(data, region))


