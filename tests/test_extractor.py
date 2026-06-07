import pytest

from doc_extractor.extractor import build_schedule_rows, extract_text


def test_build_schedule_rows_detects_cable_tags():
    text = """
    MBSRE SYSTEM DRAWING
    Cable TAG CBL-001 from MCC-1 to VFD-1, size 4Cx2.5, Type XLPE
    Cable TAG CBL-002 from MCC-1 to PUMP-1, size 4Cx4.0, Type PVC
    """

    rows = build_schedule_rows(text, "Prepare a cable schedule from the MBSRE system drawing")

    assert len(rows) >= 2
    assert rows[0]['Cable tag'] == 'CBL-001'
    assert rows[0]['From Equipment'] == 'MCC-1'
    assert rows[0]['To Equipment'] == 'VFD-1'


def test_extract_text_reads_plain_text_input():
    text = "MBSRE drawing\nCable TAG CBL-010"

    extracted = extract_text(text, source_type='text')

    assert 'MBSRE drawing' in extracted
    assert 'CBL-010' in extracted


def test_extract_text_reports_missing_poppler(monkeypatch, tmp_path):
    import doc_extractor.extractor as extractor

    bad_pdf = tmp_path / 'scanned.pdf'
    bad_pdf.write_bytes(b'%PDF-1.4\ninvalid')

    monkeypatch.setattr(extractor, 'pytesseract', object())

    def raise_poppler(path):
        from pdf2image.exceptions import PDFInfoNotInstalledError

        raise PDFInfoNotInstalledError('poppler missing')

    monkeypatch.setattr(extractor, 'convert_from_path', raise_poppler)

    with pytest.raises(RuntimeError, match='poppler|Poppler'):
        extractor.extract_text(bad_pdf, source_type='pdf')
