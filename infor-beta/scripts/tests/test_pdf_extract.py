"""Unit tests for the shared PDF text -> OCR-fallback helper.

The garble heuristic is a pure string function, so it is tested directly. The
orchestrator's text / OCR / degrade-gracefully branches are tested by
monkeypatching the extraction + OCR steps, so the suite needs neither a real PDF,
``pypdfium2``, nor the ``tesseract`` binary (CI-safe per the task brief).
"""

from pathlib import Path

import pytest

import pdf_extract
from pdf_extract import (
    OcrUnavailableError,
    PdfText,
    extract_pdf_text,
    garble_ratio,
    is_garbled,
    ocr_available,
)

# A clean, English, number-heavy statements sentence (the common case).
_CLEAN = (
    "Net income for the fiscal year ended December 31, 2025 was $1,234.5 million, "
    "up 12.3% from $1,099.2 million in the prior year. Total assets were $8,765.4M."
)
# CID-scramble: glyph ids landing in CJK / private-use ranges (the failure mode).
_CID_SCRAMBLE = "中文 アイウ 高"
# Replacement-character soup (no ToUnicode map): U+FFFD repeated.
_REPLACEMENT = "��� ���� �����"


# --- garble heuristic --------------------------------------------------------
def test_clean_english_text_is_not_garbled():
    assert is_garbled(_CLEAN) is False
    assert garble_ratio(_CLEAN) == 0.0


def test_number_heavy_table_is_not_garbled():
    # Mostly digits / currency punctuation, few words -- still all plausible chars.
    table = "Revenue 1,234.5 1,099.2 987.6\nEBITDA 456.7 412.3 388.1\n$ % (12.3)"
    assert is_garbled(table) is False


def test_empty_and_whitespace_are_garbled():
    assert is_garbled("") is True
    assert is_garbled("   \n\t  ") is True
    assert is_garbled(None) is True
    assert garble_ratio("") == 1.0


def test_too_short_text_is_garbled():
    # Below the minimum-signal threshold -> treat as a failed extraction.
    assert is_garbled("Revenue") is True


def test_cid_scramble_is_garbled():
    assert is_garbled(_CID_SCRAMBLE) is True
    assert garble_ratio(_CID_SCRAMBLE) > 0.5


def test_replacement_char_soup_is_garbled():
    assert is_garbled(_REPLACEMENT) is True


def test_accented_latin_is_plausible():
    # Accented Latin (e.g. a Quebec issuer name) must NOT trip the heuristic.
    text = (
        "Société Générale reported résultats for the fiscal "
        "year — net income rose 8.4%."
    )
    assert is_garbled(text) is False


def test_threshold_is_tunable():
    # A few implausible chars in a long otherwise-clean string: passes at the
    # default cut but fails a strict (0.0) one.
    mixed = _CLEAN + _CID_SCRAMBLE
    assert is_garbled(mixed) is False
    assert is_garbled(mixed, threshold=0.0) is True


# --- orchestrator branches (monkeypatched) ----------------------------------
@pytest.fixture
def fake_pdf(tmp_path: Path) -> Path:
    p = tmp_path / "filing.pdf"
    p.write_bytes(b"%PDF-1.7 fake")  # existence only; extraction is monkeypatched
    return p


def test_extract_uses_text_layer_when_clean(fake_pdf, monkeypatch):
    monkeypatch.setattr(pdf_extract, "extract_text_pdfium", lambda p, **k: _CLEAN)
    # OCR must NOT be called on the clean path.
    monkeypatch.setattr(
        pdf_extract, "ocr_pdf", lambda *a, **k: pytest.fail("ocr_pdf should not run")
    )
    out = extract_pdf_text(fake_pdf)
    assert isinstance(out, PdfText)
    assert out.method == "text"
    assert out.garbled is False
    assert out.text == _CLEAN


def test_extract_falls_back_to_ocr_when_garbled(fake_pdf, monkeypatch):
    monkeypatch.setattr(pdf_extract, "extract_text_pdfium", lambda p, **k: _CID_SCRAMBLE)
    monkeypatch.setattr(
        pdf_extract, "ocr_pdf", lambda p, **k: "Clean OCR text from the rendered page."
    )
    out = extract_pdf_text(fake_pdf)
    assert out.method == "ocr"
    assert out.garbled is True
    assert out.text == "Clean OCR text from the rendered page."


def test_extract_degrades_when_ocr_unavailable(fake_pdf, monkeypatch):
    monkeypatch.setattr(pdf_extract, "extract_text_pdfium", lambda p, **k: _CID_SCRAMBLE)

    def _no_ocr(*a, **k):
        raise OcrUnavailableError("tesseract not installed")

    monkeypatch.setattr(pdf_extract, "ocr_pdf", _no_ocr)
    out = extract_pdf_text(fake_pdf)
    # Does not raise -- returns the garbled text layer, flagged.
    assert out.method == "text"
    assert out.garbled is True
    assert out.text == _CID_SCRAMBLE


def test_prefer_ocr_skips_text_layer(fake_pdf, monkeypatch):
    monkeypatch.setattr(
        pdf_extract,
        "extract_text_pdfium",
        lambda p, **k: pytest.fail("text layer should be skipped when prefer_ocr"),
    )
    monkeypatch.setattr(pdf_extract, "ocr_pdf", lambda p, **k: "OCR only")
    out = extract_pdf_text(fake_pdf, prefer_ocr=True)
    assert out.method == "ocr"
    assert out.text == "OCR only"


def test_missing_file_raises(tmp_path: Path):
    with pytest.raises(FileNotFoundError):
        extract_pdf_text(tmp_path / "nope.pdf")


def test_ocr_pdf_raises_when_binary_absent(fake_pdf, monkeypatch):
    monkeypatch.setattr(pdf_extract, "_tesseract_path", lambda: None)
    assert ocr_available() is False
    with pytest.raises(OcrUnavailableError):
        pdf_extract.ocr_pdf(fake_pdf)
