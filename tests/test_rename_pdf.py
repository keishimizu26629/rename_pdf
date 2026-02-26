import datetime
import os
from pathlib import Path

import PyPDF2
from reportlab.pdfgen import canvas

# renamePdf から残っているユーティリティのテスト用
import renamePdf
from models import ProcessingResult, SheetData
from processors.base import calculate_confirm_day, trim_end_of_word
from services import merge_files_for_posting

# Initialize fonts for testing
renamePdf.register_fonts()


def create_sample_pdf(path: Path, text: str = "sample") -> None:
    """Create a simple single-page PDF for test purposes."""
    path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(path))
    c.setFont("Helvetica", 12)
    c.drawString(50, 750, text)
    c.showPage()
    c.save()


def test_trim_end_of_word_removes_trailing_spaces():
    assert trim_end_of_word("value   ") == "value"


def test_calculate_confirm_day_skips_weekend_and_holiday():
    holidays = ["2024/01/02"]
    shipping_day = datetime.date(2024, 1, 5)  # Friday
    expected = datetime.date(2023, 12, 29)  # Four business days before shipping_day

    assert calculate_confirm_day(shipping_day, holidays) == expected


def test_change_words_applies_transform():
    words = ["one", "two", "three"]
    renamePdf.change_words(words, lambda s: s.upper())
    assert words == ["ONE", "TWO", "THREE"]


def test_merge_files_for_posting_creates_prefixed_pdf(tmp_path: Path):
    target_dir = tmp_path / "pdfs"
    target_dir.mkdir()

    final_check_name = "【01-05】12345 67890 sample.pdf"
    companion_name = "12345 67890 sample.pdf"

    final_check_path = target_dir / final_check_name
    companion_path = target_dir / companion_name
    create_sample_pdf(final_check_path, text="final")
    create_sample_pdf(companion_path, text="companion")

    result = ProcessingResult(
        title_name="ユニットバスルーム納期最終確認票",
        file_name=str(final_check_path),
    )
    result.new_file_name = str(final_check_path)
    result.sheet_data_list.append(SheetData())

    output_dir = str(target_dir) + os.sep
    merge_files_for_posting([result], output_dir)

    merged_path = target_dir / f"(投函用){final_check_name}"
    assert merged_path.exists()

    with merged_path.open("rb") as merged_file:
        reader = PyPDF2.PdfReader(merged_file)
        assert len(reader.pages) == 2
