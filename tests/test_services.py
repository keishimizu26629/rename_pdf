"""services モジュールのテスト."""

from unittest.mock import MagicMock

from models import ProcessingResult, SheetData, TextElement
from services import generate_rename_string, process_single_pdf


class TestProcessSinglePdf:
    def test_returns_none_for_unknown_document(self):
        extractor = MagicMock()
        # 未知のドキュメント → classify_document が None を返す
        extractor.extract_pages.return_value = [
            [
                TextElement(word="unknown", x0=0, y0=800, x1=100, y1=810),
                TextElement(word="doc", x0=0, y0=790, x1=100, y1=800),
                TextElement(word="type", x0=0, y0=780, x1=100, y1=790),
                TextElement(word="here", x0=0, y0=770, x1=100, y1=780),
            ]
        ]
        result = process_single_pdf("dummy.pdf", extractor, [])
        assert result is None

    def test_returns_none_for_empty_pages(self):
        extractor = MagicMock()
        extractor.extract_pages.return_value = [[]]
        result = process_single_pdf("dummy.pdf", extractor, [])
        assert result is None


class TestGenerateRenameString:
    def test_final_check(self):
        result = ProcessingResult(
            title_name="ユニットバスルーム納期最終確認票",
            file_name="test.pdf",
            sheet_data_list=[
                SheetData(
                    store_code="12345",
                    management_number="67890",
                    site_name="テスト現場",
                    confirm_day="2024/01/05",
                )
            ],
        )
        rename = generate_rename_string(result)
        assert rename == "【01-05】12345 67890 テスト現場"

    def test_quotation(self):
        result = ProcessingResult(
            title_name="御 見 積 書",
            file_name="test.pdf",
            sheet_data_list=[SheetData(store_code="12345", site_name="テスト現場")],
        )
        rename = generate_rename_string(result)
        assert rename == "12345 テスト現場"

    def test_detail(self):
        result = ProcessingResult(
            title_name="ユニットバスルームご発注確認票",
            file_name="test.pdf",
            sheet_data_list=[
                SheetData(store_code="12345", management_number="67890", site_name="テスト現場")
            ],
        )
        rename = generate_rename_string(result)
        assert rename == "12345 67890 テスト現場"

    def test_cancel(self):
        result = ProcessingResult(
            title_name="キャンセル確認票",
            file_name="test.pdf",
            sheet_data_list=[
                SheetData(
                    store_code="12345",
                    management_number="67890",
                    site_name="テスト現場",
                    is_sapporo=False,
                )
            ],
        )
        rename = generate_rename_string(result)
        assert rename == "(消)12345 67890 テスト現場"

    def test_change_spec(self):
        result = ProcessingResult(
            title_name="仕様変更確認票",
            file_name="test.pdf",
            sheet_data_list=[
                SheetData(store_code="12345", management_number="67890", site_name="テスト現場")
            ],
        )
        rename = generate_rename_string(result)
        assert rename == "(変)12345 67890 テスト現場"

    def test_empty_sheet_data_list(self):
        result = ProcessingResult(title_name="ユニットバスルーム納期最終確認票")
        assert generate_rename_string(result) == ""

    def test_unknown_title(self):
        result = ProcessingResult(
            title_name="不明な種別",
            sheet_data_list=[SheetData(store_code="12345")],
        )
        assert generate_rename_string(result) == ""
