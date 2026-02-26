"""processors モジュールのテスト."""

import datetime

from models import SheetData, TextElement
from processors.base import calculate_confirm_day, clean_site_name, trim_end_of_word
from processors.cancel import extract_cancel_data, generate_cancel_rename
from processors.change_spec import generate_change_spec_rename
from processors.detail import extract_detail_data, generate_detail_rename
from processors.final_check import generate_final_check_rename
from processors.quotation import extract_quotation_data, generate_quotation_rename


# ---------------------------------------------------------------------------
# processors.base
# ---------------------------------------------------------------------------


class TestTrimEndOfWord:
    def test_trailing_spaces(self):
        assert trim_end_of_word("value   ") == "value"

    def test_no_trailing_spaces(self):
        assert trim_end_of_word("value") == "value"

    def test_empty_string(self):
        assert trim_end_of_word("") == ""

    def test_single_char(self):
        assert trim_end_of_word("a") == "a"


class TestCleanSiteName:
    def test_removes_forbidden_chars(self):
        assert clean_site_name('test\\file:name*"<>|') == "test_file_name_"

    def test_normal_name_unchanged(self):
        assert clean_site_name("山田太郎") == "山田太郎"


class TestCalculateConfirmDay:
    def test_skips_weekend_and_holiday(self):
        holidays = ["2024/01/02"]
        shipping_day = datetime.date(2024, 1, 5)  # Friday
        result = calculate_confirm_day(shipping_day, holidays)
        assert result == datetime.date(2023, 12, 29)

    def test_empty_holidays_returns_none(self):
        result = calculate_confirm_day(datetime.date(2024, 1, 5), [])
        assert result is None

    def test_normal_weekday(self):
        holidays = ["2024/06/01"]
        shipping_day = datetime.date(2024, 6, 14)  # Friday
        result = calculate_confirm_day(shipping_day, holidays)
        # 6/13(Thu), 6/12(Wed), 6/11(Tue), 6/10(Mon) = 4 business days back
        assert result == datetime.date(2024, 6, 10)


# ---------------------------------------------------------------------------
# processors.final_check
# ---------------------------------------------------------------------------


class TestGenerateFinalCheckRename:
    def test_with_confirm_day_and_lts(self):
        data = SheetData(
            store_code="12345",
            management_number="67890",
            site_name="テスト現場",
            confirm_day="2024/01/05",
            lts="※",
        )
        assert generate_final_check_rename(data) == "【01-05】※12345 67890 テスト現場"

    def test_without_confirm_day(self):
        data = SheetData(
            store_code="12345",
            management_number="67890",
            site_name="テスト現場",
        )
        assert generate_final_check_rename(data) == "【】12345 67890 テスト現場"

    def test_missing_required_fields(self):
        data = SheetData(store_code="12345")
        assert generate_final_check_rename(data) == ""


# ---------------------------------------------------------------------------
# processors.detail
# ---------------------------------------------------------------------------


class TestExtractDetailData:
    def test_extracts_store_code(self):
        elements = [
            TextElement(word="XX12345", x0=45.0, y0=759.0, x1=100.0, y1=770.0),
        ]
        data = extract_detail_data(elements, [])
        assert data.store_code == "12345"

    def test_extracts_management_number(self):
        elements = [
            TextElement(word="MGMT001", x0=88.0, y0=803.0, x1=150.0, y1=815.0),
        ]
        data = extract_detail_data(elements, [])
        assert data.management_number == "MGMT001"


class TestGenerateDetailRename:
    def test_normal(self):
        data = SheetData(store_code="12345", management_number="67890", site_name="テスト現場")
        assert generate_detail_rename(data) == "12345 67890 テスト現場"

    def test_missing_fields(self):
        data = SheetData(store_code="12345")
        assert generate_detail_rename(data) == ""


# ---------------------------------------------------------------------------
# processors.quotation
# ---------------------------------------------------------------------------


class TestExtractQuotationData:
    def test_pattern_a(self):
        elements = [
            TextElement(word="御 見 積 書\n12345", x0=19.0, y0=772.0, x1=100.0, y1=790.0),
            TextElement(word="テスト現場  ", x0=67.0, y0=744.0, x1=200.0, y1=760.0),
        ]
        data = extract_quotation_data(elements, [])
        assert data.store_code == "12345"
        assert data.site_name == "テスト現場"

    def test_pattern_b_fallback(self):
        elements = [
            TextElement(word="御 見 積 書\nABCDE", x0=19.0, y0=783.0, x1=100.0, y1=800.0),
            TextElement(word="別現場", x0=67.0, y0=719.0, x1=200.0, y1=735.0),
        ]
        data = extract_quotation_data(elements, [])
        assert data.store_code == "ABCDE"
        assert data.site_name == "別現場"


class TestGenerateQuotationRename:
    def test_normal(self):
        data = SheetData(store_code="12345", site_name="テスト現場")
        assert generate_quotation_rename(data) == "12345 テスト現場"

    def test_missing_fields(self):
        data = SheetData(store_code="12345")
        assert generate_quotation_rename(data) == ""


# ---------------------------------------------------------------------------
# processors.cancel
# ---------------------------------------------------------------------------


class TestExtractCancelData:
    def test_sapporo_true(self):
        elements = [
            TextElement(word="12345", x0=107.0, y0=722.0, x1=200.0, y1=735.0),
            TextElement(word="MGMT001", x0=79.0, y0=786.0, x1=150.0, y1=800.0),
            TextElement(word="テスト現場", x0=107.0, y0=756.0, x1=200.0, y1=770.0),
            TextElement(word="Fｼﾞﾄﾞｳｼﾖﾘ", x0=107.0, y0=586.0, x1=200.0, y1=600.0),
        ]
        data = extract_cancel_data(elements, [])
        assert data.is_sapporo is True
        assert data.store_code == "12345"

    def test_sapporo_false(self):
        elements = [
            TextElement(word="OTHER", x0=107.0, y0=586.0, x1=200.0, y1=600.0),
        ]
        data = extract_cancel_data(elements, [])
        assert data.is_sapporo is False


class TestGenerateCancelRename:
    def test_sapporo(self):
        data = SheetData(
            store_code="12345", management_number="67890", site_name="テスト現場", is_sapporo=True
        )
        assert generate_cancel_rename(data) == "(キャンセル不可！！)12345 67890 テスト現場"

    def test_not_sapporo(self):
        data = SheetData(
            store_code="12345", management_number="67890", site_name="テスト現場", is_sapporo=False
        )
        assert generate_cancel_rename(data) == "(消)12345 67890 テスト現場"

    def test_missing_fields(self):
        data = SheetData()
        assert generate_cancel_rename(data) == ""


# ---------------------------------------------------------------------------
# processors.change_spec
# ---------------------------------------------------------------------------


class TestGenerateChangeSpecRename:
    def test_normal(self):
        data = SheetData(store_code="12345", management_number="67890", site_name="テスト現場")
        assert generate_change_spec_rename(data) == "(変)12345 67890 テスト現場"

    def test_missing_fields(self):
        data = SheetData()
        assert generate_change_spec_rename(data) == ""
