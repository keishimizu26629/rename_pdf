"""models モジュールのテスト."""

from models import ProcessingResult, SheetData, TextElement


class TestTextElement:
    def test_match_exact(self):
        el = TextElement(word="test", x0=161.0, y0=650.0, x1=200.0, y1=660.0)
        assert el.match(161, 650) is True

    def test_match_floor(self):
        el = TextElement(word="test", x0=161.7, y0=650.9, x1=200.0, y1=660.0)
        assert el.match(161, 650) is True

    def test_match_fails_on_different_coords(self):
        el = TextElement(word="test", x0=161.0, y0=650.0, x1=200.0, y1=660.0)
        assert el.match(162, 650) is False
        assert el.match(161, 651) is False


class TestSheetData:
    def test_defaults(self):
        data = SheetData()
        assert data.store_code == ""
        assert data.management_number == ""
        assert data.site_name == ""
        assert data.confirm_day == ""
        assert data.lts == ""
        assert data.is_sapporo is False


class TestProcessingResult:
    def test_defaults(self):
        result = ProcessingResult()
        assert result.title_name == ""
        assert result.file_name == ""
        assert result.sheet_data_list == []
        assert result.new_file_name == ""

    def test_sheet_data_list_independent(self):
        r1 = ProcessingResult()
        r2 = ProcessingResult()
        r1.sheet_data_list.append(SheetData(store_code="12345"))
        assert len(r2.sheet_data_list) == 0
