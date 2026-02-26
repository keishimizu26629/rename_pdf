"""データクラス定義モジュール."""

from __future__ import annotations

import math
from dataclasses import dataclass, field


@dataclass
class TextElement:
    """PDFから抽出したテキスト要素."""

    word: str
    x0: float
    y0: float
    x1: float
    y1: float

    def match(self, x0: int, y0: int) -> bool:
        """座標が指定値と一致するか判定する."""
        return math.floor(self.x0) == x0 and math.floor(self.y0) == y0


@dataclass
class SheetData:
    """各シートから抽出した処理用データ."""

    store_code: str = ""
    management_number: str = ""
    site_name: str = ""
    confirm_day: str = ""
    lts: str = ""
    is_sapporo: bool = False


@dataclass
class ProcessingResult:
    """ファイル処理結果."""

    title_name: str = ""
    file_name: str = ""
    sheet_data_list: list[SheetData] = field(default_factory=list)
    new_file_name: str = ""
