"""見積書の処理モジュール."""

from __future__ import annotations

from config import (
    QUOTATION_SITE_NAME_A,
    QUOTATION_SITE_NAME_B,
    QUOTATION_STORE_CODE_A,
    QUOTATION_STORE_CODE_B,
)
from models import SheetData, TextElement
from processors.base import clean_site_name, trim_end_of_word


def _extract_quotation_pattern(
    elements: list[TextElement],
    data: SheetData,
    store_code_x0: int,
    store_code_y0: int,
    site_name_x0: int,
    site_name_y0: int,
) -> None:
    """見積書の1パターン分の抽出を行う（data を直接更新する）."""
    for element in elements:
        if element.match(store_code_x0, store_code_y0):
            data.store_code = element.word.replace("御 見 積 書\n", "")[0:5]
        elif element.match(site_name_x0, site_name_y0):
            data.site_name = trim_end_of_word(clean_site_name(element.word))


def extract_quotation_data(elements: list[TextElement], holidays: list[str]) -> SheetData:
    """見積書からデータを抽出する.

    パターンAで抽出を試み、店コードが取れなかった場合はパターンBで再抽出する。
    holidays は他 processor とのインターフェース統一のため引数に含めるが、本処理では未使用。
    """
    data = SheetData()

    # パターンA
    _extract_quotation_pattern(
        elements,
        data,
        QUOTATION_STORE_CODE_A.x0,
        QUOTATION_STORE_CODE_A.y0,
        QUOTATION_SITE_NAME_A.x0,
        QUOTATION_SITE_NAME_A.y0,
    )

    # パターンAで店コードが取れなかった場合、パターンBを試行
    if not data.store_code:
        _extract_quotation_pattern(
            elements,
            data,
            QUOTATION_STORE_CODE_B.x0,
            QUOTATION_STORE_CODE_B.y0,
            QUOTATION_SITE_NAME_B.x0,
            QUOTATION_SITE_NAME_B.y0,
        )

    return data


def generate_quotation_rename(data: SheetData) -> str:
    """見積書のリネーム文字列を生成する.

    Returns:
        リネーム文字列。必要なフィールドが不足している場合は空文字列。
    """
    if not data.store_code or not data.site_name:
        return ""

    return data.store_code + " " + data.site_name
