"""仕様明細書の処理モジュール."""

from __future__ import annotations

from config import (
    DETAIL_MANAGEMENT_NUMBER,
    DETAIL_SITE_NAME_A,
    DETAIL_SITE_NAME_B,
    DETAIL_SITE_NAME_OR_STORE_CODE,
)
from models import SheetData, TextElement
from processors.base import clean_site_name, trim_end_of_word


def extract_detail_data(elements: list[TextElement], holidays: list[str]) -> SheetData:
    """仕様明細書からデータを抽出する.

    holidays は他 processor とのインターフェース統一のため引数に含めるが、本処理では未使用。
    """
    data = SheetData()

    for element in elements:
        if element.match(DETAIL_SITE_NAME_OR_STORE_CODE.x0, DETAIL_SITE_NAME_OR_STORE_CODE.y0):
            if len(element.word.split("\n")) > 2:
                data.site_name = trim_end_of_word(clean_site_name(element.word.split("\n")[0]))
            else:
                data.store_code = element.word[-5:]
        elif element.match(DETAIL_MANAGEMENT_NUMBER.x0, DETAIL_MANAGEMENT_NUMBER.y0):
            data.management_number = element.word
        elif element.match(DETAIL_SITE_NAME_A.x0, DETAIL_SITE_NAME_A.y0) or element.match(
            DETAIL_SITE_NAME_B.x0, DETAIL_SITE_NAME_B.y0
        ):
            data.site_name = trim_end_of_word(
                clean_site_name(element.word.split("\n")[0]).replace("\u3000", "\u3000")
            )

    return data


def generate_detail_rename(data: SheetData) -> str:
    """仕様明細書のリネーム文字列を生成する.

    Returns:
        リネーム文字列。必要なフィールドが不足している場合は空文字列。
    """
    if not data.store_code or not data.management_number or not data.site_name:
        return ""

    return data.store_code + " " + data.management_number + " " + data.site_name
