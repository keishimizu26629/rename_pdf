"""キャンセル確認票の処理モジュール."""

from __future__ import annotations

from config import (
    CANCEL_MANAGEMENT_NUMBER,
    CANCEL_SAPPORO,
    CANCEL_SITE_NAME,
    CANCEL_STORE_CODE,
)
from models import SheetData, TextElement
from processors.base import clean_site_name, trim_end_of_word


def extract_cancel_data(elements: list[TextElement], holidays: list[str]) -> SheetData:
    """キャンセル確認票からデータを抽出する.

    holidays は他 processor とのインターフェース統一のため引数に含めるが、本処理では未使用。
    """
    data = SheetData()

    for element in elements:
        if element.match(CANCEL_STORE_CODE.x0, CANCEL_STORE_CODE.y0):
            data.store_code = element.word
        elif element.match(CANCEL_MANAGEMENT_NUMBER.x0, CANCEL_MANAGEMENT_NUMBER.y0):
            data.management_number = element.word
        elif element.match(CANCEL_SITE_NAME.x0, CANCEL_SITE_NAME.y0):
            data.site_name = trim_end_of_word(clean_site_name(element.word))
        elif element.match(CANCEL_SAPPORO.x0, CANCEL_SAPPORO.y0):
            if element.word == "Fｼﾞﾄﾞｳｼﾖﾘ":
                data.is_sapporo = True
            else:
                data.is_sapporo = False

    return data


def generate_cancel_rename(data: SheetData) -> str:
    """キャンセル確認票のリネーム文字列を生成する.

    Returns:
        リネーム文字列。必要なフィールドが不足している場合は空文字列。
    """
    if not data.store_code or not data.management_number or not data.site_name:
        return ""

    cancel_string = "(キャンセル不可！！)" if data.is_sapporo else "(消)"
    return cancel_string + data.store_code + " " + data.management_number + " " + data.site_name
