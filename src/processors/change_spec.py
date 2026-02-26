"""仕様変更確認票の処理モジュール."""

from __future__ import annotations

from models import SheetData, TextElement
from processors.detail import extract_detail_data


def extract_change_spec_data(elements: list[TextElement], holidays: list[str]) -> SheetData:
    """仕様変更確認票からデータを抽出する.

    DetailSheet と同じ抽出ロジックを使用する。
    """
    return extract_detail_data(elements, holidays)


def generate_change_spec_rename(data: SheetData) -> str:
    """仕様変更確認票のリネーム文字列を生成する.

    Returns:
        リネーム文字列。必要なフィールドが不足している場合は空文字列。
    """
    if not data.store_code or not data.management_number or not data.site_name:
        return ""

    return "(変)" + data.store_code + " " + data.management_number + " " + data.site_name
