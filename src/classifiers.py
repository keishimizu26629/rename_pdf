"""ドキュメント種別判定モジュール."""

from __future__ import annotations

from models import TextElement


def classify_document(elements: list[TextElement]) -> str | None:
    """PDF要素リストからドキュメント種別を判定する.

    elements を y1 の降順でソートし、上位要素のテキストで種別を判定する。

    Returns:
        "final_check" | "quotation" | "detail" | "change_spec" | "cancel" | None
    """
    sorted_elements = sorted(elements, key=lambda x: x.y1, reverse=True)

    if (
        "ユニットバスルーム納期最終確認票" in sorted_elements[2].word
        or "ユニットバスルーム納期最終確認票" in sorted_elements[1].word
    ):
        return "final_check"

    if "御 見 積 書" in sorted_elements[0].word:
        return "quotation"

    if "ユニットバスルームご発注確認票" in sorted_elements[1].word:
        return "detail"

    if "仕様変更確認票" in sorted_elements[1].word:
        return "change_spec"

    if "キ ャ ン セ ル 確 認 票" in sorted_elements[0].word:
        return "cancel"

    return None
