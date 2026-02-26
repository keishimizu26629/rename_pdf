"""座標定数・描画設定・テキスト定数を集約するモジュール."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CoordinateMatch:
    """PDF上のテキスト要素を座標でマッチングするための定数."""

    x0: int
    y0: int


# ---------------------------------------------------------------------------
# FinalCheckSheet 座標
# ---------------------------------------------------------------------------
FINAL_STORE_CODE = CoordinateMatch(x0=161, y0=650)
FINAL_MANAGEMENT_NUMBER = CoordinateMatch(x0=198, y0=523)
FINAL_SITE_NAME = CoordinateMatch(x0=198, y0=625)
FINAL_SHIPPING_DAY = CoordinateMatch(x0=158, y0=472)
FINAL_CONSTRUCTION_TYPE = CoordinateMatch(x0=334, y0=523)
FINAL_SAPPORO_ADDRESS = CoordinateMatch(x0=158, y0=344)

# ---------------------------------------------------------------------------
# DetailSheet 座標
# ---------------------------------------------------------------------------
DETAIL_SITE_NAME_OR_STORE_CODE = CoordinateMatch(x0=45, y0=759)
DETAIL_MANAGEMENT_NUMBER = CoordinateMatch(x0=88, y0=803)
DETAIL_SITE_NAME_A = CoordinateMatch(x0=45, y0=782)
DETAIL_SITE_NAME_B = CoordinateMatch(x0=45, y0=780)

# ---------------------------------------------------------------------------
# CancelSheet 座標
# ---------------------------------------------------------------------------
CANCEL_STORE_CODE = CoordinateMatch(x0=107, y0=722)
CANCEL_MANAGEMENT_NUMBER = CoordinateMatch(x0=79, y0=786)
CANCEL_SITE_NAME = CoordinateMatch(x0=107, y0=756)
CANCEL_SAPPORO = CoordinateMatch(x0=107, y0=586)

# ---------------------------------------------------------------------------
# QuotationSheet 座標 (パターンA / パターンB)
# ---------------------------------------------------------------------------
QUOTATION_STORE_CODE_A = CoordinateMatch(x0=19, y0=772)
QUOTATION_SITE_NAME_A = CoordinateMatch(x0=67, y0=744)
QUOTATION_STORE_CODE_B = CoordinateMatch(x0=19, y0=783)
QUOTATION_SITE_NAME_B = CoordinateMatch(x0=67, y0=719)

# ---------------------------------------------------------------------------
# PDF描画座標
# ---------------------------------------------------------------------------
CONFIRM_DAY_MONTH_X = 254
CONFIRM_DAY_DAY_X = 296
CONFIRM_DAY_Y_OFFSET = 130

SAPPORO_TEXT_X = 20
SAPPORO_TEXT_Y = 90

# ---------------------------------------------------------------------------
# 札幌DC テキスト
# ---------------------------------------------------------------------------
SAPPORO_DC_TEXT = (
    "キャンセル：出荷日８日前の午前中まで\n"
    "仕様変更：出荷日５日前の午前中までにご依頼お願いします。\n"
    "内容によってお受けできない場合があります。\n"
    "※ 長納期品、基準外品等に関しては都度ご確認をお願いします。"
)
