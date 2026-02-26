"""processors 共通ユーティリティ関数."""

from __future__ import annotations

import datetime
import re


def trim_end_of_word(s: str) -> str:
    """文字列末尾の空白を除去する."""
    i = 1
    maxcount = len(s)
    while i < maxcount:
        if s[-1] == " ":
            s = s[:-1]
        else:
            break
        i += 1
    return s


def clean_site_name(word: str) -> str:
    """現場名からファイル名に使用できない文字を置換する."""
    return re.sub(r'[\\/:*?"<>|]+', "_", word)


def calculate_confirm_day(shipping_day: datetime.date, holidays: list[str]) -> datetime.date | None:
    """出荷日から4営業日前の確定日を計算する.

    holidays が空の場合は None を返す。
    """
    if len(holidays) == 0:
        return None
    business_day_count = 0
    while business_day_count < 4:
        shipping_day = shipping_day - datetime.timedelta(days=1)
        if (
            shipping_day.weekday() != 5
            and shipping_day.weekday() != 6
            and shipping_day.strftime("%Y/%m/%d") not in holidays
        ):
            business_day_count += 1
    return shipping_day
