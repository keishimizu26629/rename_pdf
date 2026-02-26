"""最終確認票の処理モジュール."""

from __future__ import annotations

import datetime

import dateutil.relativedelta

from config import (
    FINAL_CONSTRUCTION_TYPE,
    FINAL_MANAGEMENT_NUMBER,
    FINAL_SAPPORO_ADDRESS,
    FINAL_SHIPPING_DAY,
    FINAL_SITE_NAME,
    FINAL_STORE_CODE,
)
from models import SheetData, TextElement
from processors.base import calculate_confirm_day, clean_site_name, trim_end_of_word


def extract_final_check_data(elements: list[TextElement], holidays: list[str]) -> SheetData:
    """最終確認票からデータを抽出する."""
    data = SheetData()

    for element in elements:
        if element.match(FINAL_STORE_CODE.x0, FINAL_STORE_CODE.y0):
            data.store_code = element.word[0:5]
        elif element.match(FINAL_MANAGEMENT_NUMBER.x0, FINAL_MANAGEMENT_NUMBER.y0):
            data.management_number = element.word.split("\n")[2]
        elif element.match(FINAL_SITE_NAME.x0, FINAL_SITE_NAME.y0):
            data.site_name = trim_end_of_word(clean_site_name(element.word))
        elif (
            element.match(FINAL_SHIPPING_DAY.x0, FINAL_SHIPPING_DAY.y0)
            and element.word.split(" ")[0].split("月")[0] != ""
            and element.word != "発送"
        ):
            shipping_day = datetime.date(
                datetime.datetime.now().date().year,
                int(element.word.split(" ")[0].split("月")[0]),
                int(element.word.split(" ")[0].split("月")[1][:-1]),
            )
            if datetime.datetime.now().date() > shipping_day:
                shipping_day = shipping_day + dateutil.relativedelta.relativedelta(years=1)
            confirm_day = calculate_confirm_day(shipping_day, holidays)
            if confirm_day is None:
                data.confirm_day = ""
            else:
                data.confirm_day = confirm_day.strftime("%Y/%m/%d")
        elif element.match(FINAL_CONSTRUCTION_TYPE.x0, FINAL_CONSTRUCTION_TYPE.y0):
            if "工事区分" in element.word:
                data.lts = "※"
        elif element.match(FINAL_SAPPORO_ADDRESS.x0, FINAL_SAPPORO_ADDRESS.y0):
            if element.word.startswith("札幌市") or element.word.startswith("北海道"):
                data.is_sapporo = True
            else:
                data.is_sapporo = False

    return data


def generate_final_check_rename(data: SheetData) -> str:
    """最終確認票のリネーム文字列を生成する.

    Returns:
        リネーム文字列。必要なフィールドが不足している場合は空文字列。
    """
    if not data.store_code or not data.management_number or not data.site_name:
        return ""

    confirm_day = ""
    if data.confirm_day:
        confirm_day = data.confirm_day.split("/")[1] + "-" + data.confirm_day.split("/")[2]

    return "【" + confirm_day + "】" + data.lts + data.store_code + " " + data.management_number + " " + data.site_name
