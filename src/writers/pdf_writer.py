"""確定日 PDF 書き込みの実装."""

from __future__ import annotations

import io

import PyPDF2
import reportlab.pdfgen.canvas

from config import (
    CONFIRM_DAY_DAY_X,
    CONFIRM_DAY_MONTH_X,
    CONFIRM_DAY_Y_OFFSET,
    SAPPORO_DC_TEXT,
    SAPPORO_TEXT_X,
    SAPPORO_TEXT_Y,
)
from models import SheetData


class DefaultPDFConfirmDayWriter:
    """確定日を PDF に書き込む実装."""

    def write(
        self,
        file_path: str,
        output_path: str,
        sheet_data_list: list[SheetData],
        page_count: int,
        *,
        holidays: list[str] | None = None,
    ) -> None:
        """確定日を既存 PDF に重ね合わせて出力する.

        Args:
            file_path: 入力 PDF のパス
            output_path: 出力 PDF のパス
            sheet_data_list: 各ページの SheetData リスト
            page_count: ページ数
            holidays: 祝日リスト（将来拡張用）
        """
        with open(file_path, "rb") as fi:
            pdf_reader = PyPDF2.PdfReader(fi)
            pdf_writer = PyPDF2.PdfWriter()

            # 確定日オーバーレイ PDF をメモリ上に作成
            bs = io.BytesIO()
            canvas = reportlab.pdfgen.canvas.Canvas(bs)
            for i in range(page_count):
                pdf_page = pdf_reader.pages[i]
                page_size = _get_page_size(pdf_page)
                _draw_confirm_day_page(canvas, page_size, sheet_data_list[i])
            canvas.save()

            # オーバーレイ PDF を読み込んで既存ページに重ね合わせ
            pdf_overlay_reader = PyPDF2.PdfReader(bs)
            for i in range(page_count):
                pdf_page = pdf_reader.pages[i]
                overlay_page = pdf_overlay_reader.pages[i]
                pdf_page.merge_page(overlay_page)
                pdf_writer.add_page(pdf_page)

            with open(output_path, "wb") as fo:
                pdf_writer.write(fo)
            bs.close()


def _get_page_size(page: PyPDF2.PageObject) -> tuple[float, float]:
    """PDF ページのサイズ（幅, 高さ）を取得する."""
    page_box = page.mediabox
    width = float(page_box.upper_right[0] - page_box.lower_left[0])
    height = float(page_box.upper_right[1] - page_box.lower_left[1])
    return width, height


def _draw_confirm_day_page(
    canvas: reportlab.pdfgen.canvas.Canvas,
    page_size: tuple[float, float],
    sheet_data: SheetData,
) -> None:
    """1ページ分の確定日描画を行う.

    Args:
        canvas: reportlab Canvas
        page_size: (幅, 高さ) タプル
        sheet_data: 該当ページの SheetData
    """
    confirm_day = sheet_data.confirm_day
    if not confirm_day:
        canvas.setPageSize(page_size)
        canvas.showPage()
        return

    # 月・日を取得（先頭ゼロ除去）
    parts = confirm_day.split("/")
    ship_month = parts[1].lstrip("0") or "0"
    ship_day = parts[2].lstrip("0") or "0"

    canvas.setPageSize(page_size)
    try:
        canvas.setFont("MS P ゴシック", 16)
    except KeyError:
        canvas.setFont("Helvetica", 16)

    canvas.drawCentredString(CONFIRM_DAY_MONTH_X, page_size[1] - CONFIRM_DAY_Y_OFFSET, ship_month)
    canvas.drawCentredString(CONFIRM_DAY_DAY_X, page_size[1] - CONFIRM_DAY_Y_OFFSET, ship_day)

    # 札幌DC対応の場合、文言を追加
    if sheet_data.is_sapporo:
        try:
            canvas.setFont("MS P ゴシック", 12)
        except KeyError:
            canvas.setFont("Helvetica", 12)
        text_object = canvas.beginText(SAPPORO_TEXT_X, SAPPORO_TEXT_Y)
        for line in SAPPORO_DC_TEXT.split("\n"):
            text_object.textLine(line)
        canvas.drawText(text_object)

    canvas.showPage()
