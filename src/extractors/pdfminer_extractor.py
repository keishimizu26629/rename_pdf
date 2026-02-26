"""pdfminer.six を使用した PDF テキスト抽出の実装."""

from __future__ import annotations

from collections.abc import Iterator

import pdfminer.converter
import pdfminer.layout
import pdfminer.pdfinterp
import pdfminer.pdfpage

from models import TextElement


class PdfMinerTextExtractor:
    """pdfminer.six ベースの PDF テキスト抽出."""

    def extract_pages(self, file_path: str) -> Iterator[list[TextElement]]:
        """PDF の各ページからテキスト要素を抽出する.

        Args:
            file_path: PDF ファイルのパス

        Yields:
            ページごとの TextElement リスト
        """
        resource_manager = pdfminer.pdfinterp.PDFResourceManager()
        device = pdfminer.converter.PDFPageAggregator(resource_manager, laparams=pdfminer.layout.LAParams())
        try:
            with open(file_path, "rb") as fp:
                interpreter = pdfminer.pdfinterp.PDFPageInterpreter(resource_manager, device)
                for page in pdfminer.pdfpage.PDFPage.get_pages(fp):
                    interpreter.process_page(page)
                    layout = device.get_result()
                    elements: list[TextElement] = []
                    for lt in layout:
                        if isinstance(lt, pdfminer.layout.LTTextContainer):
                            elements.append(
                                TextElement(
                                    word=lt.get_text().strip(),
                                    x0=lt.x0,
                                    y0=lt.y0,
                                    x1=lt.x1,
                                    y1=lt.y1,
                                )
                            )
                    yield elements
        finally:
            device.close()
