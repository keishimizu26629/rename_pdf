"""PDF テキスト抽出の Protocol 定義."""

from __future__ import annotations

from collections.abc import Iterator
from typing import Protocol

from models import TextElement


class PDFTextExtractor(Protocol):
    """PDF ファイルからページ単位でテキスト要素を抽出する Protocol."""

    def extract_pages(self, file_path: str) -> Iterator[list[TextElement]]:
        """PDF の各ページからテキスト要素のリストを yield する.

        Args:
            file_path: PDF ファイルのパス

        Yields:
            ページごとの TextElement リスト
        """
        ...
