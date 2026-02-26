"""ファイル操作・PDF書き込みの Protocol 定義."""

from __future__ import annotations

from typing import Protocol

from models import SheetData


class FileHandler(Protocol):
    """ファイル操作の Protocol."""

    def rename(self, src: str, dst: str) -> None:
        """ファイルをリネームする."""
        ...

    def remove(self, path: str) -> None:
        """ファイルを削除する."""
        ...

    def find_duplicates(self, pattern: str) -> int:
        """パターンに一致するファイル数を返す."""
        ...

    def build_output_path(self, target_folder: str, base_name: str) -> str:
        """重複を考慮した出力ファイルパスを生成する.

        Args:
            target_folder: 出力先フォルダパス
            base_name: 拡張子なしのファイル名

        Returns:
            重複回避済みのフルパス (.pdf 付き)
        """
        ...


class PDFConfirmDayWriter(Protocol):
    """確定日 PDF 書き込みの Protocol."""

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
            holidays: 祝日リスト（未使用だが将来拡張用）
        """
        ...
