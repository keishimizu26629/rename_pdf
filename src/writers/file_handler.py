"""ファイル操作の実装."""

from __future__ import annotations

import glob
import os


class DefaultFileHandler:
    """os / glob を使用したファイル操作."""

    def rename(self, src: str, dst: str) -> None:
        """ファイルをリネームする."""
        os.rename(src, dst)

    def remove(self, path: str) -> None:
        """ファイルを削除する."""
        os.remove(path)

    def find_duplicates(self, pattern: str) -> int:
        """パターンに一致するファイル数を返す."""
        return len(glob.glob(pattern))

    def build_output_path(self, target_folder: str, base_name: str) -> str:
        """重複を考慮した出力ファイルパスを生成する.

        Args:
            target_folder: 出力先フォルダパス
            base_name: 拡張子なしのファイル名

        Returns:
            重複回避済みのフルパス (.pdf 付き)
        """
        duplicated_count = self.find_duplicates(os.path.join(target_folder, base_name + "*.pdf"))
        if duplicated_count == 0:
            return os.path.join(target_folder, base_name + ".pdf")
        return os.path.join(target_folder, base_name + "(" + str(duplicated_count + 1) + ")" + ".pdf")
