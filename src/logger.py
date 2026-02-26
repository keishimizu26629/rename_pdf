"""ログ設定モジュール."""

from __future__ import annotations

import datetime
import logging
import os
import sys


def _get_log_dir() -> str:
    """ログディレクトリのパスを返す（exe / 開発 両対応）."""
    if getattr(sys, "frozen", False):
        # exe 実行時: 実行ファイルと同階層
        base_dir = os.path.dirname(sys.executable)
    else:
        # 開発時: src/ の親 = プロジェクトルート
        base_dir = os.path.dirname(os.path.dirname(__file__))
    return os.path.join(base_dir, "logs")


def setup_logging() -> None:
    """logging モジュールを設定し、日付別ログファイルへの出力を開始する."""
    log_dir = _get_log_dir()
    os.makedirs(log_dir, exist_ok=True)

    today = datetime.date.today().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"renamePdf_{today}.log")

    handler = logging.FileHandler(log_file, encoding="utf-8")
    handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(handler)


def cleanup_old_logs(days: int = 10) -> None:
    """指定日数を超えた古いログファイルを削除する."""
    log_dir = _get_log_dir()
    if not os.path.isdir(log_dir):
        return

    cutoff = datetime.date.today() - datetime.timedelta(days=days)

    for filename in os.listdir(log_dir):
        if not filename.startswith("renamePdf_") or not filename.endswith(".log"):
            continue
        # renamePdf_YYYY-MM-DD.log → YYYY-MM-DD を抽出
        date_str = filename[len("renamePdf_") : -len(".log")]
        try:
            file_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            continue
        if file_date < cutoff:
            os.remove(os.path.join(log_dir, filename))
