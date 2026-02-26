"""PDF リネームツール – GUI エントリポイント."""

from __future__ import annotations

import contextlib
import csv
import datetime
import logging
import os
import traceback

import reportlab.pdfbase
import reportlab.pdfbase.ttfonts

# GUI imports
try:
    import tkinter
    from tkinter import filedialog, ttk

    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

from extractors.pdfminer_extractor import PdfMinerTextExtractor
from logger import cleanup_old_logs, setup_logging
from services import merge_files_for_posting, process_folder
from writers.file_handler import DefaultFileHandler
from writers.pdf_writer import DefaultPDFConfirmDayWriter


def change_words(words: list[str], func) -> None:
    """リスト内の各要素に関数を適用して置換する."""
    for i, word in enumerate(words):
        words[i] = func(word)


def register_fonts() -> None:
    """PDF 生成用フォントを登録する（クロスプラットフォーム対応）."""
    import platform

    if platform.system() == "Windows":
        try:
            reportlab.pdfbase.pdfmetrics.registerFont(
                reportlab.pdfbase.ttfonts.TTFont("Meiryo UI", "c:/Windows/Fonts/meiryob.ttc")
            )
            reportlab.pdfbase.pdfmetrics.registerFont(
                reportlab.pdfbase.ttfonts.TTFont("MS P ゴシック", "c:/Windows/Fonts/msgothic.ttc")
            )
        except Exception:
            pass
    elif platform.system() == "Darwin":
        try:
            reportlab.pdfbase.pdfmetrics.registerFont(
                reportlab.pdfbase.ttfonts.TTFont("Meiryo UI", "/System/Library/Fonts/Helvetica.ttc")
            )
            reportlab.pdfbase.pdfmetrics.registerFont(
                reportlab.pdfbase.ttfonts.TTFont("MS P ゴシック", "/System/Library/Fonts/Helvetica.ttc")
            )
        except Exception:
            pass


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

if GUI_AVAILABLE:

    class ErrorDialog(tkinter.Toplevel):
        """エラー表示ダイアログ."""

        def __init__(self, parent: tkinter.Tk, error_message: str) -> None:
            super().__init__(parent)
            self.title("エラー")
            self.geometry("500x300")

            error_label = tkinter.Label(self, text="エラーが発生しました:", font=("Helvetica", 12, "bold"))
            error_label.pack(pady=10)

            error_text = tkinter.Text(self, wrap=tkinter.WORD, width=60, height=10)
            error_text.insert(tkinter.END, error_message)
            error_text.config(state=tkinter.DISABLED)
            error_text.pack(padx=10, pady=10)

            close_button = tkinter.Button(self, text="閉じる", command=self.destroy)
            close_button.pack(pady=10)


def main() -> None:
    """メインエントリポイント – GUI を起動して PDF 処理を実行する."""
    if not GUI_AVAILABLE:
        print("GUI not available. This application requires tkinter.")
        return

    setup_logging()
    cleanup_old_logs()

    register_fonts()

    # 祝日 CSV の読み込み
    # exe 実行時はカレントディレクトリ、開発時は data/ ディレクトリを参照
    _csv_candidates = ["休日.csv", os.path.join(os.path.dirname(os.path.dirname(__file__)), "data", "休日.csv")]
    holidays: list[str] = []
    _csv_path = next((p for p in _csv_candidates if os.path.exists(p)), None)
    if _csv_path:
        with open(_csv_path, encoding="utf-8") as f:
            reader = csv.reader(f)
            holidays = [rows[0] for rows in reader]
        holidays = holidays[1:]
        change_words(holidays, lambda word: datetime.datetime.strptime(word, "%Y/%m/%d").strftime("%Y/%m/%d"))
    else:
        logging.warning("祝日CSVファイルが見つかりません: %s", _csv_candidates)

    # DI: 実装インスタンスの生成
    extractor = PdfMinerTextExtractor()
    file_handler = DefaultFileHandler()
    pdf_writer = DefaultPDFConfirmDayWriter()

    # root の作成
    root = tkinter.Tk()
    root.title("renamePdf")
    with contextlib.suppress(tkinter.TclError):
        root.iconbitmap("icons/icon.ico")

    # Frame1 の作成
    frame1 = ttk.Frame(root, padding=10)
    frame1.grid(row=0, column=1, sticky=tkinter.E)

    IDirLabel = ttk.Label(frame1, text="フォルダ参照＞＞", padding=(5, 2))
    IDirLabel.pack(side=tkinter.LEFT)

    entry1 = tkinter.StringVar()
    IDirEntry = ttk.Entry(frame1, textvariable=entry1, width=30)
    IDirEntry.pack(side=tkinter.LEFT)

    if os.path.exists("./var/path.txt"):
        with open("./var/path.txt", encoding="utf-8") as f:
            s = f.read()
            entry1.set(s)

    def dirdialog_clicked() -> None:
        dirPath = entry1.get()
        iDir = dirPath if dirPath else "C:\\"
        iDirPath = filedialog.askdirectory(initialdir=iDir)
        if iDirPath != "":
            entry1.set(iDirPath)
            if not os.path.exists("./var"):
                os.makedirs("./var")
            with open("./var/path.txt", "w", encoding="utf-8") as f:
                f.write(iDirPath)

    IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    IDirButton.pack(side=tkinter.LEFT)

    def conduct_main() -> None:
        try:
            dirPath = entry1.get()
            if dirPath:
                target_folder_name = dirPath + "/"
                results = process_folder(
                    target_folder_name, extractor, holidays, file_handler, pdf_writer
                )
                merge_files_for_posting(results, target_folder_name)
            else:
                raise ValueError("フォルダを指定してください！")
        except Exception as e:
            logging.exception("conduct_main でエラーが発生しました")
            error_message = f"エラーの種類: {type(e).__name__}\n"
            error_message += f"エラーメッセージ: {str(e)}\n\n"
            error_message += "詳細なエラー情報:\n"
            error_message += traceback.format_exc()
            ErrorDialog(root, error_message)

    # Frame3 の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid(row=5, column=1, sticky=tkinter.W)

    button1 = ttk.Button(frame3, text="実行", command=conduct_main)
    button1.pack(fill="x", padx=30, side="left")

    button2 = ttk.Button(frame3, text="閉じる", command=root.destroy)
    button2.pack(fill="x", padx=30, side="left")

    root.geometry("400x130+200+300")
    root.mainloop()


if __name__ == "__main__":
    main()
