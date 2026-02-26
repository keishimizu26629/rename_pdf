"""PDF 処理のオーケストレーション."""

from __future__ import annotations

import glob
import os

import PyPDF2

from classifiers import classify_document
from extractors.base import PDFTextExtractor
from models import ProcessingResult
from processors.cancel import extract_cancel_data, generate_cancel_rename
from processors.change_spec import extract_change_spec_data, generate_change_spec_rename
from processors.detail import extract_detail_data, generate_detail_rename
from processors.final_check import extract_final_check_data, generate_final_check_rename
from processors.quotation import extract_quotation_data, generate_quotation_rename
from writers.base import FileHandler, PDFConfirmDayWriter

# ドキュメント種別 → (抽出関数, リネーム生成関数, 表示名) のマッピング
_PROCESSOR_MAP: dict[str, tuple] = {
    "final_check": (extract_final_check_data, generate_final_check_rename, "ユニットバスルーム納期最終確認票"),
    "quotation": (extract_quotation_data, generate_quotation_rename, "御 見 積 書"),
    "detail": (extract_detail_data, generate_detail_rename, "ユニットバスルームご発注確認票"),
    "change_spec": (extract_change_spec_data, generate_change_spec_rename, "仕様変更確認票"),
    "cancel": (extract_cancel_data, generate_cancel_rename, "キャンセル確認票"),
}


def process_single_pdf(
    file_path: str,
    extractor: PDFTextExtractor,
    holidays: list[str],
) -> ProcessingResult | None:
    """1つの PDF ファイルを処理する.

    Args:
        file_path: PDF ファイルのパス
        extractor: テキスト抽出器
        holidays: 祝日リスト

    Returns:
        処理結果。ドキュメント種別が判定できない場合は None。
    """
    result: ProcessingResult | None = None
    doc_type: str | None = None

    for page_elements in extractor.extract_pages(file_path):
        if not page_elements:
            continue

        if doc_type is None:
            doc_type = classify_document(page_elements)
            if doc_type is None:
                continue

        processor_info = _PROCESSOR_MAP.get(doc_type)
        if processor_info is None:
            continue

        extract_fn, _rename_fn, title_name = processor_info
        sheet_data = extract_fn(page_elements, holidays)

        if result is None:
            result = ProcessingResult(title_name=title_name, file_name=file_path)
        result.sheet_data_list.append(sheet_data)

    return result


def generate_rename_string(result: ProcessingResult) -> str:
    """ProcessingResult からリネーム文字列を生成する."""
    if not result.sheet_data_list:
        return ""

    doc_type_key = _find_doc_type_key(result.title_name)
    if doc_type_key is None:
        return ""

    _, rename_fn, _ = _PROCESSOR_MAP[doc_type_key]
    return str(rename_fn(result.sheet_data_list[0]))


def _find_doc_type_key(title_name: str) -> str | None:
    """表示名からドキュメント種別キーを逆引きする."""
    for key, (_, _, name) in _PROCESSOR_MAP.items():
        if name == title_name:
            return key
    return None


def rename_pdf_file(
    result: ProcessingResult,
    target_folder: str,
    file_handler: FileHandler,
    pdf_writer: PDFConfirmDayWriter,
) -> None:
    """PDF ファイルのリネームと確定日書き込みを行う.

    Args:
        result: 処理結果
        target_folder: 出力先フォルダ
        file_handler: ファイル操作ハンドラ
        pdf_writer: 確定日 PDF 書き込み器
    """
    rename_string = generate_rename_string(result)
    if not rename_string:
        return

    if rename_string in result.file_name:
        return

    output_path = file_handler.build_output_path(target_folder, rename_string)
    result.new_file_name = output_path

    # 確定日がある場合は PDF に書き込み
    first_data = result.sheet_data_list[0]
    if first_data.confirm_day:
        pdf_writer.write(
            result.file_name,
            output_path,
            result.sheet_data_list,
            len(result.sheet_data_list),
        )
        file_handler.remove(result.file_name)
    else:
        file_handler.rename(result.file_name, output_path)

    print(rename_string)


def process_folder(
    folder_path: str,
    extractor: PDFTextExtractor,
    holidays: list[str],
    file_handler: FileHandler,
    pdf_writer: PDFConfirmDayWriter,
) -> list[ProcessingResult]:
    """フォルダ内の全 PDF を処理する.

    Args:
        folder_path: 対象フォルダパス（末尾 / 付き）
        extractor: テキスト抽出器
        holidays: 祝日リスト
        file_handler: ファイル操作ハンドラ
        pdf_writer: 確定日 PDF 書き込み器

    Returns:
        処理済みの結果リスト
    """
    results: list[ProcessingResult] = []
    for pdf_path in glob.glob(folder_path + "*.pdf"):
        result = process_single_pdf(pdf_path, extractor, holidays)
        if result and result.sheet_data_list:
            rename_pdf_file(result, folder_path, file_handler, pdf_writer)
            if result.new_file_name:
                results.append(result)

    return results


def merge_files_for_posting(results: list[ProcessingResult], target_folder: str) -> None:
    """投函用の PDF マージ処理.

    最終確認票と対応する仕様明細書を結合し、(投函用) プレフィックス付きで保存する。

    Args:
        results: 処理済み結果リスト
        target_folder: 出力先フォルダパス（末尾 / 付き）
    """
    for result in results:
        if result.title_name != "ユニットバスルーム納期最終確認票":
            continue

        merge_file = result.new_file_name.replace(target_folder, "")
        matching_file = merge_file.split("】", 1)[1].replace("※", "")

        matching_path = os.path.join(target_folder, matching_file)
        if not os.path.exists(matching_path):
            print(f"Matching file not found for: {merge_file}")
            continue

        try:
            pdf_file_merger = PyPDF2.PdfWriter()

            with open(os.path.join(target_folder, merge_file), "rb") as f1:
                pdf1 = PyPDF2.PdfReader(f1)
                for page in pdf1.pages:
                    pdf_file_merger.add_page(page)

            with open(matching_path, "rb") as f2:
                pdf2 = PyPDF2.PdfReader(f2)
                for page in pdf2.pages:
                    pdf_file_merger.add_page(page)

            output_filename = os.path.join(target_folder, "(投函用)" + merge_file)
            with open(output_filename, "wb") as output_file:
                pdf_file_merger.write(output_file)

            print(f"Merged files: {merge_file} and {matching_file}")
        except FileNotFoundError as e:
            _log_error(target_folder, merge_file, matching_file, str(e))
            continue

    print("PDF merging process completed.")


def _log_error(target_folder: str, merge_file: str, matching_file: str, error_message: str) -> None:
    """エラーログを書き出す."""
    import datetime

    log_dir = "./var"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_file = os.path.join(log_dir, "log.txt")
    with open(log_file, "a") as f:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"[{timestamp}] Error:\n")
        f.write(f"フォルダ名: {target_folder}\n")
        f.write(f"最終確認票: {merge_file}\n")
        f.write(f"結合するファイル名: {matching_file}\n")
        f.write(f"Error Message: {error_message}\n\n")
