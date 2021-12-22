import io
import PyPDF2
from PyPDF2.pdf import PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import mm

# ページ番号の下からの位置
PAGE_BOTTOM = 10 * mm
# ページ番号のプレフィックス
PAGE_PREFIX = "12"
# フォント登録
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))

def add_page_number(input_file: str, output_file: str, start_num: int = 1):
    """
    既存PDFにページ番号を追加する
    """
    # 既存PDF（ページを付けるPDF）
    fi = open(input_file, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(fi)
    pages_num = pdf_reader.getNumPages()

    # ページ番号を付けたPDFの書き込み用
    pdf_writer = PyPDF2.PdfFileWriter()

    # ページ番号だけのPDFをメモリ（binary stream）に作成
    bs = io.BytesIO()
    c = canvas.Canvas(bs)
    for i in range(0, pages_num):
        # 既存PDF
        pdf_page = pdf_reader.getPage(i)
        # PDFページのサイズ
        page_size = get_page_size(pdf_page)
        # ページ番号のPDF作成
        create_page_number_pdf(c, page_size, i + start_num)
    c.save()

    # ページ番号だけのPDFをメモリから読み込み（seek操作はPyPDF2に実装されているので不要）
    pdf_num_reader = PyPDF2.PdfFileReader(bs)

    # 既存PDFに１ページずつページ番号を付ける
    for i in range(0, pages_num):
        # 既存PDF
        pdf_page = pdf_reader.getPage(i)
        # ページ番号だけのPDF
        pdf_num = pdf_num_reader.getPage(i)
        # ２つのPDFを重ねる
        pdf_page.mergePage(pdf_num)
        pdf_writer.addPage(pdf_page)

    # ページ番号を付けたPDFを保存
    fo = open(output_file, 'wb')
    pdf_writer.write(fo)

    bs.close()
    fi.close()
    fo.close()


def create_page_number_pdf(c: canvas.Canvas, page_size: tuple, page_num: int):
    """
    ページ番号だけのPDFを作成
    """
    c.setPageSize(page_size)
    c.setFont("HeiseiKakuGo-W5", 18)
    c.drawCentredString(50,
                        page_size[1] - 103,
                        # PAGE_PREFIX + str(page_num))
                        PAGE_PREFIX)
    c.drawCentredString(94,
                        page_size[1] - 103,
                        # PAGE_PREFIX + str(page_num))
                        PAGE_PREFIX)

    c.showPage()


def get_page_size(page: PageObject) -> tuple:
    """
    既存PDFからページサイズ（幅, 高さ）を取得する
    """
    page_box = page.mediaBox
    width = page_box.getUpperRight_x() - page_box.getLowerLeft_x()
    height = page_box.getUpperRight_y() - page_box.getLowerLeft_y()
    return float(width), float(height)


if __name__ == '__main__':
    # テスト用
    infile = './sample/test1.pdf'
    outfile = './sample/test1_paged.pdf'
    # outfile = infile
    add_page_number(infile, outfile)
