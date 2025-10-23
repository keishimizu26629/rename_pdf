import PyPDF2
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import red

def add_page_numbers(input_pdf, output_pdf):
    # 入力PDFを読み込む
    with open(input_pdf, 'rb') as f:
        pdf_reader = PyPDF2.PdfFileReader(f)
        pdf_writer = PyPDF2.PdfFileWriter()

        # ページ番号のPDFをメモリ上に作成
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        # 各ページにページ番号を追加
        for page_num in range(pdf_reader.getNumPages()):
            can.setFont("Helvetica-Bold", 12)
            can.setFillColor(red)
            can.drawString(500, 10, f"Page {page_num + 1}")
            can.showPage()
        can.save()

        packet.seek(0)
        new_pdf = PyPDF2.PdfFileReader(packet)

        # 元のPDFにページ番号を重ねる
        for page_num in range(pdf_reader.getNumPages()):
            page = pdf_reader.getPage(page_num)
            page.mergePage(new_pdf.getPage(page_num))
            pdf_writer.addPage(page)

        # 出力PDFを保存
        with open(output_pdf, 'wb') as output_f:
            pdf_writer.write(output_f)

# 使用例
input_pdf = 'input.pdf'
output_pdf = 'output.pdf'
add_page_numbers(input_pdf, output_pdf)
