import os, glob, re, math
import pdfminer.pdfinterp, pdfminer.converter, pdfminer.pdfpage, pdfminer.layout
import csv, datetime
import dateutil.relativedelta

import io, PyPDF2
# from PyPDF2.pdf import PageObject
import reportlab.pdfgen
import reportlab.pdfbase
import reportlab.pdfbase.cidfonts
import reportlab.pdfbase.ttfonts

# reportlab.pdfbase.pdfmetrics.registerFont(reportlab.pdfbase.cidfonts.UnicodeCIDFont("HeiseiKakuGo-W5"))
# reportlab.pdfbase.pdfmetrics.registerFont(reportlab.pdfbase.ttfonts.TTFont('YuGothR', 'C:\Windows\Fonts\游ゴシック\YuGothR.ttc'))
reportlab.pdfbase.pdfmetrics.registerFont(reportlab.pdfbase.ttfonts.TTFont('MS P ゴシック', 'c:/Windows/Fonts/msgothic.ttc'))

def add_page_number(input_file: str, output_file: str):

    fi = open(input_file, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(fi)
    pages_num = pdf_reader.getNumPages()

    # ページ番号を付けたPDFの書き込み用
    pdf_writer = PyPDF2.PdfFileWriter()

    # ページ番号だけのPDFをメモリ（binary stream）に作成
    bs = io.BytesIO()
    c = reportlab.pdfgen.canvas.Canvas(bs)
    # 既存PDF
    pdf_page = pdf_reader.getPage(0)
    # PDFページのサイズ
    page_size = get_page_size(pdf_page)
    # ページ番号のPDF作成
    create_page_number_pdf(c, page_size)
    c.save()

    # ページ番号だけのPDFをメモリから読み込み（seek操作はPyPDF2に実装されているので不要）
    pdf_num_reader = PyPDF2.PdfFileReader(bs)

    # 既存PDFに１ページずつページ番号を付ける
    for i in range(0, pages_num):
        # 既存PDF
        pdf_page = pdf_reader.getPage(i)
        if i == 0:
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
    os.remove(file_name)

def create_page_number_pdf(c: reportlab.pdfgen.canvas.Canvas, page_size: tuple):
    """
    ページ番号だけのPDFを作成
    """

    ship_month = number_required_for['確定日'].split('/')[1]
    ship_day = number_required_for['確定日'].split('/')[2]

    c.setPageSize(page_size)
    c.setFont('MS P ゴシック', 16)
    c.drawCentredString(50,
                        page_size[1] - 104,
                        # PAGE_PREFIX + str(page_num))
                        ship_month)
    c.drawCentredString(94,
                        page_size[1] - 104,
                        # PAGE_PREFIX + str(page_num))
                        ship_day)

    c.showPage()


def get_page_size(page: PyPDF2.pdf.PageObject) -> tuple:
    """
    既存PDFからページサイズ（幅, 高さ）を取得する
    """
    page_box = page.mediaBox
    width = page_box.getUpperRight_x() - page_box.getLowerLeft_x()
    height = page_box.getUpperRight_y() - page_box.getLowerLeft_y()
    return float(width), float(height)

def change_words(words, func):
    for i, word in enumerate(words):
        words[i] = func(word)

def confirm_day(shipping_day):
    business_day_count = 0
    target_day = shipping_day
    while business_day_count < 4:
        target_day = target_day - datetime.timedelta(days=1)
        if target_day.weekday() != 5 and target_day.weekday() != 6 and target_day.strftime("%Y/%m/%d") not in holiday:
            business_day_count += 1
    return target_day

def trim_end_of_word(trim_string):
    i = 1
    maxcount = len(trim_string)
    while i < maxcount:
        if trim_string[-1] == ' ':
            trim_string = trim_string[:-1]
        else:
            break
        i += 1
    return trim_string

def append_number(X_customer_num, Y_customer_num, X_kanri_num, Y_kanri_num, X_genba_num, Y_genba_num, title_name):
    number_required_for.clear()
    for element in elements:
        if title_name == 'ユニットバスルーム納期最終確認票' or title_name == '御 見 積 書':
            #店コード
            if X_customer_num == math.floor(element['x0']) and Y_customer_num == math.floor(element['y0']):
                number_required_for['店コード'] = element['word'].replace('御 見 積 書\n','')[0:5]
            #管理ナンバー
            elif X_kanri_num == math.floor(element['x0']) and Y_kanri_num == math.floor(element['y0']):
                number_required_for['管理ナンバー'] = element['word']
            #現場名
            elif X_genba_num == math.floor(element['x0']) and Y_genba_num == math.floor(element['y0']):
                number_required_for['現場名'] = trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word']))
            #確定日
            elif 158 == math.floor(element['x0']) and 512 == math.floor(element['y0']) and element['word'].split(' ')[0].split('月')[0] != ''and element['word'] != '発送':
                shipping_day = datetime.date(datetime.datetime.now().date().year, int(element['word'].split(' ')[0].split('月')[0]), int(element['word'].split(' ')[0].split('月')[1][:-1]))
                if datetime.datetime.now().date() > shipping_day:
                    shipping_day = shipping_day + dateutil.relativedelta.relativedelta(years=1)
                confirm_day = compute_confirm_day(shipping_day)
                if confirm_day is None:
                    number_required_for['確定日'] = ''
                else:
                    number_required_for['確定日'] = confirm_day.strftime("%Y/%m/%d")
            #工事区分
            elif 334 == math.floor(element['x0']) and 582 == math.floor(element['y0']):
                if '工事区分' in element['word']:
                    number_required_for['LTS'] = '※'
        elif title_name == 'ユニットバスルームご発注確認票' or title_name == '仕様変更確認票':
            if X_customer_num == math.floor(element['x0']) and Y_customer_num == math.floor(element['y0']):
                if 2 < len(element['word'].split('\n')):
                    number_required_for['現場名'] = trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word'].split('\n')[0]))
                else:
                    number_required_for['店コード'] = element['word'][-5:]
            #管理ナンバー
            elif X_kanri_num == math.floor(element['x0']) and Y_kanri_num == math.floor(element['y0']):
                number_required_for['管理ナンバー'] = element['word']
            #現場名
            elif X_genba_num == math.floor(element['x0']) and Y_genba_num == math.floor(element['y0']) or 45 == math.floor(element['x0']) and 780 == math.floor(element['y0']):
                number_required_for['現場名'] = trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word'].split('\n')[0]).replace('\u3000', '　'))

def rename_file(file_name, target_folder_name, rename_string):
    if not rename_string in file_name:
        duplicated_count = len(glob.glob("./☆ファイル/" + rename_string + "*.pdf"))
        # print(number_required_for['出荷日'])
        if duplicated_count == 0:
            new_rename_string = target_folder_name + rename_string  + '.pdf'
        else:
            new_rename_string = target_folder_name + rename_string + '(' + str(duplicated_count + 1) + ')' + '.pdf'
        if '確定日' in number_required_for:
            if number_required_for['確定日'] != '':
                add_page_number(file_name, new_rename_string)
        else:
            os.rename(file_name, new_rename_string)

def sort_file(file_name, target_folder_name, title_name):
    rename_string = ''
    confirm_day = ''
    lts_construction = ''
    if 'LTS' in number_required_for:
        if number_required_for['LTS'] != '':
            lts_construction = '※'
    if '確定日' in number_required_for:
        if number_required_for['確定日'] != '':
            confirm_day = number_required_for['確定日'].split('/')[1] + '-' + number_required_for['確定日'].split('/')[2]
    if title_name == 'ユニットバスルーム納期最終確認票':
        if '店コード' in number_required_for and '管理ナンバー' in number_required_for and '現場名' in number_required_for:
            rename_string = '【' + confirm_day + '】' + lts_construction + number_required_for['店コード'] + ' ' + number_required_for['管理ナンバー'] + ' ' + number_required_for['現場名']
            rename_file(file_name, target_folder_name, rename_string)
    elif title_name == '御 見 積 書':
        if '店コード' in number_required_for and '現場名' in number_required_for:
            rename_string = number_required_for['店コード'] + ' ' + number_required_for['現場名']
            rename_file(file_name, target_folder_name, rename_string)
    elif title_name == 'ユニットバスルームご発注確認票':
        if '店コード' in number_required_for and '管理ナンバー' in number_required_for and '現場名' in number_required_for:
            rename_string = number_required_for['店コード'] + ' ' + number_required_for['管理ナンバー'] + ' ' + number_required_for['現場名']
            rename_file(file_name, target_folder_name, rename_string)
    elif title_name == '仕様変更確認票':
        if '店コード' in number_required_for and '管理ナンバー' in number_required_for and '現場名' in number_required_for:
            rename_string = '(変)' + number_required_for['店コード'] + ' ' + number_required_for['管理ナンバー'] + ' ' + number_required_for['現場名']
            rename_file(file_name, target_folder_name, rename_string)

def compute_confirm_day(shipping_day):
    if len(holiday) == 0:
        return
    business_day_count = 0
    target_day = shipping_day
    while business_day_count < 4:
        target_day = target_day - datetime.timedelta(days=1)
        if target_day.weekday() != 5 and target_day.weekday() != 6 and target_day.strftime("%Y/%m/%d") not in holiday:
            business_day_count += 1
    return target_day

files = glob.glob("./☆ファイル/*.pdf")

number_required_for = {}
elements = []
holiday = []
# delete_files = []
title_name = ''
target_folder_name = './☆ファイル/'


if os.path.exists("./休日.csv"):
    with open('休日.csv','r', encoding="utf-8") as f:
        reader = csv.reader(f)
        holiday = [rows[0] for rows in reader]
    holiday = holiday[1:]
    change_words(holiday, lambda word: datetime.datetime.strptime(word, "%Y/%m/%d").strftime("%Y/%m/%d"))

for x in files:
    resourceManager = pdfminer.pdfinterp.PDFResourceManager()
    device = pdfminer.converter.PDFPageAggregator(resourceManager, laparams=pdfminer.layout.LAParams())
    elements = []
#     file_name = './☆ファイル/' + 'tes' + str(k) + '.pdf'
    file_name = x
    with open(file_name, 'rb') as fp:
        interpreter = pdfminer.pdfinterp.PDFPageInterpreter(resourceManager, device)
        for i, page in enumerate(pdfminer.pdfpage.PDFPage.get_pages(fp)):
            if i == 0:
                interpreter.process_page(page)
                layout = device.get_result()
                for lt in layout:
                    if isinstance(lt, pdfminer.layout.LTTextContainer):
        #                 print('{}, x0={:.2f}, x1={:.2f}, y0={:.2f}, y1={:.2f}, width={:.2f}, height={:.2f}'.format(
        #                         lt.get_text().strip(), lt.x0, lt.x1, lt.y0, lt.y1, lt.width, lt.height))
                        element = {
                            'word': lt.get_text().strip(),'x0': lt.x0, 'x1': lt.x1, 'y0': lt.y0, 'y1': lt.y1
                        }
                        elements.append(element)
    if len(elements) != 0:
        if 'ユニットバスルーム納期最終確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[2]['word']:
            title_name = 'ユニットバスルーム納期最終確認票'
        #     element, X_customer_num, Y_customer_num, X_kanri_num, Y_kanri_num, X_genba_num, Y_genba_num
            append_number(161, 686, 198, 562, 198, 662, title_name)
            sort_file(file_name, target_folder_name, title_name)
        elif '御 見 積 書' in sorted(elements, key=lambda x: x['y1'], reverse=True)[0]['word']:
            title_name = '御 見 積 書'
        #     element, X_customer_num, Y_customer_num, X_kanri_num, Y_kanri_num, X_genba_num, Y_genba_num
            append_number(19, 772, 0, 0, 67, 744, title_name)
            sort_file(file_name, target_folder_name, title_name)
        elif 'ユニットバスルームご発注確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[1]['word'] or '仕様変更確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[1]['word']:
            if 'ユニットバスルームご発注確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[1]['word']:
                title_name = 'ユニットバスルームご発注確認票'
            else:
                title_name = '仕様変更確認票'
        #     element, X_customer_num, Y_customer_num, X_kanri_num, Y_kanri_num, X_genba_num, Y_genba_num
            append_number(45, 759, 88, 803, 45, 782, title_name)
            sort_file(file_name, target_folder_name, title_name)
    device.close()
