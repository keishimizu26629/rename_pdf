import os
import tkinter
from tkinter import ttk, messagebox, filedialog
import glob, re, math, copy, csv, datetime, io, PyPDF2
import pdfminer.pdfinterp, pdfminer.converter, pdfminer.pdfpage, pdfminer.layout
import dateutil.relativedelta
import reportlab.pdfgen, reportlab.pdfbase, reportlab.pdfbase.cidfonts, reportlab.pdfbase.ttfonts
import reportlab.lib.colors
import traceback

# 継承用のClassの作成
class Sheet():
    def __init__(self, title_name, file_name):
        self.title_name = title_name
        self.file_name = file_name
        self.required_for_processing = []
        self.single_required_for = {}

    # 単語の末尾から空白を取り除く
    def trim_end_of_word(self, trim_string):
        i = 1
        maxcount = len(trim_string)
        while i < maxcount:
            if trim_string[-1] == ' ':
                trim_string = trim_string[:-1]
            else:
                break
            i += 1
        return trim_string

    # 見積書の情報を抽出する
    def extraction_for_quotation(self, single_required_for, customer_X_coordinate, customer_Y_coordinate, property_X_coordinate, property_Y_coordinate, elements):
        for element in elements:
            # 店コードの抽出
            if customer_X_coordinate == math.floor(element['x0']) and customer_Y_coordinate == math.floor(element['y0']):
                single_required_for['店コード'] = element['word'].replace('御 見 積 書\n', '')[0:5]
            # 現場名の抽出
            elif property_X_coordinate == math.floor(element['x0']) and property_Y_coordinate == math.floor(element['y0']):
                single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+', '_', element['word']))

    # 確定日を計算する
    def calculate_confirm_day(self, shipping_day):
        if len(holiday) == 0: return
        business_day_count = 0
        while business_day_count < 4:
            shipping_day = shipping_day - datetime.timedelta(days=1)
            if shipping_day.weekday() != 5 and shipping_day.weekday() != 6 and shipping_day.strftime("%Y/%m/%d") not in holiday:
                business_day_count += 1
        return shipping_day

    # 処理に必要な情報を追加する（サブクラスで実装）
    def append_required_for_processing(self):
        pass

    # 確定日をPDFに書き込む
    def write_confirm_day(self):
        fi = open(self.file_name, 'rb')
        pdf_reader = PyPDF2.PdfFileReader(fi)
        pages_num = pdf_reader.getNumPages()
        # ページ番号を付けたPDFの書き込み用
        pdf_writer = PyPDF2.PdfFileWriter()
        # ページ番号だけのPDFをメモリ（binary stream）に作成
        bs = io.BytesIO()
        c = reportlab.pdfgen.canvas.Canvas(bs)
        for i in range(0, pages_num):
            # 既存PDF
            pdf_page = pdf_reader.getPage(i)
            # PDFページのサイズ
            page_size = self.get_page_size(pdf_page)
            # ページ番号のPDF作成
            self.create_page_number_pdf(c, page_size, i)
        c.save()
        # ページ番号だけのPDFをメモリから読み込み（seek操作はPyPDF2に実装されているので不要）
        pdf_num_reader = PyPDF2.PdfFileReader(bs)
        # 既存PDFに１ページずつ処理する
        for i in range(0, pages_num):
            # 既存PDF
            pdf_page = pdf_reader.getPage(i)
            # ページ番号だけのPDF
            pdf_num = pdf_num_reader.getPage(i)
            # ２つのPDFを重ねる
            pdf_page.mergePage(pdf_num)
            pdf_writer.addPage(pdf_page)

        # ページ番号を付けたPDFを保存と元ファイル削除
        fo = open(self.new_rename_string, 'wb')
        pdf_writer.write(fo)
        bs.close()
        fi.close()
        fo.close()
        os.remove(self.file_name)

    # ページ番号だけのPDFを作成する
    def create_page_number_pdf(self, c: reportlab.pdfgen.canvas.Canvas, page_size: tuple, i):
        # 確定日の印字
        ship_month = self.required_for_processing[i]['確定日'].split('/')[1]
        if str(ship_month[0]) == '0': ship_month = ship_month[1]
        ship_day = self.required_for_processing[i]['確定日'].split('/')[2]
        if str(ship_day[0]) == '0': ship_day = ship_day[1]
        c.setPageSize(page_size)
        c.setFont('MS P ゴシック', 16)
        c.drawCentredString(254, page_size[1] - 130, ship_month)
        c.drawCentredString(296, page_size[1] - 130, ship_day)

        # 札幌DC対応の場合、文言を追加
        if self.single_required_for.get('札幌', False):
            c.setFont("MS P ゴシック", 12)
            text_object = c.beginText(20, 90)
            for line in self.sapporo_text.split('\n'):
                text_object.textLine(line)
            c.drawText(text_object)
        c.showPage()

    # 既存PDFからページサイズ（幅, 高さ）を取得する
    def get_page_size(self, page: PyPDF2.pdf.PageObject) -> tuple:
        page_box = page.mediaBox
        width = page_box.getUpperRight_x() - page_box.getLowerLeft_x()
        height = page_box.getUpperRight_y() - page_box.getLowerLeft_y()
        return float(width), float(height)

    # ファイル名を変更する
    def rename_file(self, rename_string, target_folder_name):
        if not rename_string in self.file_name:
            duplicated_count = len(glob.glob("./☆ファイル/" + rename_string + "*.pdf"))
            if duplicated_count == 0:
                self.new_rename_string = target_folder_name + rename_string + '.pdf'
            else:
                self.new_rename_string = target_folder_name + rename_string + '(' + str(duplicated_count + 1) + ')' + '.pdf'
            if '確定日' in self.required_for_processing[0]:
                if self.required_for_processing[0]['確定日'] != '':
                    self.write_confirm_day()
                else:
                    os.rename(self.file_name, self.new_rename_string)
            else:
                os.rename(self.file_name, self.new_rename_string)
            print(rename_string)

    def extract_file_name(self):
        pass

# 最終確認票分のClass
class Final_check_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)
        self.sapporo_text = (
                    "キャンセル：出荷日８日前の午前中まで\n"
                    "仕様変更：出荷日５日前の午前中までにご依頼お願いします。\n"
                    "内容によってお受けできない場合があります。\n"
                    "※ 長納期品、基準外品等に関しては都度ご確認をお願いします。")

    # 必要な情報を追加する
    def append_required_for_processing(self, elements):
        for element in elements:
            if 161 == math.floor(element['x0']) and 650 == math.floor(element['y0']):
                self.single_required_for['店コード'] = element['word'][0:5]
            elif 198 == math.floor(element['x0']) and 523 == math.floor(element['y0']):
                self.single_required_for['管理ナンバー'] = element['word'].split('\n')[2]
            elif 198 == math.floor(element['x0']) and 625 == math.floor(element['y0']):
                self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+', '_', element['word']))
            elif 158 == math.floor(element['x0']) and 472 == math.floor(element['y0']) and element['word'].split(' ')[0].split('月')[0] != '' and element['word'] != '発送':
                shipping_day = datetime.date(datetime.datetime.now().date().year, int(element['word'].split(' ')[0].split('月')[0]), int(element['word'].split(' ')[0].split('月')[1][:-1]))
                if datetime.datetime.now().date() > shipping_day:
                    shipping_day = shipping_day + dateutil.relativedelta.relativedelta(years=1)
                confirm_day = self.calculate_confirm_day(shipping_day)
                if confirm_day is None:
                    self.single_required_for['確定日'] = ''
                else:
                    self.single_required_for['確定日'] = confirm_day.strftime("%Y/%m/%d")
            elif 334 == math.floor(element['x0']) and 523 == math.floor(element['y0']):
                if '工事区分' in element['word']:
                    self.single_required_for['LTS'] = '※'

            # 札幌DC対応分の住所が北海道・札幌から始まる場合のみの納期返信文言表示
            elif 158 == math.floor(element['x0']) and 344 == math.floor(element['y0']):
                if element['word'].startswith('札幌市') or element['word'].startswith('北海道'):
                    self.single_required_for['札幌'] = True
                else:
                    self.single_required_for['札幌'] = False

        self.required_for_processing.append(copy.deepcopy(self.single_required_for))

    # ファイル名を抽出してリネームする
    def extract_file_name(self, target_folder_name):
        confirm_day = ''
        lts_construction = ''
        if 'LTS' in self.required_for_processing[0]:
            if self.required_for_processing[0]['LTS'] != '':
                lts_construction = '※'
        if '確定日' in self.required_for_processing[0]:
            if self.required_for_processing[0]['確定日'] != '':
                confirm_day = self.required_for_processing[0]['確定日'].split('/')[1] + '-' + self.required_for_processing[0]['確定日'].split('/')[2]
        if '店コード' in self.required_for_processing[0] and '管理ナンバー' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            rename_string = '【' + confirm_day + '】' + lts_construction + self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['管理ナンバー'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

# 仕様明細書のClass
class Detail_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)

    # 必要な情報を追加する
    def append_required_for_processing(self, elements):
        for element in elements:
            # 店コード
            if 45 == math.floor(element['x0']) and 759 == math.floor(element['y0']):
                if 2 < len(element['word'].split('\n')):
                    self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+', '_', element['word'].split('\n')[0]))
                else:
                    self.single_required_for['店コード'] = element['word'][-5:]
            # 管理ナンバー
            elif 88 == math.floor(element['x0']) and 803 == math.floor(element['y0']):
                self.single_required_for['管理ナンバー'] = element['word']
            # 現場名
            elif 45 == math.floor(element['x0']) and 782 == math.floor(element['y0']) or 45 == math.floor(element['x0']) and 780 == math.floor(element['y0']):
                self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+', '_', element['word'].split('\n')[0]).replace('\u3000', '　'))
        self.required_for_processing.append(copy.deepcopy(self.single_required_for))

    # ファイル名を抽出してリネームする
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '管理ナンバー' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            rename_string = self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['管理ナンバー'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

# 見積書のClass
class Quotation_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)

    # 必要な情報を追加する
    def append_required_for_processing(self, elements):
        self.extraction_for_quotation(self.single_required_for, 19, 772, 67, 744, elements)
        if len(self.required_for_processing) == 0:
            self.extraction_for_quotation(self.single_required_for, 19, 783, 67, 719, elements)
        elif '店コード' not in self.required_for_processing[0]:
            self.extraction_for_quotation(self.single_required_for, 19, 783, 67, 719, elements)
        self.required_for_processing.append(copy.deepcopy(self.single_required_for))

    # ファイル名を抽出してリネームする
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            rename_string = self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

# 仕様変更の仕様明細書
class Change_specifications_sheet(Detail_sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)

    # ファイル名を抽出してリネームする
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '管理ナンバー' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            rename_string = '(変)' + self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['管理ナンバー'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

# キャンセル明細書
class Cancel_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)

    # 必要な情報を追加する
    def append_required_for_processing(self, elements):
        for element in elements:
            # 店コード
            if 107 == math.floor(element['x0']) and 722 == math.floor(element['y0']):
                self.single_required_for['店コード'] = element['word']
            # 管理ナンバー
            elif 79 == math.floor(element['x0']) and 786 == math.floor(element['y0']):
                self.single_required_for['管理ナンバー'] = element['word']
            # 現場名
            elif 107 == math.floor(element['x0']) and 756 == math.floor(element['y0']):
                self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+', '_', element['word']))
            # 札幌DC分は警告の文字列を表示する
            elif 107 == math.floor(element['x0']) and 586 == math.floor(element['y0']):
                if 'Fｼﾞﾄﾞｳｼﾖﾘ' == element['word']:
                    self.single_required_for['札幌'] = True
                else:
                    self.single_required_for['札幌'] = False
        self.required_for_processing.append(copy.deepcopy(self.single_required_for))

    # ファイル名を抽出してリネームする
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '管理ナンバー' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            cancel_string = ''
            if self.required_for_processing[0]['札幌'] == True:
                cancel_string = '(キャンセル不可！！)'
            else:
                cancel_string = '(消)'
            rename_string = cancel_string + self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['管理ナンバー'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

# リスト内の単語を変換する
def change_words(words, func):
    for i, word in enumerate(words):
        words[i] = func(word)

# PDFファイルを結合する
def merge_files_for_posting(processed_files, target_folder_name):
    for file_being_processed in processed_files:
        if file_being_processed.title_name == 'ユニットバスルーム納期最終確認票':
            merge_file = file_being_processed.new_rename_string.replace(target_folder_name, '')
            print(f"Processing file: {merge_file}")

            # 対応するファイルを探す
            base_filename = merge_file.split('】', 1)[1]  # 【日付】の部分を除去
            matching_file = base_filename

            if os.path.exists(os.path.join(target_folder_name, matching_file)):
                pdf_file_merger = PyPDF2.PdfFileMerger()

                # 元のファイルを追加
                pdf_file_merger.append(os.path.join(target_folder_name, merge_file))

                # 対応するファイルを追加
                pdf_file_merger.append(os.path.join(target_folder_name, matching_file))

                # 新しいファイル名を生成（先頭に(投函用)を追加）
                output_filename = os.path.join(target_folder_name, '(投函用)' + merge_file)

                # 結合したPDFを保存
                pdf_file_merger.write(output_filename)
                pdf_file_merger.close()

                print(f"Merged files: {merge_file} and {matching_file}")
                print(f"Output file created: {output_filename}")
            else:
                print(f"Matching file not found for: {merge_file}")

    print("PDF merging process completed.")

# メイン処理
def main_process(elements, target_folder_name):
    files_to_process = glob.glob(target_folder_name + "*.pdf")
    processed_count = 0
    for j, file_being_processed in enumerate(files_to_process):
        resourceManager = pdfminer.pdfinterp.PDFResourceManager()
        device = pdfminer.converter.PDFPageAggregator(resourceManager, laparams=pdfminer.layout.LAParams())
        with open(file_being_processed, 'rb') as fp:
            interpreter = pdfminer.pdfinterp.PDFPageInterpreter(resourceManager, device)
            do_first_processing = False
            for i, page in enumerate(pdfminer.pdfpage.PDFPage.get_pages(fp)):
                interpreter.process_page(page)
                layout = device.get_result()
                elements = []
                for lt in layout:
                    if isinstance(lt, pdfminer.layout.LTTextContainer):
                        element = {
                            'word': lt.get_text().strip(), 'x0': lt.x0, 'x1': lt.x1, 'y0': lt.y0, 'y1': lt.y1
                        }
                        elements.append(element)
                if len(elements) != 0:
                    if 'ユニットバスルーム納期最終確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[2]['word'] or 'ユニットバスルーム納期最終確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[1]['word']:
                        if do_first_processing == False:
                            processed_count += 1
                            file_instance = Final_check_sheet('ユニットバスルーム納期最終確認票', file_being_processed)
                        file_instance.append_required_for_processing(elements)
                        do_first_processing = True
                    elif '御 見 積 書' in sorted(elements, key=lambda x: x['y1'], reverse=True)[0]['word']:
                        if do_first_processing == False:
                            file_instance = Quotation_sheet('御 見 積 書', file_being_processed)
                            file_instance.append_required_for_processing(elements)
                            processed_count += 1
                            do_first_processing = True
                    elif 'ユニットバスルームご発注確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[1]['word']:
                        if do_first_processing == False:
                            file_instance = Detail_sheet('ユニットバスルームご発注確認票', file_being_processed)
                            file_instance.append_required_for_processing(elements)
                            processed_count += 1
                            do_first_processing = True
                    elif '仕様変更確認票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[1]['word']:
                        if do_first_processing == False:
                            file_instance = Change_specifications_sheet('仕様変更確認票', file_being_processed)
                            file_instance.append_required_for_processing(elements)
                            processed_count += 1
                            do_first_processing = True
                    elif 'キ ャ ン セ ル 確 認 票' in sorted(elements, key=lambda x: x['y1'], reverse=True)[0]['word']:
                        if do_first_processing == False:
                            file_instance = Cancel_sheet('キャンセル確認票', file_being_processed)
                            file_instance.append_required_for_processing(elements)
                            processed_count += 1
                            do_first_processing = True
        if do_first_processing == True:
            if len(file_instance.required_for_processing) != 0:
                file_instance.extract_file_name(target_folder_name)
                if 'new_rename_string' in vars(file_instance):
                    processed_files.append(copy.deepcopy(file_instance))
                del file_instance
        device.close()

# 既存PDFからページサイズ（幅, 高さ）を取得する
def get_page_size(page: PyPDF2.pdf.PageObject) -> tuple:
    page_box = page.mediaBox
    width = page_box.getUpperRight_x() - page_box.getLowerLeft_x()
    height = page_box.getUpperRight_y() - page_box.getLowerLeft_y()
    return float(width), float(height)

# フォルダダイアログを開く
def dirdialog_clicked():
    dirPath = entry1.get()
    if dirPath:
        iDir = dirPath
    else:
        iDir = 'C:\\'
    iDirPath = filedialog.askdirectory(initialdir=iDir)
    if not iDirPath == '':
        entry1.set(iDirPath)
        if not os.path.exists('./var'):
            os.makedirs('./var')
        with open('./var/path.txt', 'w', encoding="utf-8") as f:
            f.write(iDirPath)

# メイン処理の実行
def conductMain(elements, processed_files):
    text = ""
    dirPath = entry1.get()
    if dirPath:
        target_folder_name = dirPath + '/'
        main_process(elements, target_folder_name)
        merge_files_for_posting(processed_files, target_folder_name)
    # else:
    #     text = "フォルダを指定してください！"
    #     messagebox.showinfo("info", text)

class ErrorDialog(tkinter.Toplevel):
    def __init__(self, parent, error_message):
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

def show_error(parent, error_message):
    ErrorDialog(parent, error_message)

def error_handler(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            error_message = f"エラーの種類: {type(e).__name__}\n"
            error_message += f"エラーメッセージ: {str(e)}\n\n"
            error_message += "詳細なエラー情報:\n"
            error_message += traceback.format_exc()
            show_error(root, error_message)
    return wrapper

@error_handler
def conductMain(elements, processed_files):
    text = ""
    dirPath = entry1.get()
    if dirPath:
        target_folder_name = dirPath + '/'
        main_process(elements, target_folder_name)
        merge_files_for_posting(processed_files, target_folder_name)
    else:
        raise ValueError("フォルダを指定してください！")

if __name__ == "__main__":
    reportlab.pdfbase.pdfmetrics.registerFont(reportlab.pdfbase.ttfonts.TTFont('Meiryo UI', 'c:/Windows/Fonts/meiryob.ttc'))
    reportlab.pdfbase.pdfmetrics.registerFont(reportlab.pdfbase.ttfonts.TTFont('MS P ゴシック', 'c:/Windows/Fonts/msgothic.ttc'))
    elements = []
    holiday = []
    processed_files = []

    if os.path.exists("./休日.csv"):
        with open('休日.csv', 'r', encoding="utf-8") as f:
            reader = csv.reader(f)
            holiday = [rows[0] for rows in reader]
        holiday = holiday[1:]
        change_words(holiday, lambda word: datetime.datetime.strptime(word, "%Y/%m/%d").strftime("%Y/%m/%d"))

    # rootの作成
    root = tkinter.Tk()
    root.title("renamePdf")
    root.iconbitmap('icons/icon.ico')

    # Frame1の作成
    frame1 = ttk.Frame(root, padding=10)
    frame1.grid(row=0, column=1, sticky=tkinter.E)

    # 「フォルダ参照」ラベルの作成
    IDirLabel = ttk.Label(frame1, text="フォルダ参照＞＞", padding=(5, 2))
    IDirLabel.pack(side=tkinter.LEFT)

    # 「フォルダ参照」エントリーの作成
    entry1 = tkinter.StringVar()
    IDirEntry = ttk.Entry(frame1, textvariable=entry1, width=30)
    IDirEntry.pack(side=tkinter.LEFT)

    if os.path.exists("./var/path.txt"):
        with open('./var/path.txt', 'r', encoding="utf-8") as f:
            s = f.read()
            entry1.set(s)
            target_folder_name = s + '/'

    # 「フォルダ参照」ボタンの作成
    IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    IDirButton.pack(side=tkinter.LEFT)

    # Frame3の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid(row=5, column=1, sticky=tkinter.W)

    # 実行ボタンの設置
    button1 = ttk.Button(frame3, text="実行", command=lambda: conductMain(elements, processed_files))
    button1.pack(fill="x", padx=30, side="left")

    # キャンセルボタンの設置
    button2 = ttk.Button(frame3, text=("閉じる"), command=root.destroy)
    button2.pack(fill="x", padx=30, side="left")

    root.geometry("400x130+200+300")
    root.mainloop()
