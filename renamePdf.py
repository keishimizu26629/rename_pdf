import os,sys
import tkinter
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import os, glob, re, math, copy, csv, datetime, io, PyPDF2
import pdfminer.pdfinterp, pdfminer.converter, pdfminer.pdfpage, pdfminer.layout
import dateutil.relativedelta
# from PyPDF2.pdf import PageObject
import reportlab.pdfgen, reportlab.pdfbase, reportlab.pdfbase.cidfonts, reportlab.pdfbase.ttfonts

class Sheet():
    def __init__(self, title_name, file_name):
        self.title_name = title_name
        self.file_name = file_name
        self.required_for_processing = []
        self.single_required_for = {}
    def trim_end_of_word(self,trim_string):
        i = 1
        maxcount = len(trim_string)
        while i < maxcount:
            if trim_string[-1] == ' ':
                trim_string = trim_string[:-1]
            else:
                break
            i += 1
        return trim_string
    def extraction_for_quotation(self, single_required_for, customer_X_coordinate, customer_Y_coordinate, property_X_coordinate, property_Y_coordinate, elements):
        for element in elements:
            #店コード
            if customer_X_coordinate == math.floor(element['x0']) and customer_Y_coordinate == math.floor(element['y0']):
                single_required_for['店コード'] = element['word'].replace('御 見 積 書\n','')[0:5]
            #現場名
            elif property_X_coordinate == math.floor(element['x0']) and property_Y_coordinate == math.floor(element['y0']):
                single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word']))
    def calculate_confirm_day(self, shipping_day):
        if len(holiday) == 0: return
        business_day_count = 0
        while business_day_count < 4:
            shipping_day = shipping_day - datetime.timedelta(days=1)
            if shipping_day.weekday() != 5 and shipping_day.weekday() != 6 and shipping_day.strftime("%Y/%m/%d") not in holiday:
                business_day_count += 1
        return shipping_day
    def append_required_for_processing(self):
        pass
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
        # 既存PDFに１ページずつページ番号を付ける
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

    def create_page_number_pdf(self, c: reportlab.pdfgen.canvas.Canvas, page_size: tuple, i):
        """
        ページ番号だけのPDFを作成
        """
        # 確定日の印字
        ship_month = self.required_for_processing[i]['確定日'].split('/')[1]
        if str(ship_month[0]) == '0': ship_month = ship_month[1]
        ship_day = self.required_for_processing[i]['確定日'].split('/')[2]
        if str(ship_day[0]) == '0': ship_day = ship_day[1]
        c.setPageSize(page_size)
        c.setFont('MS P ゴシック', 16)
        # c.drawCentredString(50, page_size[1] - 104, ship_month)
        # c.drawCentredString(94, page_size[1] - 104, ship_day)
        c.drawCentredString(254, page_size[1] - 130, ship_month)
        c.drawCentredString(296, page_size[1] - 130, ship_day)
        c.showPage()
    def get_page_size(self, page: PyPDF2.pdf.PageObject) -> tuple:
        """
        既存PDFからページサイズ（幅, 高さ）を取得する
        """
        page_box = page.mediaBox
        width = page_box.getUpperRight_x() - page_box.getLowerLeft_x()
        height = page_box.getUpperRight_y() - page_box.getLowerLeft_y()
        return float(width), float(height)
    def rename_file(self, rename_string, target_folder_name):
        if not rename_string in self.file_name:
            duplicated_count = len(glob.glob("./☆ファイル/" + rename_string + "*.pdf"))
            # print(self.required_for_processing['出荷日'])
            if duplicated_count == 0:
                self.new_rename_string = target_folder_name + rename_string  + '.pdf'
            else:
                self.new_rename_string = target_folder_name + rename_string + '(' + str(duplicated_count + 1) + ')' + '.pdf'
            if '確定日' in self.required_for_processing[0]:
                if self.required_for_processing[0]['確定日'] != '':
                    self.write_confirm_day()
                else:
                    os.rename(self.file_name, self.new_rename_string)
            else:
                os.rename(self.file_name, self.new_rename_string)
            # if '【' == new_rename_string.replace(target_folder_name, '')[0]:
            #     marge_files_list.append(new_rename_string.replace(target_folder_name, ''))
    def extract_file_name(self):
        pass

class Final_check_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)
    def append_required_for_processing(self, elements):
        for element in elements:
            #店コード
            # if 161 == math.floor(element['x0']) and 686 == math.floor(element['y0']):
            if 161 == math.floor(element['x0']) and 650 == math.floor(element['y0']):
                #店名が長くなった時のイレギュラー対応
                # if 719 == math.floor(element['y1']):
                #     self.single_required_for['店コード'] = element['word'].split('\n')[1][0:5]
                # else:
                #     self.single_required_for['店コード'] = element['word'].replace('御 見 積 書\n','')[0:5]
                self.single_required_for['店コード'] = element['word'][0:5]
            #管理ナンバー
            # elif 198 == math.floor(element['x0']) and 562 == math.floor(element['y0']):
            #     self.single_required_for['管理ナンバー'] = element['word']
            elif 198 == math.floor(element['x0']) and 523 == math.floor(element['y0']):
                self.single_required_for['管理ナンバー'] = element['word'].split('\n')[2]
            #現場名
            # elif 198 == math.floor(element['x0']) and 662 == math.floor(element['y0']):
            #     self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word']))
            elif 198 == math.floor(element['x0']) and 625 == math.floor(element['y0']):
                self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word']))
            #確定日
            # elif 158 == math.floor(element['x0']) and 512 == math.floor(element['y0']) and element['word'].split(' ')[0].split('月')[0] != ''and element['word'] != '発送':
            elif 158 == math.floor(element['x0']) and 472 == math.floor(element['y0']) and element['word'].split(' ')[0].split('月')[0] != ''and element['word'] != '発送':
                shipping_day = datetime.date(datetime.datetime.now().date().year, int(element['word'].split(' ')[0].split('月')[0]), int(element['word'].split(' ')[0].split('月')[1][:-1]))
                if datetime.datetime.now().date() > shipping_day:
                    shipping_day = shipping_day + dateutil.relativedelta.relativedelta(years=1)
                confirm_day = self.calculate_confirm_day(shipping_day)
                if confirm_day is None:
                    self.single_required_for['確定日'] = ''
                else:
                    self.single_required_for['確定日'] = confirm_day.strftime("%Y/%m/%d")
            #工事区分
            # elif 334 == math.floor(element['x0']) and 582 == math.floor(element['y0']):
            elif 334 == math.floor(element['x0']) and 523 == math.floor(element['y0']):
                if '工事区分' in element['word']:
                    self.single_required_for['LTS'] = '※'
        self.required_for_processing.append(copy.deepcopy(self.single_required_for))
        # print(self.single_required_for)
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

class Detail_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)
    def append_required_for_processing(self, elements):
        for element in elements:
            #店コード
            if 45 == math.floor(element['x0']) and 759 == math.floor(element['y0']):
                if 2 < len(element['word'].split('\n')):
                    self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word'].split('\n')[0]))
                else:
                    self.single_required_for['店コード'] = element['word'][-5:]
            #管理ナンバー
            elif 88 == math.floor(element['x0']) and 803 == math.floor(element['y0']):
                self.single_required_for['管理ナンバー'] = element['word']
            #現場名
            elif 45 == math.floor(element['x0']) and 782 == math.floor(element['y0']) or 45 == math.floor(element['x0']) and 780 == math.floor(element['y0']):
                self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word'].split('\n')[0]).replace('\u3000', '　'))
        self.required_for_processing.append(copy.deepcopy(self.single_required_for))
        # print(self.single_required_for)
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '管理ナンバー' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            rename_string = self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['管理ナンバー'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

class Quotation_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)
    def append_required_for_processing(self, elements):
        self.extraction_for_quotation(self.single_required_for, 19, 772, 67, 744, elements)
        if len(self.required_for_processing) == 0:
            self.extraction_for_quotation(self.single_required_for,19, 783, 67, 719, elements)
        elif not '店コード' in self.required_for_processing[0]:
            self.extraction_for_quotation(self.single_required_for,19, 783, 67, 719, elements)
        self.required_for_processing.append(copy.deepcopy(self.single_required_for))
        # print(self.single_required_for)
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            rename_string = self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

class Change_specifications_sheet(Detail_sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '管理ナンバー' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
                rename_string = '(変)' + self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['管理ナンバー'] + ' ' + self.required_for_processing[0]['現場名']
                # print(rename_string)
                self.rename_file(rename_string, target_folder_name)

class Cancel_sheet(Sheet):
    def __init__(self, title_name, file_name):
        super().__init__(title_name, file_name)
    def append_required_for_processing(self, elements):
        for element in elements:
            #店コード
            if 107 == math.floor(element['x0']) and 722 == math.floor(element['y0']):
                self.single_required_for['店コード'] = element['word']
            #管理ナンバー
            elif 79 == math.floor(element['x0']) and 786 == math.floor(element['y0']):
                self.single_required_for['管理ナンバー'] = element['word']
            #現場名
            elif 107 == math.floor(element['x0']) and 756 == math.floor(element['y0']):
                self.single_required_for['現場名'] = self.trim_end_of_word(re.sub(r'[\\/:*?"<>|]+','_',element['word']))
        self.required_for_processing.append(copy.deepcopy(self.single_required_for))
    def extract_file_name(self, target_folder_name):
        if '店コード' in self.required_for_processing[0] and '管理ナンバー' in self.required_for_processing[0] and '現場名' in self.required_for_processing[0]:
            rename_string = '(消)' + self.required_for_processing[0]['店コード'] + ' ' + self.required_for_processing[0]['管理ナンバー'] + ' ' + self.required_for_processing[0]['現場名']
            self.rename_file(rename_string, target_folder_name)

def change_words(words, func):
    for i, word in enumerate(words):
        words[i] = func(word)

def merge_files_for_posting(processed_files, target_folder_name):
    for file_being_processed in processed_files:
        if file_being_processed.title_name == 'ユニットバスルーム納期最終確認票' and not 'LTS' in file_being_processed.required_for_processing[0]:
            merge_file = file_being_processed.new_rename_string.replace(target_folder_name,'')
            matched_file = list(filter(lambda x: x[0] != '【', list(map(lambda x: x.replace('./☆ファイル\\', ''), glob.glob(target_folder_name + '*' + merge_file.replace('※', '')[13:20] + '*')))))
            if len(matched_file) != 0:
                pdf_file_merger = PyPDF2.PdfFileMerger()
                pdf_file_merger.append(target_folder_name + merge_file)
                pdf_file_merger.append(target_folder_name + matched_file[0])
                pdf_file_merger.write(target_folder_name + '(投函用)' + merge_file)
                pdf_file_merger.close()

def main_process(elements, target_folder_name):
    files_to_process = glob.glob(target_folder_name + "*.pdf")
    processed_count = 0
    for j, file_being_processed in enumerate(files_to_process):
        resourceManager = pdfminer.pdfinterp.PDFResourceManager()
        device = pdfminer.converter.PDFPageAggregator(resourceManager, laparams=pdfminer.layout.LAParams())
    #     file_name = './☆ファイル/' + 'tes' + str(k) + '.pdf'
        with open(file_being_processed, 'rb') as fp:
            interpreter = pdfminer.pdfinterp.PDFPageInterpreter(resourceManager, device)
            do_first_processing = False
            for i, page in enumerate(pdfminer.pdfpage.PDFPage.get_pages(fp)):
                interpreter.process_page(page)
                layout = device.get_result()
                elements = []
                for lt in layout:
                    if isinstance(lt, pdfminer.layout.LTTextContainer):
                        #                 print('{}, x0={:.2f}, x1={:.2f}, y0={:.2f}, y1={:.2f}, width={:.2f}, height={:.2f}'.format(
                        #                         lt.get_text().strip(), lt.x0, lt.x1, lt.y0, lt.y1, lt.width, lt.height))
                        element = {
                            'word': lt.get_text().strip(),'x0': lt.x0, 'x1': lt.x1, 'y0': lt.y0, 'y1': lt.y1
                        }
                        elements.append(element)
                if len(elements) != 0:
                    # customer_X_coordinate, customer_Y_coordinate, X_consecutive_num, Y_consecutive_num, property_X_coordinate, property_Y_coordinate, title_name
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
                            file_instance.append_required_for_processing()
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
    messagebox.showinfo("info", '処理が完了しました。')

def dirdialog_clicked():
    # iDir = os.path.abspath(os.path.dirname(__file__))
    dirPath = entry1.get()
    if dirPath:
        iDir = dirPath
    else:
        iDir = 'C:\\'
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    if not iDirPath == '':
        entry1.set(iDirPath)
        if not os.path.exists('./var'):
            os.makedirs('./var')
        with open('./var/path.txt','w', encoding="utf-8") as f:
            f.write(iDirPath)

def conductMain(elements):
    text = ""
    dirPath = entry1.get()
    if dirPath:
        target_folder_name = dirPath + '/'
        main_process(elements, target_folder_name)
    else:
        text = "フォルダを指定してください！"
        messagebox.showinfo("info", text)

# def on_closing():
#     with open('./var/path.txt','w', encoding="utf-8") as f:
#         dirPath = entry1.get()
#         f.write(dirPath)
    # root.destroy()

if __name__ == "__main__":

    reportlab.pdfbase.pdfmetrics.registerFont(reportlab.pdfbase.ttfonts.TTFont('MS P ゴシック', 'c:/Windows/Fonts/msgothic.ttc'))
    elements = []
    holiday = []
    processed_files = []
    # target_folder_name = './☆ファイル/'

    if os.path.exists("./休日.csv"):
        with open('休日.csv','r', encoding="utf-8") as f:
            reader = csv.reader(f)
            holiday = [rows[0] for rows in reader]
        holiday = holiday[1:]
        change_words(holiday, lambda word: datetime.datetime.strptime(word, "%Y/%m/%d").strftime("%Y/%m/%d"))

    # rootの作成
    root = tkinter.Tk()
    root.title("renamePdf")

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
        with open('./var/path.txt','r', encoding="utf-8") as f:
            s = f.read()
            entry1.set(s)
            target_folder_name = s + '/'

    # 「フォルダ参照」ボタンの作成
    IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    IDirButton.pack(side=tkinter.LEFT)

    # Frame3の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid(row=5,column=1,sticky=tkinter.W)

    # 実行ボタンの設置
    button1 = ttk.Button(frame3, text="実行", command=lambda: conductMain(elements))
    button1.pack(fill = "x", padx=30, side = "left")

    # キャンセルボタンの設置
    button2 = ttk.Button(frame3, text=("閉じる"), command=root.destroy)
    button2.pack(fill = "x", padx=30, side = "left")

    # root.protocol("WM_DELETE_WINDOW", root.destroy)

    root.geometry("400x130+200+300")

    root.mainloop()
