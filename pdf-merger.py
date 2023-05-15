import tkinter as tk
import tkinter.filedialog
import pandas as pd
from natsort import natsorted

import PyPDF2 as pp2
import glob
import os

import re
from io import StringIO

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage

class MpdfClass:
  def __init__(self):
    root = tk.Tk()
    root.title('PDF merger D')
    root.geometry("500x680")

    button_1 = tk.Button(root, text='参照するエクセルファイルを選択', font=('', 20),
                        width=29, height=1, bg='#999999', activebackground="#aaaaaa")
    button_1.bind('<ButtonPress>', self.file_dialog_1)
    button_1.pack(pady=40)

    self.file_name_1 = tk.StringVar()
    self.file_name_1.set('未選択です')
    label_1 = tk.Label(textvariable=self.file_name_1, font=('', 12))
    label_1.pack(pady=0)

    button_2 = tk.Button(root, text='参照するフォルダーを選択', font=('', 20),
                        width=24, height=1, bg='#999999', activebackground="#aaaaaa")
    button_2.bind('<ButtonPress>', self.folder_dialog_2)
    button_2.pack(pady=40)

    self.folder_name_2 = tk.StringVar()
    self.folder_name_2.set('未選択です')
    label_2 = tk.Label(textvariable=self.folder_name_2, font=('', 12))
    label_2.pack(pady=10)
    
    button_3 = tk.Button(root, text='出力先フォルダーを選択', font=('', 20),
                width=24, height=1, bg='#999999', activebackground="#aaaaaa")
    button_3.bind('<ButtonPress>', self.folder_dialog_3)
    button_3.pack(pady=40)

    self.folder_name_3 = tk.StringVar()
    self.folder_name_3.set('未選択です')
    label_3 = tk.Label(textvariable=self.folder_name_3, font=('', 12))
    label_3.pack(pady=10)
    
    button_4 = tk.Button(root, text='スタート', font=('', 20),
                width=10, height=1, bg='#999999', activebackground="#aaaaaa")
    button_4.bind('<ButtonPress>', self.merge_pdf_4)
    button_4.pack(pady=40)

    self.status_4 = tk.StringVar()
    self.status_4.set('待機中')
    label_4 = tk.Label(textvariable=self.status_4, font=('', 12))
    label_4.pack(pady=10)

    root.mainloop()

  def file_dialog_1(self, event):
    fTyp = [("xlsx", "xlsx"), ("xls", "xls")]
    
    global file_name_1
    file_name_1 = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir="C:")
    if len(file_name_1) == 0:
      self.file_name_1.set('選択をキャンセルしました')
    else:
      self.file_name_1.set(file_name_1)

  def folder_dialog_2(self, event):
    global folder_name_2
    folder_name_2 = tk.filedialog.askdirectory(initialdir="C:")
    if len(folder_name_2) == 0:
      self.folder_name_2.set('選択をキャンセルしました')
    else:
      self.folder_name_2.set(folder_name_2)

  def folder_dialog_3(self, event):
    global folder_name_3
    folder_name_3 = tk.filedialog.askdirectory(initialdir="C:")
    if len(folder_name_3) == 0:
      self.folder_name_3.set('選択をキャンセルしました')
    else:
      self.folder_name_3.set(folder_name_3)

  def merge_pdf_4(self, event):
    try:
      xlsx_file=file_name_1
      trgt_folder=folder_name_2
      day_file='*.pdf'
      dst_folder=folder_name_3

      self.status_4.set('進行中…')

      df = pd.read_excel(xlsx_file, sheet_name=0, header=None)
      _df = df.iloc[0:df.shape[0]+1, 0]
      _df = _df.to_list()

      search_text_list = [] 
      for i in range(len(_df)):
        str = _df[i]
        search_text_list.append(str)

      files = glob.glob(os.path.join(trgt_folder, day_file))
      files = natsorted(files)

      merger = pp2.PdfFileMerger()
      for file in files:
        merger.append(file)

      tmp = os.path.join(dst_folder, 'tmp.pdf')
      merger.write(tmp)
      merger.close()

      rsrcmgr = PDFResourceManager()
      codec = 'utf-8'
      laparams = LAParams()
      laparams.detext_vertical=True

      output_nums_list = []
      with open(tmp, 'rb') as fp:
        for nums in range(len(search_text_list)):
          output_nums = []
          output_nums_list.append(output_nums)
          for i, page in enumerate(PDFPage.get_pages(fp)):
            outfp = StringIO()
            device = TextConverter(
              rsrcmgr,
              outfp,
              codec=codec,
              laparams=laparams
              )
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            interpreter.process_page(page)

            extracted_text = outfp.getvalue()

            search_text = search_text_list[nums]
            extracted_page = re.search(search_text, extracted_text)

            if extracted_page:
              output_nums.append(i) #put the page numbers into 'output_nums'


      reader = pp2.PdfFileReader(tmp)
      for name, nums in enumerate(output_nums_list): #get indexes and page numbers from 'output_nums_list'
        file_name = os.path.join(dst_folder, search_text_list[name]+'.pdf')

        writer = pp2.PdfFileWriter()
        with open(file_name, 'wb') as f:
          for i in nums: #get page numbers from each list
            writer.addPage(reader.getPage(i))
            writer.write(f)

      os.remove(tmp)

      self.status_4.set('完了しました')
    
    except NameError:
      self.status_4.set('選択漏れがあります')
    
if __name__ == '__main__':
  MpdfClass()