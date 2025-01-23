#!/usr/bin/env python
# -*- coding: utf8 -*-
# Module     : unprotect.py
# Synopsis   : 破解excel和word文件
# Programmer : 张杰
# Date       : 20250116
# Notes      :
#

import datetime
import logging
import os
import re
import shutil
import threading
import tkinter as tk
from pathlib import Path
from queue import Queue
from tkinter import ttk

import pythoncom
import sys
import time
import win32com.client
from PIL import ImageTk, Image
from tkinterdnd2 import DND_FILES, TkinterDnD

# 用于pyinstaller取得程序目录
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

folder_paths = {}
for k in ("temp", "out"):
    folder_path = os.path.join(application_path, k)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    folder_paths[k] = folder_path

logging.basicConfig(handlers=[logging.FileHandler(filename=f'{application_path}\error.log'
                                                  , encoding='utf-8', mode='a+'),
                              logging.StreamHandler()]
                    , format='%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s'
                    , level=logging.INFO)
logger = logging.getLogger(__name__)


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title('破解word和excel')
        # 窗口尺寸设为实际需要的尺寸
        width = self.winfo_screenwidth()
        height = self.winfo_screenheight()
        # self.minsize(width=800, height=600)
        self.maxsize(width=width, height=height)
        # 禁止改变窗口尺寸
        self.resizable(False, False)
        # 施放文件图标
        image = ImageTk.PhotoImage(
            Image.open(os.path.join(application_path, "dnd.jpg")).resize((250, 250), Image.Resampling.LANCZOS))
        self.lbl_image = ttk.Label(self, image=image)
        self.lbl_image.image = image
        self.lbl_image.drop_target_register(DND_FILES)
        self.lbl_image.dnd_bind('<<Drop>>', lambda e: self.unprotect(e.data))
        self.lbl_image.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        self.lbl_image.grid_columnconfigure(0, weight=1)
        self.p = ttk.Progressbar(self, orient="horizontal", length=200, mode="determinate",
                                 takefocus=True, maximum=100)
        self.p.grid(row=1, column=0, sticky='nsew', padx=5, pady=(0, 5))
        # self.p['value'] = 0
        self.desc_var = tk.StringVar()
        self.msg_desc = tk.Message(self, textvariable=self.desc_var, width=200)
        self.msg_desc.grid(row=2, column=0, sticky='nsw', padx=5, pady=(0, 5))
        self.desc_var.set('状态:')
        self.q = Queue()
        self.is_busy = False
        self.update()
        self.attributes("-topmost", True)
        self.after(20, self.check_message)

    def check_message(self):
        """检查停止标志，检测到就退出"""
        if not self.q.empty():
            message = self.q.get_nowait()
            self.desc_var.set(message["text"])
            self.p['value'] = message["progress_value"]
            if self.p['value'] == 99.9:
                self.is_busy = False
            self.update()
            file_path = message.get("file_path")
            if file_path:
                # print(f"convert sucess, begin cracking {file_path}")
                self.unprotect(file_path)
        self.after(20, self.check_message)

    @staticmethod
    def clear_temp():
        if os.path.exists(folder_paths["temp"]):
            shutil.rmtree(folder_paths["temp"])
        os.makedirs(folder_paths["temp"])

    @staticmethod
    def modify_xml(file_path, pattern):
        with open(file_path, 'r+', encoding="utf-8") as f:
            content = f.read()
            pat = re.compile(pattern)
            match = re.search(pat, content)
            if match:
                cracked = re.sub(pat, "", content)
                f.truncate(0)
                f.seek(0)
                f.write(cracked)
                protected = True
            else:
                # print("can not find protection element")
                protected = False
            return protected

    @staticmethod
    def convert_file(q, id, file_path):
        # print("enter convert_file")
        path = Path(file_path)
        file_name = path.name
        # print(file_name)
        file_base = path.stem
        # print(file_base)
        file_ext = path.suffix
        # print(file_ext)
        now = datetime.datetime.now()  # current date and time
        date_time = now.strftime('_%Y%m%d_%H%M%S')
        message = {}
        message["text"] = '状态:开始格式转换'
        message["progress_value"] = 0
        q.put(message)
        # Initialize
        pythoncom.CoInitialize()
        if file_ext == ".doc":
            word = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(id, pythoncom.IID_IDispatch))
            word.visible = 0
            wb = word.Documents.Open(file_path)
            temp_file = os.path.join(folder_paths["temp"], f"{file_base}{file_ext}x")
            wb.SaveAs2(temp_file, FileFormat=16)  # file format for docx
            wb.Close()
            word.Quit()
            pythoncom.CoUninitialize()
            message["text"] = '状态:格式转换完成，开始破解'
            message["progress_value"] = 99.9
            message["file_path"] = temp_file
            q.put(message)
        else:
            excel = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(id, pythoncom.IID_IDispatch))
            excel.visible = 0
            wb = excel.Workbooks.Open(path)
            temp_file = os.path.join(folder_paths["temp"], f"{file_base}{file_ext}x")
            wb.SaveAs(temp_file, FileFormat=51)  # FileFormat = 51 is for .xlsx extension
            wb.Close()  # FileFormat = 56 is for .xls extension
            excel.Quit()
            pythoncom.CoUninitialize()
            message["text"] = '状态:格式转换完成，开始破解'
            message["progress_value"] = 99.9
            message["file_path"] = temp_file
            q.put(message)

    def unprotect(self, file_path):
        if not self.is_busy:
            file_path = file_path.replace("{", "").replace("}", "")
            path = Path(file_path)
            file_ext = path.suffix
            if file_ext == ".doc":
                # Initialize
                pythoncom.CoInitialize()
                # Get instance
                wd = win32com.client.Dispatch("Word.Application")
                # Create id
                wd_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, wd)
                t = threading.Thread(target=App.convert_file, args=(self.q, wd_id, file_path))
                t.start()
            elif file_ext == ".xls":
                # Initialize
                pythoncom.CoInitialize()
                # Get instance
                xl = win32com.client.Dispatch("Excel.Application")
                # Create id
                xl_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, xl)
                t = threading.Thread(target=App.convert_file, args=(self.q, xl_id, file_path))
                t.start()
            elif file_ext == ".docx":
                file_size = os.path.getsize(file_path)
                if file_size == 0:
                    message = {}
                    message["text"] = '状态:docx为空文档，无需破解'
                    message["progress_value"] = 99.9
                    self.q.put(message)
                else:
                    t = threading.Thread(target=App.main_work, args=(self.q, file_path))
                    t.start()
            elif file_ext == ".xlsx":
                t = threading.Thread(target=App.main_work, args=(self.q, file_path))
                t.start()

    @staticmethod
    def main_work(q, file_path):
        # print("enter main_work")
        file_path = file_path.replace("{", "").replace("}", "")
        message = {}
        message["text"] = "状态：开始破解"
        message["progress_value"] = 0
        q.put(message)
        path = Path(file_path)
        file_name = path.name
        file_base = path.stem
        file_ext = path.suffix
        now = datetime.datetime.now()  # current date and time
        date_time = now.strftime('_%Y%m%d_%H%M%S')
        zip_path = os.path.join(folder_paths["temp"], file_name + date_time)
        if file_ext in (".xlsx", ".docx"):
            pythoncom.CoUninitialize()
            if os.path.isdir(zip_path):
                shutil.rmtree(zip_path)
            shutil.unpack_archive(path, zip_path, "zip")
            message["text"] = '状态:解压完成'
            message["progress_value"] = 20
            q.put(message)
            if os.path.isfile(zip_path):
                os.remove(zip_path)
            protected = False
            if file_ext == ".xlsx":
                workbook_path = os.path.join(zip_path, "xl", "workbook.xml")
                result = App.modify_xml(workbook_path, r"<workbookProtection.*?/>")
                if result:
                    protected = result
                result = App.modify_xml(workbook_path, r"<fileSharing.*?/>")
                if result:
                    protected = result
                sheet_path = os.path.join(zip_path, "xl", "worksheets")
                file_list = os.listdir(sheet_path)
                pat = re.compile(r"^sheet\d+\.xml$")
                for file in file_list:
                    match = re.match(pat, file)
                    if match:
                        result = App.modify_xml(os.path.join(sheet_path, file), r"<sheetProtection.*?/>")
                        if result:
                            protected = result
            else:
                settings_path = os.path.join(zip_path, "word", "settings.xml")
                protected = App.modify_xml(settings_path, r"<w:documentProtection.*?/>")
            if protected:
                message["text"] = '状态:完成修改xml'
                message["progress_value"] = 40
                q.put(message)
                shutil.make_archive(os.path.join(folder_paths["temp"], f"{file_base}{date_time}"), "zip",
                                    root_dir=zip_path, base_dir=".")
                message["text"] = '状态:重新压缩成zip'
                message["progress_value"] = 60
                q.put(message)
                if os.path.isdir(zip_path):
                    shutil.rmtree(zip_path)
                os.rename(os.path.join(folder_paths["temp"], f"{file_base}{date_time}.zip"),
                          os.path.join(folder_paths["out"], f"{file_base}{date_time}{file_ext}"))
                message["text"] = '状态:完成，请到out文件夹中查看'
                message["progress_value"] = 99.9
                q.put(message)
            else:
                message["text"] = '状态:文件没有被保护！'
                message["progress_value"] = 99.9
                q.put(message)
            App.clear_temp()


if __name__ == "__main__":
    app = App()
    app.mainloop()
