import tkinter as tk
from tkinter import messagebox
import pyautogui
import pygetwindow as gw
import os
import tempfile
import time
try:
    import win32com.client
except ImportError:
    win32com = None

class ScreenshotExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title('スクリーンショットアプリ')
        self.root.geometry('300x150')
        self.excel = None
        self.wb = None
        self.ws = None
        self.current_row = 2
        self.setup_excel()
        btn1 = tk.Button(root, text='スクリーンショット', font=('Arial', 14), width=20, command=self.take_screenshot)
        btn1.pack(pady=10)
        btn2 = tk.Button(root, text='次のシート', font=('Arial', 14), width=20, command=self.next_sheet)
        btn2.pack(pady=10)

    def setup_excel(self):
        if win32com is None:
            messagebox.showerror('エラー', 'この機能はWindowsでのみ動作します')
            self.root.destroy()
            return
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.excel.Visible = True
        self.wb = self.excel.Workbooks.Add()
        self.ws = self.wb.Worksheets(1)
        self.ws.Name = 'Sheet1'
        self.current_row = 2

    def take_screenshot(self):
        # Chromeウィンドウ取得
        chrome_windows = [w for w in gw.getAllWindows() if 'chrome' in w.title.lower() and w.isActive]
        if not chrome_windows:
            messagebox.showerror('エラー', 'アクティブなChromeウィンドウが見つかりません')
            return
        win = chrome_windows[0]
        # ウィンドウ位置・サイズ取得
        left, top, width, height = win.left, win.top, win.width, win.height
        # スクリーンショット
        img = pyautogui.screenshot(region=(left, top, width, height))
        tmpfile = os.path.join(tempfile.gettempdir(), f'ss_{int(time.time())}.png')
        img.save(tmpfile)
        # Excelに画像貼り付け
        self.ws.Pictures().Insert(tmpfile).Select()
        self.excel.Selection.Top = self.ws.Rows(self.current_row).Top
        self.current_row += 25  # 画像の高さ分だけ行を下げる
        os.remove(tmpfile)

    def next_sheet(self):
        new_sheet = self.wb.Worksheets.Add(After=self.wb.Worksheets(self.wb.Worksheets.Count))
        idx = self.wb.Worksheets.Count
        new_sheet.Name = f'Sheet{idx}'
        self.ws = new_sheet
        self.current_row = 2

def main():
    root = tk.Tk()
    app = ScreenshotExcelApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
