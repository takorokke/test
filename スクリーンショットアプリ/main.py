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
        self.current_row = 4  # 3行空けて4行目から貼り付け

    def take_screenshot(self):
        # アプリウィンドウの位置とサイズを取得
        self.root.update_idletasks()
        app_x = self.root.winfo_rootx()
        app_y = self.root.winfo_rooty()
        app_w = self.root.winfo_width()
        app_h = self.root.winfo_height()
        app_rect = (app_x, app_y, app_x + app_w, app_y + app_h)

        # アプリウィンドウと重なる他のウィンドウをリストアップ
        candidates = []
        for w in gw.getAllWindows():
            if not w.isVisible or w._hWnd == self.root.winfo_id():
                continue
            wx1, wy1, wx2, wy2 = w.left, w.top, w.left + w.width, w.top + w.height
            # 重なり判定
            if (wx1 < app_rect[2] and wx2 > app_rect[0] and wy1 < app_rect[3] and wy2 > app_rect[1]):
                area = w.width * w.height
                candidates.append((area, w))
        if not candidates:
            messagebox.showerror('エラー', '下にあるウィンドウが見つかりません')
            return
        # 一番大きいウィンドウを選択
        candidates.sort(reverse=True)
        win = candidates[0][1]
        # スクリーンショット
        left, top, width, height = win.left, win.top, win.width, win.height
        img = pyautogui.screenshot(region=(left, top, width, height))
        tmpfile = os.path.join(tempfile.gettempdir(), f'ss_{int(time.time())}.png')
        img.save(tmpfile)
        # Excelに画像貼り付け
        self.ws.Pictures().Insert(tmpfile).Select()
        self.excel.Selection.Top = self.ws.Rows(self.current_row).Top
        self.current_row += 25
        os.remove(tmpfile)

    def next_sheet(self):
        new_sheet = self.wb.Worksheets.Add(After=self.wb.Worksheets(self.wb.Worksheets.Count))
        idx = self.wb.Worksheets.Count
        new_sheet.Name = f'Sheet{idx}'
        self.ws = new_sheet
        self.current_row = 4  # 新しいシートも4行目から貼り付け

def main():
    root = tk.Tk()
    app = ScreenshotExcelApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
