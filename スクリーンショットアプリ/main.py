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
        # ボタンの真下の座標（アプリ中央下部）
        target_x = app_x + app_w // 2
        target_y = app_y + app_h + 5  # アプリ下端より少し下

        # 真下にあるウィンドウを特定
        target_window = None
        for w in gw.getAllWindows():
            if not w.visible or w._hWnd == self.root.winfo_id():
                continue
            wx1, wy1, wx2, wy2 = w.left, w.top, w.left + w.width, w.top + w.height
            if wx1 <= target_x <= wx2 and wy1 <= target_y <= wy2:
                target_window = w
                break
        if not target_window:
            messagebox.showerror('エラー', 'ボタンの真下にあるウィンドウが見つかりません')
            return
        # スクリーンショット
        left, top, width, height = target_window.left, target_window.top, target_window.width, target_window.height
        img = pyautogui.screenshot(region=(left, top, width, height))
        # G列右端に合わせてリサイズ（G列は約7列目、幅は約7*8.43=59.01ポイント=約780px）
        target_width = 780
        if width > target_width:
            from PIL import Image
            img = img.resize((target_width, int(height * target_width / width)), Image.LANCZOS)
        tmpfile = os.path.join(tempfile.gettempdir(), f'ss_{int(time.time())}.png')
        img.save(tmpfile)
        # Excelに画像貼り付け
        pic = self.ws.Pictures().Insert(tmpfile)
        pic.Select()
        self.excel.Selection.Top = self.ws.Rows(self.current_row).Top
        self.excel.Selection.Left = self.ws.Columns(1).Left  # A列左端
        # 画像の高さを取得し、次回貼り付け位置を自動調整
        img_height = img.height if hasattr(img, 'height') else 400
        row_height = 20  # Excelの1行の高さ（おおよそ）
        add_rows = int(img_height / row_height) + 3  # 画像の高さ分＋3行空ける
        self.current_row += add_rows
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
