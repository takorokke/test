import sys
from unittest import mock
# テスト時のみWindows依存/GUI系モジュールを自動でモック化
if not (__name__ == '__main__'):
    sys.modules['pyautogui'] = mock.Mock()
    sys.modules['screeninfo'] = mock.Mock()
    sys.modules['pywinauto'] = mock.Mock()
    sys.modules['win32com'] = mock.Mock()
    sys.modules['win32com.client'] = mock.Mock()
    sys.modules['pygetwindow'] = mock.Mock()
    import tkinter
    tkinter.Tk = mock.Mock()
    tkinter.Button = mock.Mock()
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
        self.root.geometry('500x200')  # ウィンドウサイズを大きく
        self.excel = None
        self.wb = None
        self.ws = None
        self.current_row = 4
        self.current_col = 1  # 1=A列
        self.last_img_col = 1
        self.setup_excel()
        btn1 = tk.Button(root, text='スクリーンショット', font=('Arial', 14), width=20, command=self.take_screenshot)
        btn1.pack(pady=5)
        btn2 = tk.Button(root, text='右へスクリーンショット', font=('Arial', 14), width=20, command=self.take_screenshot_right)
        btn2.pack(pady=5)
        btn3 = tk.Button(root, text='次のシート', font=('Arial', 14), width=20, command=self.next_sheet)
        btn3.pack(pady=5)

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
        # 常にA列から貼り付け開始
        self.current_col = 1
        # 直前の画像の下の行に進める（右貼り付け後も必ず下に進む）
        if hasattr(self, 'last_img_row') and hasattr(self, 'last_img_height'):
            self.current_row = self.last_img_row + self.last_img_height + 1
        # アプリウィンドウの中央座標を取得
        self.root.update_idletasks()
        app_x = self.root.winfo_rootx()
        app_y = self.root.winfo_rooty()
        app_w = self.root.winfo_width()
        app_h = self.root.winfo_height()
        center_x = app_x + app_w // 2
        center_y = app_y + app_h // 2

        # すべてのディスプレイ情報を取得
        try:
            import screeninfo
            screens = screeninfo.get_monitors()
        except ImportError:
            messagebox.showerror('エラー', 'screeninfoパッケージが必要です。\npip install screeninfo を実行してください。')
            return
        # アプリが表示されているディスプレイを特定
        target_screen = None
        for s in screens:
            if s.x <= center_x < s.x + s.width and s.y <= center_y < s.y + s.height:
                target_screen = s
                break
        if not target_screen:
            messagebox.showerror('エラー', 'アプリが表示されている画面が見つかりません')
            return
        # アプリを一時的に最小化
        self.root.iconify()
        self.root.update()
        time.sleep(0.5)
        # ブラウザウィンドウを特定（中心座標がこの画面内にあるものを優先）
        import pygetwindow as gw
        browser_keywords = ['chrome', 'edge', 'firefox', 'opera', 'safari', 'brave']
        browser_window = None
        for w in gw.getAllWindows():
            if not w.visible:
                continue
            title = w.title.lower()
            if not any(k in title for k in browser_keywords):
                continue
            wx_center = w.left + w.width // 2
            wy_center = w.top + w.height // 2
            if (target_screen.x <= wx_center < target_screen.x + target_screen.width and
                target_screen.y <= wy_center < target_screen.y + target_screen.height):
                browser_window = w
                break
        if not browser_window:
            self.root.deiconify()
            self.root.update()
            messagebox.showerror('エラー', 'この画面内にブラウザウィンドウが見つかりません')
            return
        # pywinautoでクライアント領域（Web表示部分）を取得
        try:
            from pywinauto import Application
            app = Application(backend="uia").connect(handle=browser_window._hWnd)
            dlg = app.window(handle=browser_window._hWnd)
            child = None
            for c in dlg.children():
                cname = c.friendly_class_name().lower()
                if 'render' in cname or 'chrome' in cname or 'tabwindow' in cname or 'content' in cname:
                    child = c
                    break
            if child:
                rect = child.rectangle()
                bbox = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
            else:
                bbox = (browser_window.left, browser_window.top, browser_window.width, browser_window.height)
        except Exception as e:
            bbox = (browser_window.left, browser_window.top, browser_window.width, browser_window.height)
        # pyautoguiでキャプチャ（黒画像対策）
        import pyautogui
        img = pyautogui.screenshot(region=bbox)
        # O列右端に合わせてリサイズ（O列は15列目、幅は約15*64=960px）
        target_width = 960
        if img.width > target_width:
            from PIL import Image
            img = img.resize((target_width, int(img.height * target_width / img.width)), Image.LANCZOS)
        tmpfile = os.path.join(tempfile.gettempdir(), f'ss_{int(time.time())}.png')
        img.save(tmpfile)
        # Excelに画像貼り付け
        pic = self.ws.Pictures().Insert(tmpfile)
        pic.Select()
        self.excel.Selection.Top = self.ws.Rows(self.current_row).Top
        self.excel.Selection.Left = self.ws.Columns(self.current_col).Left
        self.excel.Selection.Top = self.ws.Rows(self.current_row).Top
        self.last_img_row = self.current_row
        self.last_img_height = int(img.height / 20)  # 1行=約20px
        self.right_row = self.current_row
        self.last_img_col = 1 + int(img.width / 64)  # 1列=約64px
        # 右貼り付け用の行位置をリセット
        self.right_row = self.current_row
        os.remove(tmpfile)
        self.root.deiconify()
        self.root.update()
        # スクリーンショット貼り付け後、次の行に進める
        self.current_row += self.last_img_height + 1

    def take_screenshot_right(self):
        # アプリウィンドウの中央座標を取得
        self.root.update_idletasks()
        app_x = self.root.winfo_rootx()
        app_y = self.root.winfo_rooty()
        app_w = self.root.winfo_width()
        app_h = self.root.winfo_height()
        center_x = app_x + app_w // 2
        center_y = app_y + app_h // 2

        # すべてのディスプレイ情報を取得
        try:
            import screeninfo
            screens = screeninfo.get_monitors()
        except ImportError:
            messagebox.showerror('エラー', 'screeninfoパッケージが必要です。\npip install screeninfo を実行してください。')
            return
        # アプリが表示されているディスプレイを特定
        target_screen = None
        for s in screens:
            if s.x <= center_x < s.x + s.width and s.y <= center_y < s.y + s.height:
                target_screen = s
                break
        if not target_screen:
            messagebox.showerror('エラー', 'アプリが表示されている画面が見つかりません')
            return
        # アプリを一時的に最小化
        self.root.iconify()
        self.root.update()
        time.sleep(0.5)
        # ブラウザウィンドウを特定（中心座標がこの画面内にあるものを優先）
        import pygetwindow as gw
        browser_keywords = ['chrome', 'edge', 'firefox', 'opera', 'safari', 'brave']
        browser_window = None
        for w in gw.getAllWindows():
            if not w.visible:
                continue
            title = w.title.lower()
            if not any(k in title for k in browser_keywords):
                continue
            wx_center = w.left + w.width // 2
            wy_center = w.top + w.height // 2
            if (target_screen.x <= wx_center < target_screen.x + target_screen.width and
                target_screen.y <= wy_center < target_screen.y + target_screen.height):
                browser_window = w
                break
        if not browser_window:
            self.root.deiconify()
            self.root.update()
            messagebox.showerror('エラー', 'この画面内にブラウザウィンドウが見つかりません')
            return
        # pywinautoでクライアント領域（Web表示部分）を取得
        try:
            from pywinauto import Application
            app = Application(backend="uia").connect(handle=browser_window._hWnd)
            dlg = app.window(handle=browser_window._hWnd)
            child = None
            for c in dlg.children():
                cname = c.friendly_class_name().lower()
                if 'render' in cname or 'chrome' in cname or 'tabwindow' in cname or 'content' in cname:
                    child = c
                    break
            if child:
                rect = child.rectangle()
                bbox = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
            else:
                bbox = (browser_window.left, browser_window.top, browser_window.width, browser_window.height)
        except Exception as e:
            bbox = (browser_window.left, browser_window.top, browser_window.width, browser_window.height)
        # pyautoguiでキャプチャ（黒画像対策）
        import pyautogui
        img = pyautogui.screenshot(region=bbox)
        # O列右端に合わせてリサイズ（O列は15列目、幅は約15*64=960px）
        target_width = 960
        if img.width > target_width:
            from PIL import Image
            img = img.resize((target_width, int(img.height * target_width / img.width)), Image.LANCZOS)
        tmpfile = os.path.join(tempfile.gettempdir(), f'ss_{int(time.time())}.png')
        img.save(tmpfile)
        # 直前の画像の右側に数列空けて貼り付け
        col_gap = 3  # 画像間の列の空き
        if not hasattr(self, 'right_row') or self.right_row != getattr(self, 'last_img_row', 4):
            self.right_row = getattr(self, 'last_img_row', 4)
            self.last_img_col = 1
        self.current_col = self.last_img_col + col_gap
        self.current_row = self.right_row
        pic = self.ws.Pictures().Insert(tmpfile)
        pic.Select()
        self.excel.Selection.Top = self.ws.Rows(self.current_row).Top
        self.excel.Selection.Left = self.ws.Columns(self.current_col).Left
        img_height = img.height if hasattr(img, 'height') else 400
        self.last_img_col = self.current_col + int(img.width / 64)
        self.last_img_row = self.current_row
        self.last_img_height = int(img_height / 20)
        os.remove(tmpfile)
        self.root.deiconify()
        self.root.update()

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
