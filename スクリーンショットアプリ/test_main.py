import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import types
import pytest
from unittest import mock
import builtins

# テスト対象のクラスをインポート
importスクリーンショットアプリ_main = __import__("スクリーンショットアプリ.main", fromlist=["ScreenshotExcelApp"])

@pytest.fixture
def mock_root():
    root = mock.Mock()
    root.title = mock.Mock()
    root.geometry = mock.Mock()
    root.destroy = mock.Mock()
    root.update_idletasks = mock.Mock()
    root.winfo_rootx = mock.Mock(return_value=100)
    root.winfo_rooty = mock.Mock(return_value=100)
    root.winfo_width = mock.Mock(return_value=500)
    root.winfo_height = mock.Mock(return_value=200)
    root.iconify = mock.Mock()
    root.update = mock.Mock()
    root.deiconify = mock.Mock()
    return root

@pytest.fixture
def patch_win32com(monkeypatch):
    # win32com.client.Dispatchのモック
    class MockExcel:
        def __init__(self):
            self.Visible = False
            self.Workbooks = mock.Mock()
            self.Workbooks.Add.return_value = self
            self.Worksheets = mock.Mock()
            self.Worksheets.__getitem__ = lambda s, i: self
            self.Worksheets.__call__ = lambda s, i=None: self
            self.Worksheets.Add = mock.Mock(return_value=self)
            self.Worksheets.Count = 1
            self.Name = ""
            self.Rows = mock.Mock()
            self.Rows.__getitem__ = lambda s, i: self
            self.Rows.Top = 0
            self.Columns = mock.Mock()
            self.Columns.__getitem__ = lambda s, i: self
            self.Columns.Left = 0
            self.Pictures = mock.Mock()
            self.Pictures().Insert = mock.Mock(return_value=mock.Mock(Select=mock.Mock()))
            self.Selection = mock.Mock()
            self.Selection.Top = 0
            self.Selection.Left = 0

    mock_win32com = types.SimpleNamespace()
    mock_win32com.client = types.SimpleNamespace()
    mock_win32com.client.Dispatch = mock.Mock(return_value=MockExcel())
    monkeypatch.setitem(sys.modules, "win32com", mock_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", mock_win32com.client)
    return mock_win32com

@pytest.fixture(autouse=True)
def patch_pyautogui(monkeypatch):
    mock_img = mock.Mock()
    mock_img.width = 960
    mock_img.height = 400
    mock_img.save = mock.Mock()
    monkeypatch.setitem(sys.modules, "pyautogui", mock.Mock(screenshot=mock.Mock(return_value=mock_img)))
    return mock_img

def test_setup_excel_win(monkeypatch, mock_root, patch_win32com):
    # messagebox.showerrorが呼ばれないこと
    monkeypatch.setattr("tkinter.messagebox.showerror", mock.Mock())
    app = importスクリーンショットアプリ_main.ScreenshotExcelApp(mock_root)
    assert app.excel is not None
    assert app.wb is not None
    assert app.ws is not None
    assert app.current_row == 4

def test_setup_excel_nonwin(monkeypatch, mock_root):
    # win32comがNoneの場合
    monkeypatch.setattr(importスクリーンショットアプリ_main, "win32com", None)
    showerror = mock.Mock()
    monkeypatch.setattr("tkinter.messagebox.showerror", showerror)
    app = importスクリーンショットアプリ_main.ScreenshotExcelApp(mock_root)
    showerror.assert_called_once()
    mock_root.destroy.assert_called_once()

def test_next_sheet(monkeypatch, mock_root, patch_win32com):
    app = importスクリーンショットアプリ_main.ScreenshotExcelApp(mock_root)
    # Worksheets.Addの戻り値をモック
    new_sheet = mock.Mock()
    app.wb.Worksheets.Add.return_value = new_sheet
    app.wb.Worksheets.Count = 2
    app.next_sheet()
    assert app.ws == new_sheet
    assert app.current_row == 4
    new_sheet.Name = f"Sheet2"

def test_take_screenshot_screeninfo_missing(monkeypatch, mock_root, patch_win32com):
    app = importスクリーンショットアプリ_main.ScreenshotExcelApp(mock_root)
    monkeypatch.setitem(sys.modules, "screeninfo", None)
    showerror = mock.Mock()
    monkeypatch.setattr("tkinter.messagebox.showerror", showerror)
    app.take_screenshot()
    showerror.assert_called_once()
    # エラー時はroot.destroyは呼ばれない
    mock_root.destroy.assert_not_called()

def test_take_screenshot_right_screeninfo_missing(monkeypatch, mock_root, patch_win32com):
    app = importスクリーンショットアプリ_main.ScreenshotExcelApp(mock_root)
    monkeypatch.setitem(sys.modules, "screeninfo", None)
    showerror = mock.Mock()
    monkeypatch.setattr("tkinter.messagebox.showerror", showerror)
    app.take_screenshot_right()
    showerror.assert_called_once()
    mock_root.destroy.assert_not_called()