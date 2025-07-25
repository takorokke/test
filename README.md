# スクリーンショットアプリ

## 概要
Windowsで動作する、ボタン一つでブラウザ画面のスクリーンショットをExcelに自動貼り付けするアプリです。

## 主な機能
- アプリ起動時にExcelも自動起動
- 「スクリーンショット」ボタンで、アプリが表示されている画面内の最前面ブラウザウィンドウのWeb表示領域を自動キャプチャし、ExcelのO列右端に合わせて貼り付け
- 「右へスクリーンショット」ボタンで、前の画像の右側に数列空けて貼り付け
- 連続で押すと画像が重ならないように自動配置
- 「次のシート」ボタンで新しいシートに切り替え
- マルチディスプレイ対応

## 必要環境
- Windows 10/11
- Python 3.8以降
- Excel（Microsoft Office）

## インストール
1. Pythonをインストール
2. コマンドプロンプトで本フォルダに移動し、以下を実行

```
pip install -r requirements.txt
```

## 使い方
1. `main.py` を実行

```
python main.py
```

2. アプリ画面が表示されたら、スクリーンショットを撮りたいブラウザをアプリを開いた画面に移動し、最前面にして「スクリーンショット」または「右へスクリーンショット」ボタンを押す
3. Excelに自動で画像が貼り付けられます

## 注意
- スクリーンショットは「アプリを開いた画面内の最前面ブラウザウィンドウ」だけが対象です
- アプリを開いた画面と異なる画面を撮影しようとすると黒画像になる場合があります。必ずアプリを開いた画面を撮影してください。
- ブラウザをF11キーで全画面表示にするとアドレスバーやブックマークバーが消え、Webページ部分だけをキャプチャできます
- アドレスバー高さは`main.py`内で調整可能です
- 権限や環境によっては管理者実行が必要な場合があります