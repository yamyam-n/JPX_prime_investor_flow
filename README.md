# JPX Prime Investor Flow v1.4

JPX「投資部門別売買状況（株式・週間）」から東証プライムの金額データを取得し、週次/月次で可視化するStreamlitアプリです。

## v1.4 修正点
- JPXのExcel添付URLが `.xls/.xlsx` で終わらない場合にも対応
- 「金額」セル内のExcel/添付リンクをセル位置から直接取得
- PDFを除外し、ダウンロード後はファイル先頭バイト/Content-TypeでExcelか検証
- バックナンバーの `<select><option>` URLも探索
- iPhone向けレスポンシブ表示を維持

## 起動
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
