# JPX Prime Investor Flow v1.8

JPX「投資部門別売買状況」の週次・金額データから、東証プライムだけを抽出して可視化するStreamlitアプリです。

## v1.8 の変更点

- JPX旧形式でExcel hrefがHTMLから取れない場合に、**金額PDF (`stock_val_1_*.pdf`) の実hrefから同じディレクトリの `.xls` URLを復元**します。
- 復元URLは推測しただけで採用せず、HTTP GET後に **Content-Type / XLS magic bytes (D0 CF 11 E0) / XLSX ZIP signature (PK)** を確認してから読み込みます。
- Excelを実際に開けた場合だけシート解析へ進み、**「東証プライム / Prime」シートを明示的に選択**します。
- 売り・買い・差引について、`差引 = 買い - 売り`、負の売買金額がないこと、売り買いの桁が極端に崩れていないことを検算します。
- 週次 / 月次、iPhone向けレスポンシブ表示を継続しています。

## JPX旧形式のURL

2026年9月28日までの金額資料は、PDFとExcelが同じ添付ディレクトリに置かれ、例として `stock_val_1_260104.pdf` に対し `stock_val_1_260104.xls` が存在します。v1.8はページからランダムな添付ディレクトリを取得するためにPDF hrefを利用します。

## デプロイ

1. ZIPを展開
2. GitHubリポジトリの `streamlit_app.py` と `requirements.txt` を上書き
3. Streamlit Community CloudをReboot

エラー時は「データ取得状況 / エラー」に、試したExcel候補URLと取得エラーが表示されます。
