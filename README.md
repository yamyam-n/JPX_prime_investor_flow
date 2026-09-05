# JPX Prime Investor Flow v1.2

東証プライムの「投資部門別売買状況」を可視化する Streamlit アプリです。

## v1.2
- iPhone / スマートフォン向けレスポンシブ表示
- スマホでは設定・KPI・表を縦1列に配置
- Plotlyツールバーを非表示にして表示領域を確保
- 週次X軸を `8/03` のような短い日付に変更
- 長いExcelファイル名がX軸に出るケースを抑制
- 週次 / 月次切替
- 投資家別に「売り=マイナス」「買い=プラス」「差引=折れ線」を表示

## Deploy
GitHub リポジトリへファイル一式を置き、Streamlit Community Cloud で `streamlit_app.py` を指定してください。
