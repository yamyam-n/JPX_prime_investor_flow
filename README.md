# JPX Prime Investor Flow v1.5

JPX「投資部門別売買状況（株式・週間）」から東証プライムの金額データを取得して可視化する Streamlit アプリです。

## v1.5 fixes
- JPXのリンク名・拡張子・アイコン文言に依存しません。
- 各週の行から href / data-* / onclick 等の候補URLを収集し、実際のレスポンス先頭バイトで Excel を判定します。
- 同じ行に「株数」「金額」Excelがある場合は、ブック内の「金額」「千円」等を読んで金額ファイルを選びます。
- 取得できない場合は Streamlit 画面に「リンク診断」を表示します。
- 週次／月次、iPhone向け表示を維持しています。

## deploy
ZIPを展開し、GitHubリポジトリのファイルを置き換えてください。Main file path は `streamlit_app.py` です。
