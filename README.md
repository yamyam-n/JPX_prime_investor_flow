# JPX Prime Investor Flow

JPX「投資部門別売買状況（株式・週間）」から東証プライムの金額データを取得し、投資主体別のネット売買を可視化するStreamlitアプリです。

## 主な表示
- 海外投資家 / 個人 / 信託銀行 / 事業法人の直近ネット売買
- 投資部門別ネット売買の週次棒グラフ
- 期間累積ネット売買
- 直近週ランキング
- 売り・買い・差引表

## 起動
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Community Cloud
GitHubリポジトリ直下に `streamlit_app.py` と `requirements.txt` を置き、Main file path に `streamlit_app.py` を指定してください。

## JPXフォーマット変更への対応
2026-09-29掲載分から週間Excelが1ファイル・1シート集約形式へ変更予定です。このv1は現行形式と新形式の双方をラベル探索で読む設計にしています。ただし実際の新ファイル公開後、列配置に応じた微調整が必要になる可能性があります。
