import io
import re
from datetime import datetime
from urllib.parse import urljoin

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title='JPX プライム 投資部門別売買', page_icon='📊', layout='wide', initial_sidebar_state='collapsed')

st.markdown('''
<style>
.block-container {padding-top: .9rem; padding-bottom: 2rem; max-width: 1180px;}
[data-testid="stMetricValue"] {font-size: 1.48rem;}
[data-testid="stPlotlyChart"] {width: 100% !important;}
@media (max-width: 700px) {
  .block-container {padding-left: .55rem !important; padding-right: .55rem !important; padding-top: .45rem !important;}
  h1 {font-size: 1.45rem !important; line-height: 1.22 !important;}
  h2, h3 {font-size: 1.12rem !important;}
  [data-testid="stHorizontalBlock"] {flex-wrap: wrap !important; gap: .28rem !important;}
  [data-testid="column"] {min-width: 100% !important; width: 100% !important; flex: 1 1 100% !important;}
  [data-testid="stMetric"] {padding: .38rem .5rem !important;}
  [data-testid="stMetricValue"] {font-size: 1.28rem !important;}
  .stPlotlyChart {overflow-x: hidden !important;}
}
</style>
''', unsafe_allow_html=True)

JPX_WEEKLY = 'https://www.jpx.co.jp/markets/statistics-equities/investor-type/index.html'
JPX_WEEKLY_ARCHIVE_2026 = 'https://www.jpx.co.jp/markets/statistics-equities/investor-type/00-00-archives-00.html'
HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/152 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Language': 'ja,en-US;q=0.8,en;q=0.6'}
MAJOR = ['海外投資家', '個人', '信託銀行', '事業法人', '投資信託', '証券会社']
ALIASES = {
    '海外投資家': ['海外投資家', '外国人', '外国法人等', 'Foreigners'],
    '個人': ['個人', '個人・その他', 'Individuals'],
    '信託銀行': ['信託銀行', 'TrustBK', 'TrustBank'],
    '事業法人': ['事業法人', 'BusinessCos'],
    '投資信託': ['投資信託', 'Investment'],
    '証券会社': ['証券会社', 'SecuritiesCos'],
    '生保・損保': ['生保・損保', '生損保', 'Life&Non-Life'],
    '都銀・地銀等': ['都銀・地銀等', '都銀・地銀', 'City&RegionalBK'],
    'その他金融機関': ['その他金融機関', 'OtherFinancials'],
    'その他法人等': ['その他法人等', 'その他法人', 'OtherCos'],
    '自己計': ['自己計', 'Proprietary'],
    '委託計': ['委託計', 'Brokerage'],
}


def clean_text(x):
    if pd.isna(x):
        return ''
    s = str(x).replace('\u3000', '').replace('\n', '').replace('\r', '')
    return re.sub(r'\s+', '', s)


def num(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).replace(',', '').replace('△', '-').replace('▲', '-').strip()
    s = re.sub(r'[^0-9.\-]', '', s)
    if s in ('', '-', '.', '-.'):
        return np.nan
    try:
        return float(s)
    except Exception:
        return np.nan


def canonical_label(text):
    t = clean_text(text)
    for canon, aliases in ALIASES.items():
        if any(clean_text(a).lower() in t.lower() for a in aliases):
            return canon
    return None


def parse_period_text(text):
    """JPX行見出しから週の開始日・終了日を取り出す。"""
    t = str(text)
    # 2026年8月第4週(8月24日～8月28日)
    ym = re.search(r'(20\d{2})年\s*(\d{1,2})月', t)
    year = int(ym.group(1)) if ym else None
    md = re.findall(r'(\d{1,2})月\s*(\d{1,2})日', t)
    if year and md:
        dates = [pd.Timestamp(year, int(m), int(d)) for m, d in md]
        return min(dates), max(dates)
    # 8/24 - 8/28, year may appear separately
    slash = re.findall(r'(?<!\d)(\d{1,2})[/-](\d{1,2})(?!\d)', t)
    if slash and year:
        dates = [pd.Timestamp(year, int(m), int(d)) for m, d in slash]
        return min(dates), max(dates)
    return pd.NaT, pd.NaT


def _all_jpx_links_from_row(tr, base_url):
    """行内のJPXリンクを拡張子で絞らず回収する。"""
    hits = []
    for a in tr.find_all('a', href=True):
        href = (a.get('href') or '').strip()
        if not href or href.lower().startswith(('javascript:', '#')):
            continue
        u = urljoin(base_url, href)
        if 'jpx.co.jp' in u.lower():
            hits.append({'url': u, 'text': clean_text(a.get_text(' ', strip=True))})
    return hits


def _excel_candidates_from_links(links):
    """JPX旧形式の金額Excel候補を作る。

    2026/9/28までのJPX資料では、金額PDFは `stock_val_1_YYMMWW.pdf`、
    Excelは同じディレクトリ・同じベース名の `.xls` で公開されている。
    HTML側でExcel hrefが取得できない場合でも、金額PDFの実hrefからExcelを復元する。
    """
    direct = []
    derived = []
    amount_pdfs = []
    generic_excels = []
    for item in links:
        u = item['url']
        lo = u.lower()
        txt = item.get('text', '')
        if re.search(r'\.xlsx?(?:$|\?)', lo):
            if 'stock_val' in lo or '金額' in txt or 'value' in txt.lower():
                direct.append(u)
            else:
                generic_excels.append(u)
        if re.search(r'\.pdf(?:$|\?)', lo) and ('stock_val' in lo or '金額' in txt or 'value' in txt.lower()):
            amount_pdfs.append(u)

    # 実在する金額PDFのディレクトリをそのまま使い、拡張子だけExcelへ。
    for pdf in amount_pdfs:
        base = re.sub(r'\.pdf(?=($|\?))', '', pdf, flags=re.I)
        # 旧形式の本命は .xls。念のため .xlsx も候補にする。
        derived.extend([base + '.xls', base + '.xlsx'])

    # 新形式は固定ファイル名 stock_1_w_YYYYMMDD_YYYYMMDD.xlsx
    unified = [item['url'] for item in links if re.search(r'/stock_1_w_\d{8}_\d{8}\.xlsx(?:$|\?)', item['url'].lower())]
    out = unified + direct + derived + generic_excels
    return list(dict.fromkeys(out)), amount_pdfs


def extract_week_rows(html, base_url):
    """週次表を日付ブロック化し、金額Excel候補を作る。"""
    soup = BeautifulSoup(html, 'html.parser')
    blocks, current = [], None
    for tr in soup.find_all('tr'):
        row_text = tr.get_text(' ', strip=True)
        start, end = parse_period_text(row_text)
        links = _all_jpx_links_from_row(tr, base_url)
        if pd.notna(end):
            if current is not None:
                blocks.append(current)
            current = {'text': row_text, 'start': start, 'end': end, 'links': list(links)}
        elif current is not None and links:
            current['links'].extend(links)
    if current is not None:
        blocks.append(current)

    found, diagnostics = [], []
    for b in blocks:
        # URL単位で重複排除
        uniq = []
        seen = set()
        for x in b['links']:
            if x['url'] not in seen:
                seen.add(x['url']); uniq.append(x)
        candidates, pdfs = _excel_candidates_from_links(uniq)
        if candidates:
            found.append({
                'text': b['text'], 'start': b['start'], 'end': b['end'],
                'candidates': candidates, 'amount_pdfs': pdfs,
                'all_links': [x['url'] for x in uniq],
            })
        else:
            diagnostics.append({'text': b['text'], 'candidates': [x['url'] for x in uniq]})
    return found, diagnostics


@st.cache_data(ttl=3600, show_spinner=False)
def discover_week_files(max_weeks=26):
    pages = [JPX_WEEKLY_ARCHIVE_2026, JPX_WEEKLY]
    records, diagnostics, page_errors = [], [], []
    for page in pages:
        try:
            r = requests.get(page, headers=HEADERS, timeout=30)
            r.raise_for_status()
            more, diag = extract_week_rows(r.text, page)
            records.extend(more); diagnostics.extend(diag)
        except Exception as e:
            page_errors.append(f'{page}: {e}')
    by_end = {}
    for rec in records:
        key = rec['end']
        if pd.notna(key) and key not in by_end:
            by_end[key] = rec
    out = sorted(by_end.values(), key=lambda x: x['end'], reverse=True)
    if not out and page_errors:
        diagnostics.append({'text': 'ページ取得エラー', 'candidates': page_errors})
    return out[:max_weeks], diagnostics


def _excel_kind(content, ctype=''):
    ctype = (ctype or '').lower()
    if content[:2] == b'PK' or 'spreadsheetml' in ctype or 'xlsx' in ctype:
        return 'xlsx'
    if content[:4] == bytes.fromhex('D0CF11E0') or 'ms-excel' in ctype:
        return 'xls'
    return None


@st.cache_data(ttl=3600, show_spinner=False)
def read_excel_candidates(candidates):
    """候補URLを順に実GETし、Excel本体であることをmagic bytes/content-typeで確認して開く。"""
    errors = []
    for url in candidates:
        try:
            r = requests.get(url, headers=HEADERS, timeout=30, allow_redirects=True)
            r.raise_for_status()
            kind = _excel_kind(r.content, r.headers.get('Content-Type'))
            if not kind:
                errors.append(f'{url} -> Excelではない ({r.headers.get("Content-Type")})')
                continue
            engine = 'openpyxl' if kind == 'xlsx' else 'xlrd'
            sheets = pd.read_excel(io.BytesIO(r.content), sheet_name=None, header=None, engine=engine)
            if not sheets:
                errors.append(f'{url} -> Excelは開けたがシートなし')
                continue
            return url, kind, sheets
        except Exception as e:
            errors.append(f'{url} -> {e}')
    raise ValueError('Excel候補を取得・読込できませんでした:\n' + '\n'.join(errors[:8]))


def sheet_score_for_prime(name, df):
    score = 0
    n = clean_text(name).lower()
    if 'プライム' in n or 'prime' in n:
        score += 100
    top = ''.join(clean_text(v) for v in df.iloc[:15, :min(5, df.shape[1])].values.flatten()).lower()
    if 'プライム' in top or 'prime' in top:
        score += 80
    # 新市場移行前の古い「東証1部」はPrimeとは別物なので自動採用しない
    if 'スタンダード' in top or 'グロース' in top:
        score -= 50
    return score


def choose_prime_sheet(sheets):
    scored = [(sheet_score_for_prime(n, d), n, d) for n, d in sheets.items()]
    scored.sort(key=lambda x: x[0], reverse=True)
    if scored and scored[0][0] > 0:
        return scored[0][1], scored[0][2]
    # 新形式は全市場1シートなので、そのまま最大シートを返す
    name = max(sheets, key=lambda k: sheets[k].shape[0] * sheets[k].shape[1])
    return name, sheets[name]


def detect_unit_multiplier(df):
    """Excel記載の単位を億円への倍率に変換。"""
    text = ' '.join(clean_text(v) for v in df.iloc[:20, :].values.flatten())
    if '千円' in text:
        return 1 / 100000.0, '千円'
    if '百万円' in text:
        return 1 / 100.0, '百万円'
    if '億円' in text:
        return 1.0, '億円'
    if re.search(r'(?<!千)(?<!万)円', text):
        return 1 / 100000000.0, '円'
    # 現行の金額ファイルは原則千円。明示が読めない場合も千円として扱う。
    return 1 / 100000.0, '千円(推定)'


def old_value_column(df):
    """旧形式の金額列を特定。既知形式は0:A=部門,1:B=売買,8:I=金額。"""
    if df.shape[1] >= 9:
        return 8
    # 保険: 売り/買い行で数値が最も多い列
    scores = []
    for c in range(df.shape[1]):
        cnt = sum(pd.notna(num(v)) for v in df.iloc[:, c])
        scores.append((cnt, c))
    return max(scores)[1]


def parse_old_prime(df):
    """2026/9/28までの市場別シートを読む。JPX既知形式(A=部門,B=売買,I=金額)を優先。"""
    if df.empty:
        return pd.DataFrame(columns=['投資部門', '売り', '買い', '差引'])
    c_value = old_value_column(df)
    current = None
    values = {}
    for r in range(df.shape[0]):
        first = clean_text(df.iat[r, 0]) if df.shape[1] > 0 else ''
        canon = canonical_label(first)
        if canon:
            current = canon
        side = clean_text(df.iat[r, 1]) if df.shape[1] > 1 else ''
        if current and ('売り' in side or side.lower() == 'sell'):
            v = num(df.iat[r, c_value])
            if pd.notna(v):
                values.setdefault(current, {})['売り'] = v
        elif current and ('買い' in side or side.lower() == 'buy'):
            v = num(df.iat[r, c_value])
            if pd.notna(v):
                values.setdefault(current, {})['買い'] = v

    rows = []
    for investor, x in values.items():
        if '売り' in x and '買い' in x:
            rows.append([investor, x['売り'], x['買い'], x['買い'] - x['売り']])
    return pd.DataFrame(rows, columns=['投資部門', '売り', '買い', '差引'])


def find_prime_rows_new(df):
    rows = []
    for r in range(df.shape[0]):
        txt = ''.join(clean_text(x) for x in df.iloc[r, :min(df.shape[1], 8)].tolist())
        if 'プライム' in txt or 'Prime' in txt:
            rows.append(r)
    return rows


def parse_new_format(df):
    """2026/9/29以降の1シート集約形式。市場行=Prime、列見出し=投資部門/売買を復元。"""
    prime_rows = find_prime_rows_new(df)
    if not prime_rows:
        return pd.DataFrame(columns=['投資部門', '売り', '買い', '差引'])

    # 各列の上側見出しを連結し、どの投資部門・売買区分かを決める
    header_rows = max(0, min(prime_rows) - 8)
    records = []
    for r in prime_rows:
        for c in range(df.shape[1]):
            v = num(df.iat[r, c])
            if pd.isna(v):
                continue
            header = ''.join(clean_text(df.iat[rr, c]) for rr in range(header_rows, r))
            canon = canonical_label(header)
            if not canon:
                continue
            h = header.lower()
            metric = None
            if '売り' in h or 'sell' in h:
                metric = '売り'
            elif '買い' in h or 'buy' in h:
                metric = '買い'
            elif '差引' in h or 'balance' in h or 'net' in h:
                metric = '差引'
            if metric:
                records.append((canon, metric, v))

    if not records:
        return pd.DataFrame(columns=['投資部門', '売り', '買い', '差引'])
    p = pd.DataFrame(records, columns=['投資部門', 'metric', 'value'])
    p = p.pivot_table(index='投資部門', columns='metric', values='value', aggfunc='last').reset_index()
    for col in ['売り', '買い', '差引']:
        if col not in p:
            p[col] = np.nan
    p['差引'] = p['差引'].fillna(p['買い'] - p['売り'])
    return p[['投資部門', '売り', '買い', '差引']]


def validate_parsed(d):
    """誤列取得を弾く。売り買いが同桁で、差引=買い-売りであることを確認。"""
    if d.empty:
        return False, '投資部門を抽出できませんでした'
    check = d.dropna(subset=['売り', '買い']).copy()
    if check.empty:
        return False, '売り・買いを抽出できませんでした'
    if (check[['売り', '買い']] < 0).any().any():
        return False, '売買金額に負値が含まれています'
    # 売りと買いは通常同程度。明らかな片側ゼロ/誤列を排除。
    ratio = (check[['売り', '買い']].max(axis=1) / check[['売り', '買い']].min(axis=1).replace(0, np.nan))
    if ratio.dropna().gt(20).mean() > 0.25:
        return False, '売りと買いの桁が一致せず、列判定に失敗した可能性があります'
    calc = check['買い'] - check['売り']
    err = (calc - check['差引']).abs().fillna(0)
    if (err > np.maximum(1, calc.abs() * .001)).any():
        return False, '差引の検算に失敗しました'
    return True, ''


def parse_file(candidates):
    resolved_url, kind, sheets = read_excel_candidates(tuple(candidates))
    name, df = choose_prime_sheet(sheets)
    # unified fileはファイル名で判定、それ以外は旧形式
    is_new = 'stock_1_w_' in resolved_url.lower()
    parsed = parse_new_format(df) if is_new else parse_old_prime(df)
    ok, msg = validate_parsed(parsed)
    if not ok:
        raise ValueError(f'{name}: {msg}')
    multiplier, unit = detect_unit_multiplier(df)
    parsed = parsed.copy()
    for col in ['売り', '買い', '差引']:
        parsed[col] = parsed[col] * multiplier
    return resolved_url, name, unit, parsed


st.title('📊 JPX 投資部門別売買状況 — 東証プライム')
st.caption('週次・金額ベース / 売りはマイナス表示 / 差引 = 買い − 売り / iPhone対応 v1.8')

with st.sidebar:
    st.header('表示設定')
    weeks = st.slider('取得する直近週数', 4, 26, 12)
    st.caption('JPX週間データの金額Excelのみを読み込みます。')

try:
    files, link_diagnostics = discover_week_files(weeks)
except Exception as e:
    st.error(f'JPXページの取得に失敗しました: {e}')
    st.stop()

if not files:
    st.error('JPXの週次ブロックは確認できましたが、金額Excelを解決できませんでした。v1.8では金額PDFの実hrefから同一ディレクトリのExcel URLも復元し、Excel本体を検証してから読み込みます。')
    if link_diagnostics:
        with st.expander('リンク診断', expanded=True):
            for d in link_diagnostics[:6]:
                st.write(d['text'])
                st.code('\n'.join(d['candidates']) if d['candidates'] else '(候補URLなし)')
    st.stop()

rows, errors = [], []
progress = st.progress(0, text='JPXデータを取得中…')
for i, rec in enumerate(files):
    try:
        resolved_url, sheet, unit, d = parse_file(rec['candidates'])
        for _, r in d.iterrows():
            rows.append({
                '週開始': rec['start'], '週終了': rec['end'],
                '投資部門': r['投資部門'], '売り': r['売り'], '買い': r['買い'], '差引': r['差引'],
                'source': resolved_url, 'sheet': sheet, 'source_unit': unit,
            })
    except Exception as e:
        errors.append((rec['text'], '\n'.join(rec['candidates'][:4]), str(e)))
    progress.progress((i + 1) / max(1, len(files)), text=f'JPXデータを取得中… {i + 1}/{len(files)}')
progress.empty()

all_df = pd.DataFrame(rows)
if all_df.empty:
    st.error('プライム市場のデータを解析できませんでした。下の「データ取得状況」を確認してください。')
    if errors:
        with st.expander('データ取得状況 / エラー', expanded=True):
            st.code('\n\n'.join(f'{t}\n{u}\n{e}' for t, u, e in errors[:10]))
    st.stop()

all_df = all_df.sort_values(['週終了', '投資部門']).drop_duplicates(['週終了', '投資部門'], keep='last')
latest_end = all_df['週終了'].max()
latest = all_df[all_df['週終了'].eq(latest_end)].copy()

# 週の解析件数を見える化（誤データに気づきやすくする）
parsed_weeks = all_df['週終了'].nunique()
st.caption(f'正常解析: {parsed_weeks}週 / 取得候補: {len(files)}週' + (f' / エラー: {len(errors)}週' if errors else ''))

st.subheader(f"直近週：{latest_end.month}/{latest_end.day:02d} 終了")
cols = st.columns(4)
for c, name in zip(cols, ['海外投資家', '個人', '信託銀行', '事業法人']):
    s = latest.loc[latest['投資部門'].eq(name), '差引']
    if s.empty:
        c.metric(name, '—')
    else:
        v = float(s.iloc[0])
        c.metric(name, f'{v:+,.0f} 億円', '買い越し' if v > 0 else ('売り越し' if v < 0 else '均衡'))

st.divider()
st.subheader('投資家別 売買推移')
st.caption('買い＝プラスの棒、売り＝マイナスの棒、差引＝折れ線。月次は週次公表値を月ごとに合計します。')

available = [x for x in MAJOR + ['生保・損保', '都銀・地銀等', 'その他金融機関', 'その他法人等', '自己計', '委託計'] if x in all_df['投資部門'].unique()]
if not available:
    available = sorted(all_df['投資部門'].unique())

c1, c2 = st.columns([2, 1])
with c1:
    investor = st.selectbox('投資家を選択', available, index=available.index('海外投資家') if '海外投資家' in available else 0)
with c2:
    period = st.radio('集計単位', ['週次', '月次'], horizontal=True)

one = all_df[all_df['投資部門'].eq(investor)].copy().sort_values('週終了')
one['売り表示'] = -one['売り'].abs()
one['買い表示'] = one['買い'].abs()

if period == '月次':
    one['月'] = one['週終了'].dt.to_period('M')
    detail = one.groupby('月', as_index=False)[['売り表示', '買い表示', '差引']].sum()
    detail['期間'] = detail['月'].map(lambda p: f'{p.year}/{p.month:02d}')
else:
    detail = one[['週終了', '売り表示', '買い表示', '差引']].copy()
    detail['期間'] = detail['週終了'].map(lambda d: f'{d.month}/{d.day:02d}')

fig = go.Figure()
fig.add_trace(go.Bar(
    x=detail['期間'], y=detail['売り表示'], name='売り',
    hovertemplate='%{x}<br>売り %{customdata:,.0f} 億円<extra></extra>', customdata=detail['売り表示'].abs(),
))
fig.add_trace(go.Bar(
    x=detail['期間'], y=detail['買い表示'], name='買い',
    hovertemplate='%{x}<br>買い %{y:,.0f} 億円<extra></extra>',
))
fig.add_trace(go.Scatter(
    x=detail['期間'], y=detail['差引'], name='差引', mode='lines+markers', yaxis='y2',
    hovertemplate='%{x}<br>差引 %{y:+,.0f} 億円<extra></extra>',
))
fig.update_layout(
    title=f'{investor} — {period}', barmode='relative', height=420,
    margin=dict(l=8, r=8, t=48, b=56),
    legend=dict(orientation='h', yanchor='top', y=-0.12, xanchor='center', x=.5),
    xaxis=dict(title='', tickangle=0, automargin=True, nticks=min(8, max(2, len(detail)))),
    yaxis=dict(title='売買金額（億円）', zeroline=True, zerolinewidth=1, tickformat=',~s'),
    yaxis2=dict(title='差引（億円）', overlaying='y', side='right', showgrid=False, zeroline=True, tickformat=',~s'),
    hovermode='x unified',
)
st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False, 'responsive': True, 'scrollZoom': False})

if not detail.empty:
    m1, m2, m3 = st.columns(3)
    m1.metric('期間内 買い', f"{detail['買い表示'].sum():,.0f} 億円")
    m2.metric('期間内 売り', f"{detail['売り表示'].abs().sum():,.0f} 億円")
    m3.metric('期間内 差引', f"{detail['差引'].sum():+,.0f} 億円")

st.divider()
st.subheader('直近週の投資部門別差引')
rank = latest[['投資部門', '差引']].dropna().sort_values('差引', ascending=False)
fig2 = go.Figure(go.Bar(
    y=rank['投資部門'], x=rank['差引'], orientation='h',
    hovertemplate='%{y}<br>%{x:+,.0f} 億円<extra></extra>',
))
fig2.update_layout(height=max(350, 31 * len(rank)), margin=dict(l=8, r=8, t=20, b=35), xaxis_title='買い越し / 売り越し（億円）', yaxis_title='')
st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False, 'responsive': True})

with st.expander('直近週の数値表'):
    tbl = latest[['投資部門', '売り', '買い', '差引']].copy().sort_values('差引', ascending=False)
    for col in ['売り', '買い', '差引']:
        tbl[col] = tbl[col].map(lambda x: f'{x:,.0f}')
    st.dataframe(tbl, hide_index=True, use_container_width=True)

with st.expander('データ取得状況 / エラー'):
    st.write(f'取得候補: {len(files)}週 / 正常解析: {parsed_weeks}週 / エラー: {len(errors)}週')
    if errors:
        st.code('\n\n'.join(f'{t}\n{u}\n  {e}' for t, u, e in errors[:12]))

st.caption('出典: 日本取引所グループ（JPX）「投資部門別売買状況」株式・週間。対象は東証プライム。非公式アプリです。')
