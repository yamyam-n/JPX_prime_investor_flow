import io
import re
from datetime import datetime
from urllib.parse import urljoin

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title='JPX プライム 投資部門別売買', page_icon='📊', layout='wide', initial_sidebar_state='collapsed')

# iPhone / mobile responsive tuning
st.markdown("""
<style>
.block-container {padding-top: 1.0rem; padding-bottom: 2rem; max-width: 1200px;}
[data-testid="stMetricValue"] {font-size: 1.55rem;}
[data-testid="stPlotlyChart"] {width: 100% !important;}
@media (max-width: 700px) {
  .block-container {padding-left: .65rem !important; padding-right: .65rem !important; padding-top: .55rem !important;}
  h1 {font-size: 1.55rem !important; line-height: 1.25 !important;}
  h2, h3 {font-size: 1.18rem !important;}
  [data-testid="stHorizontalBlock"] {flex-wrap: wrap !important; gap: .35rem !important;}
  [data-testid="column"] {min-width: 100% !important; width: 100% !important; flex: 1 1 100% !important;}
  [data-testid="stMetric"] {padding: .45rem .6rem !important;}
  [data-testid="stMetricValue"] {font-size: 1.35rem !important;}
  [data-testid="stDataFrame"] {font-size: .78rem !important;}
  .stPlotlyChart {overflow-x: hidden !important;}
}
</style>
""", unsafe_allow_html=True)

JPX_WEEKLY = 'https://www.jpx.co.jp/markets/statistics-equities/investor-type/index.html'
HEADERS = {'User-Agent': 'Mozilla/5.0 (compatible; JPXInvestorFlow/1.0)'}
MAJOR = ['海外投資家', '個人', '信託銀行', '事業法人', '投資信託', '証券会社']
ALIASES = {
    '海外投資家': ['海外投資家', '外国人', '外国法人等'],
    '個人': ['個人', '個人・その他'],
    '信託銀行': ['信託銀行'],
    '事業法人': ['事業法人'],
    '投資信託': ['投資信託'],
    '証券会社': ['証券会社'],
    '生保・損保': ['生保・損保', '生損保'],
    '都銀・地銀等': ['都銀・地銀等', '都銀・地銀'],
    'その他金融機関': ['その他金融機関'],
    'その他法人等': ['その他法人等', 'その他法人'],
    '自己計': ['自己計', '自己'],
    '委託計': ['委託計', '委託'],
}


def clean_text(x):
    if pd.isna(x):
        return ''
    return re.sub(r'\s+', '', str(x)).replace('\u3000', '')


def num(x):
    if pd.isna(x): return np.nan
    if isinstance(x, (int, float, np.number)): return float(x)
    s = str(x).replace(',', '').replace('△', '-').replace('▲', '-').strip()
    s = re.sub(r'[^0-9.\-]', '', s)
    if s in ('', '-', '.', '-.'): return np.nan
    try: return float(s)
    except: return np.nan


@st.cache_data(ttl=3600, show_spinner=False)
def discover_excel_links():
    r = requests.get(JPX_WEEKLY, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, 'html.parser')
    links = []
    for a in soup.find_all('a', href=True):
        href = urljoin(JPX_WEEKLY, a['href'])
        text = clean_text(a.get_text(' ', strip=True))
        if re.search(r'\.xlsx?(?:\?|$)', href, re.I):
            links.append((text, href))
    # 金額ファイルを優先。新形式 stock_1_w_*.xlsx も対象。
    money = [(t,u) for t,u in links if ('val' in u.lower() or 'stock_1_w_' in u.lower())]
    return money or links


def dates_from_text(text):
    """URL/リンク文言から8桁日付を拾う。"""
    vals = re.findall(r'(?<!\d)(20\d{6})(?!\d)', str(text))
    out = []
    for v in vals:
        try:
            out.append(pd.to_datetime(v, format='%Y%m%d'))
        except Exception:
            pass
    return out


def week_label_from_url(url, fallback=''):
    # ファイル名をそのままX軸に出さず、期間末の日付を短く表示
    dates = dates_from_text(url) + dates_from_text(fallback)
    if dates:
        d = max(dates)
        return f'{d.month}/{d.day:02d}'
    # "8/3" や "08/03" のような表記があれば末尾を採用
    md = re.findall(r'(?<!\d)(\d{1,2})[/-](\d{1,2})(?!\d)', str(fallback))
    if md:
        m, d = md[-1]
        return f'{int(m)}/{int(d):02d}'
    return ''


@st.cache_data(ttl=3600, show_spinner=False)
def read_excel_url(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    bio = io.BytesIO(r.content)
    engine = 'xlrd' if url.lower().split('?')[0].endswith('.xls') else 'openpyxl'
    return pd.read_excel(bio, sheet_name=None, header=None, engine=engine)


def choose_prime_sheet(sheets):
    # 現行形式: プライム市場の専用シートを探す
    for name, df in sheets.items():
        key = clean_text(name)
        if 'プライム' in key or 'Prime' in str(name):
            return name, df
    # シート名で見つからない場合、セル内に「プライム」を含むシート
    for name, df in sheets.items():
        txt = ''.join(clean_text(v) for v in df.head(30).astype(object).values.flatten())
        if 'プライム' in txt:
            return name, df
    # fallback: 最大シート
    name = max(sheets, key=lambda k: sheets[k].shape[0] * sheets[k].shape[1])
    return name, sheets[name]


def canonical_label(text):
    t = clean_text(text)
    for canon, aliases in ALIASES.items():
        if any(clean_text(a) in t for a in aliases):
            return canon
    return None


def parse_old_prime(df):
    """現行（〜2026/9/28）横持ちExcelを、ラベル探索で読み取る。"""
    arr = df.astype(object).copy()
    out = []
    # 投資家ラベルのセルを探し、近傍の「売り」「買い」「差引」または数値列を読む
    for r in range(arr.shape[0]):
        for c in range(arr.shape[1]):
            canon = canonical_label(arr.iat[r,c])
            if not canon:
                continue
            # 右方向最大12列、下方向最大4行を候補にする
            block = arr.iloc[r:min(r+5, arr.shape[0]), c:min(c+13, arr.shape[1])]
            vals = []
            tagged = {}
            for rr in range(block.shape[0]):
                rowtxt = [clean_text(x) for x in block.iloc[rr].tolist()]
                nums = [num(x) for x in block.iloc[rr].tolist()]
                nnums = [x for x in nums if pd.notna(x)]
                joined = ''.join(rowtxt)
                if '売' in joined and nnums: tagged['sell'] = nnums[-1]
                if '買' in joined and nnums: tagged['buy'] = nnums[-1]
                if ('差引' in joined or '差引き' in joined) and nnums: tagged['net'] = nnums[-1]
                vals.extend(nnums)
            if 'buy' in tagged and 'sell' in tagged:
                sell, buy = tagged['sell'], tagged['buy']
                net = tagged.get('net', buy-sell)
                out.append((canon, sell, buy, net)); continue
            # 1行形式: ラベル右側の数値から売・買・差引を推定
            rownums = [num(x) for x in arr.iloc[r, c+1:min(c+15, arr.shape[1])].tolist()]
            rownums = [x for x in rownums if pd.notna(x)]
            if len(rownums) >= 3:
                # 売・買・差引の並びであることが多い。最後3つを利用。
                sell, buy, netv = rownums[-3:]
                if abs((buy-sell)-netv) <= max(10, abs(netv)*0.05):
                    out.append((canon, sell, buy, netv))
    if not out:
        return pd.DataFrame(columns=['投資部門','売り','買い','差引'])
    res = pd.DataFrame(out, columns=['投資部門','売り','買い','差引'])
    # 同じ部門が複数ヒットした場合、売買額が最大の行を採用
    res['scale'] = res['売り'].abs() + res['買い'].abs()
    res = res.sort_values('scale').drop_duplicates('投資部門', keep='last').drop(columns='scale')
    return res


def parse_new_format(df):
    """2026/9/29以降の1シート集約形式を汎用的に読む。"""
    # プライムを含む行を中心に、列ヘッダを上方向から復元する
    prime_rows = []
    for r in range(df.shape[0]):
        joined = ''.join(clean_text(x) for x in df.iloc[r].tolist())
        if 'プライム' in joined:
            prime_rows.append(r)
    records = []
    for r in prime_rows:
        for c in range(df.shape[1]):
            v = num(df.iat[r,c])
            if pd.isna(v): continue
            header = ''.join(clean_text(df.iat[rr,c]) for rr in range(max(0,r-5), r))
            canon = canonical_label(header)
            if not canon: continue
            # ヘッダに差引があれば直接採用。売/買は別辞書にためる
            metric = 'net' if '差引' in header else ('buy' if '買' in header else ('sell' if '売' in header else None))
            if metric:
                records.append((canon, metric, v))
    if not records:
        return pd.DataFrame(columns=['投資部門','売り','買い','差引'])
    d = pd.DataFrame(records, columns=['投資部門','metric','value']).pivot_table(index='投資部門', columns='metric', values='value', aggfunc='last').reset_index()
    d = d.rename(columns={'sell':'売り','buy':'買い','net':'差引'})
    for col in ['売り','買い','差引']:
        if col not in d: d[col] = np.nan
    d['差引'] = d['差引'].fillna(d['買い']-d['売り'])
    return d[['投資部門','売り','買い','差引']]


def parse_file(url):
    sheets = read_excel_url(url)
    name, df = choose_prime_sheet(sheets)
    parsed = parse_new_format(df) if ('stock_1_w_' in url.lower() or len(sheets)==1) else parse_old_prime(df)
    return name, parsed


def to_oku(x):
    # JPX金額Excelの基本単位は千円であるケースを想定。桁から自動判定。
    # 1億円 = 100,000千円。既に億円級ならそのまま。
    if pd.isna(x): return np.nan
    ax = abs(x)
    return x / 100000.0 if ax > 1000000 else x


st.title('📊 JPX 投資部門別売買状況 — 東証プライム')
st.caption('週次・金額ベース / 差引 = 買い − 売り / JPX公表資料から取得 / iPhone対応 v1.2')

with st.sidebar:
    st.header('表示設定')
    weeks = st.slider('取得する直近週数', 4, 26, 12)
    selected = st.multiselect('表示する投資部門', MAJOR, default=['海外投資家','個人','信託銀行','事業法人','投資信託'])
    st.caption('JPXの週間公表は通常、第4営業日15:30。')

try:
    links = discover_excel_links()
except Exception as e:
    st.error(f'JPXページの取得に失敗しました: {e}')
    st.stop()

# 同じURLを除去し、ページ掲載順の新しいものから使う
seen = set(); unique=[]
for text,url in links:
    if url not in seen:
        seen.add(url); unique.append((text,url))
links = unique[:weeks]

rows=[]; errors=[]
progress = st.progress(0, text='JPXデータを取得中…')
for i,(text,url) in enumerate(links):
    try:
        sheet, d = parse_file(url)
        label = week_label_from_url(url, text) or f'週{i+1}'
        for _,r in d.iterrows():
            rows.append({'週':label,'投資部門':r['投資部門'],'売り':to_oku(r['売り']),'買い':to_oku(r['買い']),'差引':to_oku(r['差引']),'source':url})
    except Exception as e:
        errors.append((url,str(e)))
    progress.progress((i+1)/max(1,len(links)), text=f'JPXデータを取得中… {i+1}/{len(links)}')
progress.empty()

all_df = pd.DataFrame(rows)
if all_df.empty:
    st.error('データを解析できませんでした。JPXのExcelレイアウト変更の可能性があります。')
    if errors: st.code('\n'.join(f'{u} :: {e}' for u,e in errors[:5]))
    st.stop()

# ページ順を保持
week_order=[]
for t,u in links:
    candidates = all_df.loc[all_df.source.eq(u),'週'].unique().tolist()
    week_order += [x for x in candidates if x not in week_order]
week_order = list(reversed(week_order))
all_df['週'] = pd.Categorical(all_df['週'], categories=week_order, ordered=True)
all_df = all_df.sort_values(['週','投資部門'])

latest_week = week_order[-1]
latest = all_df[all_df['週']==latest_week].copy().sort_values('差引', ascending=False)

# KPI
st.subheader(f'直近週：{latest_week}')
kpis=[]
for name in ['海外投資家','個人','信託銀行','事業法人']:
    s=latest.loc[latest['投資部門'].eq(name),'差引']
    val=s.iloc[0] if len(s) else np.nan
    kpis.append((name,val))
cols=st.columns(4)
for c,(name,val) in zip(cols,kpis):
    if pd.isna(val): c.metric(name,'—')
    else: c.metric(name,f'{val:+,.0f} 億円')


# --- 投資家別 売買グラフ -------------------------------------------------
st.divider()
st.subheader('投資家別 売買推移')
st.caption('売りはマイナス、買いはプラス、差引は折れ線で表示します。')

available_investors = [x for x in ALIASES.keys() if x in all_df['投資部門'].astype(str).unique().tolist()]
if not available_investors:
    available_investors = sorted(all_df['投資部門'].astype(str).unique().tolist())

g1, g2 = st.columns([2, 1])
with g1:
    investor_pick = st.selectbox(
        '投資家を選択',
        available_investors,
        index=available_investors.index('海外投資家') if '海外投資家' in available_investors else 0,
        key='investor_detail_pick',
    )
with g2:
    period_pick = st.radio('集計単位', ['週次', '月次'], horizontal=True, key='period_detail_pick')

one = all_df[all_df['投資部門'].astype(str).eq(investor_pick)].copy()
one['売り表示'] = -one['売り'].abs()
one['買い表示'] = one['買い'].abs()

# JPXリンクのURLやラベルから週の終了日をできるだけ復元する
url_to_order = {u:i for i,(_,u) in enumerate(reversed(links))}
one['_order'] = one['source'].map(url_to_order).fillna(0)

def extract_date(row):
    # URLまたは表示ラベルから週末日を復元
    dates = dates_from_text(str(row.get('source','')))
    if dates:
        return max(dates)
    label = str(row.get('週',''))
    m = re.search(r'(?<!\d)(\d{1,2})/(\d{1,2})(?!\d)', label)
    if m:
        # 年が取れない場合は現在年を使用（表示順は_source順で補助）
        return pd.Timestamp(datetime.now().year, int(m.group(1)), int(m.group(2)))
    return pd.NaT

one['_date'] = one.apply(extract_date, axis=1)
one = one.sort_values(['_date','_order'], na_position='last')

if period_pick == '月次':
    valid = one.dropna(subset=['_date']).copy()
    if not valid.empty:
        valid['月'] = valid['_date'].dt.to_period('M').astype(str)
        detail = valid.groupby('月', as_index=False)[['売り表示','買い表示','差引']].sum()
        detail = detail.rename(columns={'月':'期間'})
    else:
        # 日付が取れない場合は4週単位の簡易集計
        tmp = one.reset_index(drop=True).copy()
        tmp['_mgrp'] = np.arange(len(tmp)) // 4
        detail = tmp.groupby('_mgrp', as_index=False)[['売り表示','買い表示','差引']].sum()
        detail['期間'] = detail['_mgrp'].map(lambda x: f'4週集計 {x+1}')
else:
    detail = one[['週','売り表示','買い表示','差引','_date']].copy().rename(columns={'週':'期間'})
    # iPhoneで読める短い週ラベル
    detail['期間'] = detail.apply(lambda r: f"{r['_date'].month}/{r['_date'].day:02d}" if pd.notna(r['_date']) else str(r['期間']), axis=1)
    detail = detail.drop(columns=['_date'])

# 添付イメージに近い、上下バー＋差引線
fig_detail = go.Figure()
fig_detail.add_trace(go.Bar(
    x=detail['期間'], y=detail['売り表示'], name='Sales',
    hovertemplate='%{x}<br>売り: %{customdata:,.0f} 億円<extra></extra>',
    customdata=detail['売り表示'].abs(),
))
fig_detail.add_trace(go.Bar(
    x=detail['期間'], y=detail['買い表示'], name='Purchases',
    hovertemplate='%{x}<br>買い: %{y:,.0f} 億円<extra></extra>',
))
fig_detail.add_trace(go.Scatter(
    x=detail['期間'], y=detail['差引'], name='Balance', mode='lines+markers',
    yaxis='y2', hovertemplate='%{x}<br>差引: %{y:+,.0f} 億円<extra></extra>',
))
fig_detail.update_layout(
    title=f'{investor_pick} — {period_pick}',
    barmode='relative',
    height=430,
    margin=dict(l=6,r=6,t=52,b=72),
    legend=dict(orientation='h', yanchor='top', y=-0.16, xanchor='center', x=0.5),
    xaxis=dict(title='', tickangle=0, automargin=True, nticks=min(8, max(2, len(detail)))),
    yaxis=dict(title='売買金額（億円）', zeroline=True, zerolinewidth=1),
    yaxis2=dict(title='差引（億円）', overlaying='y', side='right', showgrid=False, zeroline=False),
)
st.plotly_chart(fig_detail, use_container_width=True, config={'displayModeBar': False, 'responsive': True, 'scrollZoom': False})

if not detail.empty:
    a,b,c = st.columns(3)
    a.metric('期間内 買い', f"{detail['買い表示'].sum():,.0f} 億円")
    b.metric('期間内 売り', f"{detail['売り表示'].abs().sum():,.0f} 億円")
    c.metric('期間内 差引', f"{detail['差引'].sum():+,.0f} 億円")

st.divider()

plot_df = all_df[all_df['投資部門'].isin(selected)].copy()
fig = px.bar(plot_df, x='週', y='差引', color='投資部門', barmode='group',
             labels={'差引':'買い越し / 売り越し（億円）'},
             title='投資部門別 ネット売買（週次）')
fig.add_hline(y=0, line_width=1)
fig.update_layout(height=400, legend_title_text='', margin=dict(l=6,r=6,t=50,b=55), legend=dict(orientation='h', y=-0.18), xaxis=dict(tickangle=0, nticks=8, automargin=True))
st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False, 'responsive': True, 'scrollZoom': False})

cum = plot_df.copy()
cum['累積'] = cum.groupby('投資部門', observed=True)['差引'].cumsum()
fig2 = px.line(cum, x='週', y='累積', color='投資部門', markers=True,
               labels={'累積':'累積ネット売買（億円）'}, title='期間累積の買い越し・売り越し')
fig2.add_hline(y=0, line_width=1)
fig2.update_layout(height=390, legend_title_text='', margin=dict(l=6,r=6,t=50,b=55), legend=dict(orientation='h', y=-0.18), xaxis=dict(tickangle=0, nticks=8, automargin=True))
st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False, 'responsive': True, 'scrollZoom': False})

c1,c2=st.columns([1,1])
with c1:
    st.subheader('直近週ランキング')
    show=latest[['投資部門','差引']].dropna().copy()
    show['差引（億円）']=show['差引'].map(lambda x:f'{x:+,.0f}')
    st.dataframe(show[['投資部門','差引（億円）']], hide_index=True, use_container_width=True)
with c2:
    st.subheader('売り・買い・差引')
    tbl=latest[['投資部門','売り','買い','差引']].copy()
    for col in ['売り','買い','差引']:
        tbl[col]=tbl[col].map(lambda x:'' if pd.isna(x) else f'{x:,.0f}')
    st.dataframe(tbl, hide_index=True, use_container_width=True)

with st.expander('データ取得状況 / エラー'):
    st.write(f'取得候補: {len(links)}週 / 解析エラー: {len(errors)}件')
    if errors:
        st.code('\n'.join(f'{u}\n  {e}' for u,e in errors))

st.caption('出典: 日本取引所グループ（JPX）「投資部門別売買状況」株式・週間。非公式アプリです。')
