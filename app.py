# -*- coding: utf-8 -*-
"""
============================================================
 Streamlit 版：整理任天堂 Nintendo Switch 銷量表 + 圖像化 + 下載 Excel
============================================================
執行方式（本機）：
    pip install -r requirements.txt
    streamlit run app.py

部署到 Streamlit Community Cloud（GitHub）：
    1. 把本資料夾推上 GitHub（app.py、requirements.txt、README.md）
    2. 到 https://streamlit.io/cloud 用 GitHub 帳號登入
    3. New app -> 選 repo 與 branch，Main file 填 app.py -> Deploy
============================================================
"""

import io
import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import streamlit as st

# ---------- 頁面設定 ----------
st.set_page_config(page_title="Nintendo Switch 銷量整理", page_icon="🎮", layout="wide")

# 設定中文字型，避免圖表中文亂碼
_candidates = ['SimHei', 'Noto Sans CJK TC', 'Noto Sans CJK SC',
               'PingFang SC', 'Microsoft JhengHei', 'DejaVu Sans']
_avail = {f.name for f in fm.fontManager.ttflist}
_use = next((f for f in _candidates if f in _avail), 'DejaVu Sans')
matplotlib.rcParams['font.sans-serif'] = [_use]
matplotlib.rcParams['axes.unicode_minus'] = False


# ============================================================
# 資料解析（從上傳的 xlsx 讀出「類別 / 地區 / 期間 / 銷量」長表）
# ============================================================
@st.cache_data(show_spinner="解析資料中…")
def parse_excel(raw_bytes: bytes) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # 找到表頭列（含有 'FY3/2017' 的那一列）
    header_idx = None
    for i, r in enumerate(rows):
        if any(v is not None and str(v).startswith('FY3/') for v in r):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('找不到表頭列（FY3/2017）')
    header = rows[header_idx]

    # 欄位索引 -> 期間名稱（年份 + Life to date）
    col_period = {}
    for c, v in enumerate(header):
        if v is None:
            continue
        s = str(v).strip()
        if s.startswith('FY3/'):
            col_period[c] = s.replace('FY3/', 'FY')
        elif 'Life' in s or 'date' in s.lower():
            col_period[c] = 'Life to date'

    records = []
    category = None
    for r in rows[header_idx + 1:]:
        if r is None or all(v is None for v in r):
            continue
        if r[0] is not None and str(r[0]).strip():
            category = str(r[0]).strip()
        region = r[3]
        if region is None:
            continue
        region = str(region).strip()
        for c, period in col_period.items():
            v = r[c]
            records.append({
                '類別 Category': category,
                '地區 Region': region,
                '期間 Period': period,
                '銷量(萬台) Units': float(v) if v is not None else np.nan,
            })
    return pd.DataFrame(records)


# ============================================================
# 樞紐表與累計摘要
# ============================================================
def make_pivot(df, category, region='Total', drop_life=True):
    sub = df[(df['類別 Category'] == category) &
             (df['地區 Region'] == region)].copy()
    if drop_life:
        sub = sub[sub['期間 Period'] != 'Life to date']
    piv = sub.pivot_table(index='期間 Period', columns='地區 Region',
                          values='銷量(萬台) Units', aggfunc='sum')
    piv = piv.reindex([p for p in df['期間 Period'].unique() if p != 'Life to date'])
    return piv.dropna(how='all')


def life_summary(df, category):
    sub = df[(df['類別 Category'] == category) &
             (df['期間 Period'] == 'Life to date')]
    return sub.set_index('地區 Region')['銷量(萬台) Units']


# ============================================================
# 產生圖表（回傳 figure 清單 + 檔名清單）
# ============================================================
def _placeholder(text, size=(9, 4.8)):
    """資料不足時顯示的說明圖，避免崩潰。"""
    fig, ax = plt.subplots(figsize=size, dpi=110)
    ax.text(0.5, 0.5, text, ha='center', va='center', fontsize=14, color='gray')
    ax.axis('off')
    return fig


def _has_data(piv):
    """樞紐表是否有可用數值。"""
    if piv is None or piv.empty:
        return False
    num = piv.select_dtypes(include='number')
    return not num.empty and float(num.abs().sum().sum()) > 0


def build_charts(hard_pivot, soft_pivot, models, hard_life):
    figs = []
    names = []

    # 圖1：硬體各區年度趨勢
    if _has_data(hard_pivot):
        fig, ax = plt.subplots(figsize=(9, 4.8), dpi=110)
        hard_pivot.plot(ax=ax, marker='o')
        ax.set_title('Nintendo 硬體主機各區年度銷量趨勢（萬台）')
        ax.set_xlabel('期間'); ax.set_ylabel('銷量(萬台)')
        ax.grid(True, alpha=0.3); ax.legend(title='地區')
        fig.tight_layout()
    else:
        fig = _placeholder('Hardware 資料不足')
    figs.append(fig); names.append('chart_hardware.png')

    # 圖2：軟體各區年度趨勢
    if _has_data(soft_pivot):
        fig, ax = plt.subplots(figsize=(9, 4.8), dpi=110)
        soft_pivot.plot(ax=ax, marker='o')
        ax.set_title('Nintendo 軟體各區年度銷量趨勢（萬台）')
        ax.set_xlabel('期間'); ax.set_ylabel('銷量(萬台)')
        ax.grid(True, alpha=0.3); ax.legend(title='地區')
        fig.tight_layout()
    else:
        fig = _placeholder('Software 資料不足')
    figs.append(fig); names.append('chart_software.png')

    # 圖3：三大機型全球年度銷量
    if _has_data(models):
        fig, ax = plt.subplots(figsize=(9, 4.8), dpi=110)
        models.plot(ax=ax, marker='o')
        ax.set_title('Nintendo Switch 三大機型全球年度銷量（萬台）')
        ax.set_xlabel('期間'); ax.set_ylabel('銷量(萬台)')
        ax.grid(True, alpha=0.3); ax.legend(title='機型')
        fig.tight_layout()
    else:
        fig = _placeholder('三大機型資料不足')
    figs.append(fig); names.append('chart_models.png')

    # 圖4：硬體 vs 軟體 全球年度銷量
    h_total = hard_pivot['Total'] if (_has_data(hard_pivot) and 'Total' in hard_pivot.columns) else None
    s_total = soft_pivot['Total'] if (_has_data(soft_pivot) and 'Total' in soft_pivot.columns) else None
    if h_total is not None and s_total is not None:
        fig, ax = plt.subplots(figsize=(9, 4.8), dpi=110)
        x = hard_pivot.index
        ax.bar(x, h_total, label='硬體 Hardware', alpha=0.7)
        ax.bar(x, s_total, label='軟體 Software', alpha=0.7)
        ax.set_ylabel('銷量(萬台)'); ax.set_title('Nintendo 全球硬體 vs 軟體 年度銷量')
        ax.legend(); ax.grid(True, axis='y', alpha=0.3)
        plt.setp(ax.get_xticklabels(), rotation=45)
        fig.tight_layout()
    else:
        fig = _placeholder('硬體或軟體資料不足')
    figs.append(fig); names.append('chart_hw_vs_sw.png')

    # 圖5：累計銷量占比（Life to date，硬體）
    pie_data = hard_life.drop('Total', errors='ignore').dropna() if hard_life is not None else pd.Series(dtype=float)
    if len(pie_data) >= 2 and float(pie_data.abs().sum()) > 0:
        fig, ax = plt.subplots(figsize=(6, 6), dpi=110)
        pie_data.plot.pie(ax=ax, autopct='%1.1f%%', startangle=90, counterclock=False)
        ax.set_title('Nintendo 硬體累計銷量地區占比（Life to date）')
        ax.set_ylabel('')
        fig.tight_layout()
    else:
        fig = _placeholder('地區資料不足以繪製占比圖', size=(6, 6))
    figs.append(fig); names.append('chart_pie.png')

    return figs, names


# ============================================================
# 產生整理後的 Excel（含多分頁 + 嵌入圖表），回傳 BytesIO
# ============================================================
def build_excel(df, hard_pivot, soft_pivot, models, hard_life, soft_life, figs, names):
    wb_out = Workbook()

    # Sheet 1：整齊長表
    ws1 = wb_out.active
    ws1.title = '整理後資料'
    ws1.append(df.columns.tolist())
    for _, row in df.iterrows():
        ws1.append(row.tolist())
    for col in ws1.columns:
        ws1.column_dimensions[col[0].column_letter].width = 18

    # Sheet 2~4：樞紐表
    for title, data in [('硬體-各區年度', hard_pivot),
                        ('軟體-各區年度', soft_pivot),
                        ('三大機型-全球年度', models)]:
        ws = wb_out.create_sheet(title)
        ws.append(['期間'] + list(data.columns))
        for idx, row in data.iterrows():
            ws.append([idx] + list(row))

    # Sheet 5：累計摘要
    ws5 = wb_out.create_sheet('累計摘要-LifeToDate')
    ws5.append(['地區', '硬體(萬台)', '軟體(萬台)'])
    for reg in hard_life.index:
        ws5.append([reg, hard_life.get(reg, np.nan), soft_life.get(reg, np.nan)])

    # Sheet 6：圖表總覽（把圖先存成 PNG 再嵌入）
    ws6 = wb_out.create_sheet('圖表總覽')
    ws6['A1'] = 'Nintendo Switch 銷量分析圖表'
    ws6['A1'].font = openpyxl.styles.Font(size=14, bold=True)
    for i, (fig, name) in enumerate(zip(figs, names), start=1):
        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        try:
            img = XLImage(buf)
            img.anchor = f'A{(i - 1) * 32 + 3}'
            ws6.add_image(img)
        except Exception as e:
            st.warning(f'嵌入圖片失敗：{name} -> {e}')

    bio = io.BytesIO()
    wb_out.save(bio)
    bio.seek(0)
    return bio


# ============================================================
# 網頁 UI
# ============================================================
st.title('🎮 Nintendo Switch 銷量表整理工具')
st.caption('上傳任天堂銷量 Excel → 自動整理、畫圖、下載整理後的 Excel（Streamlit / GitHub 版）')

# 側欄：上傳
with st.sidebar:
    st.header('① 上傳檔案')
    uploaded = st.file_uploader('上傳 NSe2503.xlsx', type=['xlsx'],
                                help='選擇你的銷量表 Excel 檔')
    st.markdown('---')
    st.caption('流程：上傳 → 整理 → 看圖 → 下載 Excel')

if uploaded is None:
    st.info('👈 請先在上方側欄上傳你的 Excel 檔（例如 NSe2503.xlsx）。')
    st.stop()

# ---------- 解析 ----------
df = parse_excel(uploaded.getvalue())

# ---------- 樞紐 / 摘要 ----------
hard_pivot = make_pivot(df, 'Hardware')
soft_pivot = make_pivot(df, 'Software')
base_switch = make_pivot(df, 'of which Nintendo Switch')
oled = make_pivot(df, 'Nintendo Switch – OLED Model')
lite = make_pivot(df, 'Nintendo Switch Lite')

def _total_series(piv):
    """安全取樞紐表的 'Total' 欄，缺列時回傳空序列，避免崩潰。"""
    if 'Total' in piv.columns:
        return piv['Total']
    return pd.Series(dtype=float)

models = pd.DataFrame({
    'Nintendo Switch': _total_series(base_switch),
    'OLED Model': _total_series(oled),
    'Lite': _total_series(lite),
})
models.index.name = '期間 Period'

hard_life = life_summary(df, 'Hardware')
soft_life = life_summary(df, 'Software')

# ---------- 總覽指標 ----------
st.subheader('📊 數據總覽')
c1, c2, c3, c4 = st.columns(4)
c1.metric('資料筆數', f'{len(df):,}')
c2.metric('硬體累計銷量(萬台)', f"{hard_life.get('Total', np.nan):,.0f}")
c3.metric('軟體累計銷量(萬台)', f"{soft_life.get('Total', np.nan):,.0f}")
c4.metric('涵蓋年度', f"{len(hard_pivot.index)} 個財年")

# ---------- 整理後資料表 ----------
with st.expander('📋 查看整理後資料（長表）', expanded=False):
    st.dataframe(df, use_container_width=True)

# ---------- 圖表 ----------
st.subheader('📈 圖像化總結')
figs, names = build_charts(hard_pivot, soft_pivot, models, hard_life)

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ['硬體各區趨勢', '軟體各區趨勢', '三大機型比較', '硬體 vs 軟體', '累計占比'])
for tab, fig in zip([tab1, tab2, tab3, tab4, tab5], figs):
    with tab:
        st.pyplot(fig, use_container_width=True)

# ---------- 下載 ----------
st.subheader('⬇️ 下載整理後成果')
excel_bio = build_excel(df, hard_pivot, soft_pivot, models,
                        hard_life, soft_life, figs, names)
out_name = '整理後_' + (uploaded.name.split('.')[0] if uploaded.name else '銷量表') + '.xlsx'

col_a, col_b = st.columns(2)
with col_a:
    st.download_button('📥 下載整理後的 Excel（含圖表）', excel_bio,
                       file_name=out_name, mime=(
                           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
with col_b:
    # 讓使用者可挑一張圖單獨下載
    sel = st.selectbox('或單獨下載某一張圖', names)
    idx = names.index(sel)
    buf = io.BytesIO()
    figs[idx].savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    st.download_button('📥 下載這張圖 (PNG)', buf, file_name=sel, mime='image/png')

st.success('完成！Excel 與圖表已可直接下載。')

