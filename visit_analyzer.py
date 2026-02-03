import streamlit as st
import pandas as pd
import unicodedata
import re
import time
from io import BytesIO

# ---------------------------------------------------------
# 1. 共通関数 & 設定
# ---------------------------------------------------------
def normalize_and_extract_minutes(text):
    if pd.isna(text): return None, None
    text_norm = unicodedata.normalize('NFKC', str(text))
    match = re.search(r'(\d+)', text_norm)
    if match: return int(match.group(1)), text_norm
    return None, text_norm

def extract_job_title(name_str):
    if pd.isna(name_str): return ""
    match = re.search(r'[（\(](.*?)[）\)]', str(name_str))
    if match: return match.group(1).strip()
    return "不明"

def is_emergency(service_content):
    if pd.isna(service_content): return False
    return "緊" in str(service_content) or "緊急" in str(service_content)

def navigate_to(page):
    st.session_state.current_page = page
    st.rerun()

def load_file_content(file):
    results = []
    if file.name.endswith('.xlsx'):
        try:
            xls = pd.read_excel(file, sheet_name=None, header=None)
            for sname, df in xls.items(): results.append((df, f"{file.name}[{sname}]"))
        except Exception as e: return [], str(e)
    else:
        try:
            df = pd.read_csv(file, header=None)
            results.append((df, file.name))
        except Exception as e: return [], str(e)
    return results, None

def parse_single_dataframe(df_raw, source_name):
    try:
        lines_df = df_raw.fillna("").astype(str)
        # 最低限の行数チェック
        if len(lines_df) < 2: return [], "データ行不足"
        
        # A2セル(インデックス1,0)付近にある氏名を取得トライ
        full_name = lines_df.iloc[1, 0].strip()
        if not full_name: return [], "氏名欄(A2)が空欄"

        # ヘッダー行を探索
        header_row_idx = -1
        for idx, row in lines_df.iterrows():
            row_str = " ".join(row.values)
            if "訪問日" in row_str and "S提供時間" in row_str:
                header_row_idx = idx
                break
        if header_row_idx == -1: return [], "ヘッダー行なし"

        # データ抽出
        df_data = df_raw.iloc[header_row_idx + 1:].copy()
        df_data.columns = [str(c).strip() for c in df_raw.iloc[header_row_idx].values]
        
        required = ['訪問日', 'S提供時間', 'サービス内容', '保険適用']
        missing = [c for c in required if c not in df_data.columns]
        if missing: return [], f"必須列不足: {','.join(missing)}"

        job_title = extract_job_title(full_name)
        records = []
        target_mins = [20, 30, 40, 60, 90]

        for _, row in df_data.iterrows():
            try:
                v_date = pd.to_datetime(row['訪問日'])
                if pd.isna(v_date): continue
            except: continue

            minute, _ = normalize_and_extract_minutes(row['S提供時間'])
            mins = minute if minute else 0
            ins = "医療" if "医療" in str(row['保険適用']) else ("介護" if "介護" in str(row['保険適用']) else "その他")
            cat_min = f"{mins}分" if mins in target_mins else "その他時間"
            
            records.append({
                '氏名': full_name, '職種': job_title, '訪問日': v_date,
                '時間(分)': mins, '保険': ins, 'カテゴリ': f"{cat_min}（{ins}）",
                'サービス内容': row['サービス内容'], '緊急フラグ': is_emergency(row['サービス内容']),
                '元ファイル': source_name
            })
        return records, None
    except Exception as e: return [], str(e)

# ページ設定
st.set_page_config(page_title="VISIT ANALYZER Lite", layout="wide", page_icon="⚡")
st.set_page_config(page_title="VISIT ANALYZER Lite", layout="wide", page_icon="⚡")
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

if 'first_load' not in st.session_state: st.session_state.first_load = True
if 'current_page' not in st.session_state: st.session_state.current_page = "HOME"
if 'master_df' not in st.session_state: st.session_state.master_df = pd.DataFrame()

# ---------------------------------------------------------
# 2. CSS & デザイン定義
# ---------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;900&family=Noto+Sans+JP:wght@400;700&display=swap');
    .stApp { background-color: #050505; color: #e0e0e0; font-family: 'Noto Sans JP', sans-serif; }
    h1, h2, h3 { font-family: 'Montserrat', 'Noto Sans JP'; color: #ffffff; text-transform: uppercase; }
    h1 { text-shadow: 0 0 15px #00FF41; }
    .stButton>button { background: #000; color: #00FF41; border: 1px solid #00FF41; font-weight: bold; width: 100%; }
    .stButton>button:hover { background: #00FF41; color: #000; }
    
    #intro-overlay {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: #000; z-index: 9999;
        display: flex; justify-content: center; align-items: center; animation: fadeOutOverlay 2.5s forwards; pointer-events: none;
    }
    #intro-logo { font-family: 'Montserrat', sans-serif; font-size: 3rem; color: #00FF41; opacity: 0; animation: popInLogo 2s forwards; }
    @keyframes popInLogo { 0% { opacity: 0; transform: scale(0.8); } 50% { opacity: 1; transform: scale(1.1); } 100% { opacity: 0; transform: scale(1.5); } }
    @keyframes fadeOutOverlay { 0% { opacity: 1; } 80% { opacity: 1; } 100% { opacity: 0; visibility: hidden; } }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 3. 起動アニメーション
# ---------------------------------------------------------
if st.session_state.first_load:
    st.markdown('<div id="intro-overlay"><div id="intro-logo">VISIT ANALYZER Lite</div></div>', unsafe_allow_html=True)
    time.sleep(2.0)
    st.session_state.first_load = False

# ---------------------------------------------------------
# 4. メイン処理
# ---------------------------------------------------------
st.title("VISIT ANALYZER Lite")

# --- HOME ---
if st.session_state.current_page == "HOME":
    st.markdown("#### シンプル訪問集計ツール")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📂 データ読み込み (UPLOAD)", use_container_width=True): navigate_to("UPLOAD")
    with c2:
        if st.button("📊 集計レポート (REPORTS)", use_container_width=True): navigate_to("REPORTS")

else:
    # --- HEADER ---
    col_head1, col_head2 = st.columns([9, 1])
    with col_head1:
        titles = {"UPLOAD": "データ読み込み", "REPORTS": "集計レポート"}
        st.markdown(f"### :: {titles.get(st.session_state.current_page, '')}")
    with col_head2:
        if st.button("✕", key="close_main"): navigate_to("HOME")

    # --- UPLOAD ---
    if st.session_state.current_page == "UPLOAD":
        st.info("実績簿 (CSV/Excel) をドラッグ＆ドロップしてください。")
        uploaded_files = st.file_uploader("", type=['csv', 'xlsx'], accept_multiple_files=True)
        
        if uploaded_files:
            all_recs = []
            bar = st.progress(0)
            for i, f in enumerate(uploaded_files):
                d_list, err = load_file_content(f)
                if not err:
                    for df_raw, src in d_list:
                        recs, perr = parse_single_dataframe(df_raw, src)
                        if recs: all_recs.extend(recs)
                bar.progress((i+1)/len(uploaded_files))
            
            if all_recs:
                st.session_state.master_df = pd.DataFrame(all_recs)
                st.success(f"{len(all_recs)} 件のデータを読み込みました。")
            else:
                st.warning("有効なデータが見つかりませんでした。")

    # --- REPORTS ---
    elif st.session_state.current_page == "REPORTS":
        if not st.session_state.master_df.empty:
            df = st.session_state.master_df.copy()
            names = sorted(df['氏名'].unique())
            sel = st.multiselect("スタッフ絞り込み:", names, default=names)
            
            if sel:
                df = df[df['氏名'].isin(sel)]
                t1, t2 = st.tabs(["日次・週次", "月間サマリー"])
                
                with t1:
                    m = st.radio("表示モード", ["日次", "週次"], horizontal=True)
                    if m == "日次":
                        p = df.pivot_table(index=['氏名', '職種', '訪問日'], columns='カテゴリ', aggfunc='size', fill_value=0)
                        p['合計'] = p.sum(axis=1)
                        st.dataframe(p.style.background_gradient(cmap='Greens', subset=['合計']), use_container_width=True)
                    else:
                        df['週'] = df['訪問日'] - pd.to_timedelta(df['訪問日'].dt.weekday, unit='D')
                        p = df.pivot_table(index=['氏名', '職種', '週'], columns='カテゴリ', aggfunc='size', fill_value=0)
                        p['合計'] = p.sum(axis=1)
                        st.dataframe(p.style.format({"週": "{:%Y-%m-%d}"}).background_gradient(cmap='Greens', subset=['合計']), use_container_width=True)
                
                with t2:
                    df['月'] = df['訪問日'].dt.strftime('%Y-%m')
                    p = df.pivot_table(index=['氏名', '職種', '月'], columns='カテゴリ', aggfunc='size', fill_value=0)
                    p['合計'] = p.sum(axis=1)
                    st.dataframe(p.style.background_gradient(cmap='Greens', subset=['合計']), use_container_width=True)
            else:
                st.warning("スタッフを選択してください。")
        else:

            st.error("データがありません。「データ読み込み」を行ってください。")
