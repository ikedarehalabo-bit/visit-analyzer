import streamlit as st
import pandas as pd
import unicodedata
import re
import time
import math
import os
import json
import altair as alt
from io import BytesIO

# =============================================================================
# 1. CONSTANTS & LOGIC PARAMETERS
# =============================================================================
STAFF_MASTER_FILE = "staff_master.csv"
OFFICE_MASTER_FILE = "office_master.json"

# --- 職種ランク定義 ---
JOB_RANK = {
    "看護師": 1, "准看護師": 1, "保健師": 1,
    "PT": 2, "理学療法士": 2, "OT": 3, "作業療法士": 3, "ST": 4, "言語聴覚士": 4,
    "マネージャー": 80, "事務員": 90, "その他": 99
}

# --- 標準給与 (FTE 1.0時) ---
STD_SALARY = {
    "NURSE": 360000, 
    "REHAB": 270000
}

# --- 介護報酬 (単位数・地域単価) ---
KAIGO_UNITS = {
    20: 313, 30: 470, 40: 470, 60: 821, 90: 1125, "other": 821
}
AREA_GRADES = {
    "1級地 (11.40円)": 11.40, "2級地 (11.26円)": 11.26, "3級地 (11.12円)": 11.12,
    "4級地 (10.90円)": 10.90, "5級地 (10.70円)": 10.70, "6級地 (10.42円)": 10.42,
    "7級地 (10.14円)": 10.14, "その他 (10.00円)": 10.00
}

# --- 医療報酬 (基本療養費) ---
IRYO_BASE = {
    30: 4250, 60: 5550, 90: 11250, "other": 5550
}

# --- 管理療養費 (月額) ---
IRYO_MANAGE_FEES = {
    "機能強化型1": 12830, "機能強化型2": 9800, "機能強化型3": 8400, "その他": 7440
}

# --- 各種加算単価 ---
ADDON_PRICES = {
    "iryo_emerg_visit": 2650,    # 医療: 緊急訪問看護加算(1回)
    "nanbyo_2nd": 4500,          # 難病複数回(1日2回目)
    "nanbyo_3rd": 8000,          # 難病複数回(1日3回目以降)
    "iryo_24h_base": 5400,       # 24H体制加算
    "terminal_base": 25000,      # ターミナルケア
    "kaigo_emerg_unit": 574      # 介護緊急時(単位)
}

# --- 自費単価 ---
PRIVATE_PRICES = {
    "NURSE_60": 10000,
    "REHAB_40": 6500
}

# =============================================================================
# 2. UTILITY FUNCTIONS
# =============================================================================
def ceil_decimal(value, decimals=1):
    if pd.isna(value): return 0.0
    factor = 10 ** decimals
    return math.ceil(value * factor) / factor

def normalize_text(text):
    if pd.isna(text): return ""
    return unicodedata.normalize('NFKC', str(text))

def extract_minutes(text):
    text_norm = normalize_text(text)
    match = re.search(r'(\d+)', text_norm)
    return int(match.group(1)) if match else 0

def get_job_rank_num(job_name):
    norm_name = normalize_text(job_name)
    for key, rank in JOB_RANK.items():
        if key in norm_name: return rank
    return 99

def is_rehab_staff(job_name):
    return get_job_rank_num(job_name) in [2, 3, 4]

def is_nurse_staff(job_name):
    return get_job_rank_num(job_name) == 1

def check_flag(text, keywords):
    norm_text = normalize_text(text)
    return any(k in norm_text for k in keywords)

def get_default_salary(job_title, fte=1.0):
    rank = get_job_rank_num(job_title)
    if rank == 1: return int(STD_SALARY["NURSE"] * fte)
    elif rank in [2, 3, 4]: return int(STD_SALARY["REHAB"] * fte)
    return 0

def to_excel(df, sheet_name='Sheet1'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True, sheet_name=sheet_name)
    return output.getvalue()

# =============================================================================
# 3. PRECISE FILE PARSER
# =============================================================================
@st.cache_data
def parse_files(uploaded_files):
    all_records = []
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.xlsx'):
                xls = pd.read_excel(file, sheet_name=None, header=None)
                sheets = xls.items()
            else:
                df_c = pd.read_csv(file, header=None, encoding='utf-8-sig')
                sheets = [("CSV", df_c)]
        except Exception:
            continue

        for sheet_name, df_raw in sheets:
            df_str = df_raw.fillna("").astype(str)
            if df_str.shape[0] < 6 or df_str.shape[1] < 10: continue
            
            staff_info_cell = df_str.iloc[1, 0] # A2
            staff_name = str(staff_info_cell).strip()
            if not staff_name: continue

            job_title = "不明"
            match = re.search(r'[（\(](.*?)[）\)]', staff_name)
            if match: job_title = match.group(1).strip()

            START_ROW_IDX = 5 
            COL_DATE = 1
            COL_USER = 2
            COL_TIME = 7
            COL_SERVICE = 8
            COL_INSURANCE = 9

            df_data = df_raw.iloc[START_ROW_IDX:].copy()
            
            for _, row in df_data.iterrows():
                date_val = row.iloc[COL_DATE]
                if pd.isna(date_val) or str(date_val).strip() == "": continue
                try: 
                    visit_date = pd.to_datetime(date_val, errors='coerce')
                    if pd.isna(visit_date): continue
                except: continue

                user_name = str(row.iloc[COL_USER]).strip()
                time_str = str(row.iloc[COL_TIME])
                service_txt = str(row.iloc[COL_SERVICE])
                ins_txt = str(row.iloc[COL_INSURANCE])

                mins = extract_minutes(time_str)
                
                if "医療" in ins_txt: ins_type = "医療"
                elif "介護" in ins_txt: ins_type = "介護"
                else: ins_type = "その他"

                f_em = check_flag(service_txt, ["緊急", "緊"])
                f_psy = check_flag(service_txt, ["精", "精神"])
                f_nb = "難病複数回" in service_txt
                f_pvt = "自費" in service_txt
                f_term = check_flag(service_txt, ["看取", "ターミナル"])

                all_records.append({
                    '氏名': staff_name,
                    '利用者名': user_name,
                    '職種': job_title,
                    '訪問日': visit_date,
                    '時間(分)': mins,
                    '保険': ins_type,
                    'カテゴリ': f"{mins}分({ins_type})",
                    'サービス内容': service_txt,
                    '緊急フラグ': f_em,
                    '精神科フラグ': f_psy,
                    '難病フラグ': f_nb,
                    '自費フラグ': f_pvt,
                    'ターミナルフラグ': f_term,
                    '元ファイル': file.name
                })

    return pd.DataFrame(all_records)
# =============================================================================
# 4. MASTER DATA MANAGEMENT
# =============================================================================
def load_masters():
    # Staff Master
    if os.path.exists(STAFF_MASTER_FILE):
        try:
            df_s = pd.read_csv(STAFF_MASTER_FILE)
            if '固定給与' in df_s.columns: df_s.rename(columns={'固定給与':'基準給与'}, inplace=True)
        except:
            df_s = pd.DataFrame(columns=['氏名','職種','役職','人員換算','基準給与'])
    else:
        df_s = pd.DataFrame(columns=['氏名','職種','役職','人員換算','基準給与'])

    # Office Master
    default_off = {
        "area_grade": "3級地 (11.12円)", 
        "kaigo_em_cnt": 0,
        "fac_type": "機能強化型1", 
        "is_24h": "あり",
        "pl_manual": {
            "iryo_24h_contract": 0,
            "terminal_cases": 0,
            "ot_pay_total": 0
        },
        "manual_addons": []
    }
    if os.path.exists(OFFICE_MASTER_FILE):
        try:
            with open(OFFICE_MASTER_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                for k,v in default_off.items():
                    if k not in saved: saved[k] = v
                return df_s, saved
        except: pass
    
    return df_s, default_off

def save_masters(df_s, dict_o):
    df_s.to_csv(STAFF_MASTER_FILE, index=False)
    with open(OFFICE_MASTER_FILE, 'w', encoding='utf-8') as f:
        json.dump(dict_o, f, ensure_ascii=False, indent=4)

# =============================================================================
# 5. CORE CALCULATION ENGINE (P/L)
# =============================================================================
def run_pl_engine(df, smst, conf):
    area_p = AREA_GRADES.get(conf['area_grade'], 11.12)
    manage_p = IRYO_MANAGE_FEES.get(conf['fac_type'], 7440)
    
    # --- 1. 収入計算 (Revenue) ---
    if not df.empty:
        nb_df = df[(df['保険'] == '医療') & (df['難病フラグ'] == True)].copy()
        nb_df = nb_df.sort_values(['訪問日', '時間(分)'])
        nb_df['seq'] = nb_df.groupby(['訪問日', '利用者名']).cumcount() + 1
        df['難病回数'] = 0
        df.loc[nb_df.index, '難病回数'] = nb_df['seq']
    else:
        df['難病回数'] = 0

    r_kaigo, r_iryo, r_pvt, r_nb = 0, 0, 0, 0
    
    for _, r in df.iterrows():
        m = r['時間(分)']
        job = r['職種']
        
        # A. 自費
        if r['自費フラグ']:
            if is_nurse_staff(job): r_pvt += PRIVATE_PRICES["NURSE_60"]
            elif is_rehab_staff(job): r_pvt += PRIVATE_PRICES["REHAB_40"]
            continue

        # B. 介護
        if r['保険'] == '介護':
            u = KAIGO_UNITS.get(m, 821)
            r_kaigo += (u * area_p)
        
        # C. 医療
        elif r['保険'] == '医療':
            rank = r['難病回数']
            if rank <= 1: r_iryo += IRYO_BASE.get(m, 5550)
            elif rank == 2: r_nb += ADDON_PRICES['nanbyo_2nd']
            elif rank >= 3: r_nb += ADDON_PRICES['nanbyo_3rd']

    r_em_iryo = df[(df['保険']=='医療') & (df['緊急フラグ'])].shape[0] * ADDON_PRICES['iryo_emerg_visit']
    users_manage = df[(df['保険']=='医療') & (df['利用者名']!='不明')]['利用者名'].nunique()
    r_man = users_manage * manage_p
    
    m_in = conf['pl_manual']
    p24 = ADDON_PRICES['iryo_24h_base'] if conf['is_24h'] == "あり" else 0
    r_24 = m_in.get('iryo_24h_contract', 0) * p24
    r_term = m_in.get('terminal_cases', 0) * ADDON_PRICES['terminal_base']
    r_man_add = sum([int(x['price']*x['count']) for x in conf.get('manual_addons', []) if x.get('name')])
    r_kaigo_em = conf['kaigo_em_cnt'] * ADDON_PRICES['kaigo_emerg_unit'] * area_p
    
    total_rev = int(r_kaigo + r_iryo + r_pvt + r_nb + r_em_iryo + r_man + r_24 + r_term + r_kaigo_em + r_man_add)

    # --- 2. 支出計算 (Expenditure) ---
    df['cost_min'] = df['時間(分)']
    rehab_40_mask = (df['保険'] == '医療') & (df['職種'].apply(is_rehab_staff)) & (df['時間(分)'] == 40)
    df.loc[rehab_40_mask, 'cost_min'] = 60

    agg = df.groupby(['氏名', '職種']).agg(時間=('cost_min','sum'), 緊急=('緊急フラグ','sum')).reset_index()
    merged = pd.merge(smst, agg, on=['氏名','職種'], how='left').fillna(0)
    
    total_exp, details = 0, []
    
    for _, r in merged.iterrows():
        fix = int(r['基準給与'])
        job = r['職種']
        role = r['役職']
        inc = 0
        if is_rehab_staff(job) and role not in ["管理者", "リーダー"]:
            th = ceil_decimal(r['時間']/60, 1)
            if th > 70: inc = int(ceil_decimal(th-70, 1) * 4350)
        
        em = int(r['緊急'] * 5000) if is_nurse_staff(job) else 0
        gross = fix + inc + em
        cost = int(gross * 1.15)
        
        total_exp += cost
        details.append({"氏名": r['氏名'], "固定": fix, "インセン": inc, "緊急手当": em, "コスト": cost})
    
    ot_pay = m_in.get('ot_pay_total', 0)
    total_exp += ot_pay

    return total_rev, total_exp, details, {
        "管理人数": users_manage, 
        "医療緊急回数": int(r_em_iryo / ADDON_PRICES['iryo_emerg_visit'])
    }
# =============================================================================
# 6. UI IMPLEMENTATION (SIDEBAR NAV)
# =============================================================================
st.set_page_config(page_title="VISIT ANALYZER V9", layout="wide", page_icon="⚡")
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

# Cyberpunk Style
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;900&family=Noto+Sans+JP:wght@400;700&display=swap');
    .stApp { background-color: #050505; color: #e0e0e0; font-family: 'Noto Sans JP', sans-serif; }
    h1, h2, h3 { font-family: 'Montserrat', 'Noto Sans JP'; color: #fff; text-shadow: 0 0 10px #00FFFF; }
    div[data-testid="stMetricValue"] { color: #00FFFF !important; font-family: 'Montserrat'; }
    .stButton>button { background: #000; color: #00FFFF; border: 1px solid #00FFFF; font-weight: bold; }
    .stButton>button:hover { background: #00FFFF; color: #000; }
    /* Sidebar */
    section[data-testid="stSidebar"] { background-color: #0a0a0a; border-right: 1px solid #333; }
</style>
""", unsafe_allow_html=True)

# Init State
if 'master_df' not in st.session_state: st.session_state.master_df = pd.DataFrame()
if 'staff_master' not in st.session_state:
    s, o = load_masters()
    st.session_state.staff_master = s
    st.session_state.office_master = o

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.title("MENU")
    page = st.radio("Go to:", 
        ["HOME", "UPLOAD", "REPORTS", "P/L ANALYZER", "BI DASHBOARD", "SETTINGS"]
    )
    st.divider()
    st.caption("VISIT ANALYZER V9")

# --- MAIN PAGE ROUTING ---

# 1. HOME
if page == "HOME":
    st.title("VISIT ANALYZER V9")
    st.markdown("### 訪問看護経営・集計完全統合モデル")
    st.info("左側のサイドバーメニューから機能を選択してください。")
    st.image("https://streamlit.io/images/brand/streamlit-mark-color.png", width=100)
    st.markdown("""
    - **UPLOAD**: 実績簿（Excel）の読み込み
    - **REPORTS**: スタッフ別・日次/月次レポート（Excel出力可）
    - **P/L ANALYZER**: 収支シミュレーション・給与分析
    - **BI DASHBOARD**: 稼働率・生産性分析
    - **SETTINGS**: マスタ管理（単価・人員など）
    """)

# 2. UPLOAD
elif page == "UPLOAD":
    st.subheader("📂 実績簿アップロード")
    up = st.file_uploader("実績簿(Excel)をドロップ", type=['xlsx'], accept_multiple_files=True)
    if up:
        with st.spinner("Processing..."):
            df = parse_files(up)
            if not df.empty:
                st.session_state.master_df = df
                # Auto Register Logic
                curr = st.session_state.staff_master
                exist = curr['氏名'].tolist()
                new_r = []
                for _, r in df[['氏名','職種']].drop_duplicates().iterrows():
                    n = r['氏名']
                    if n not in exist and n != "不明":
                        j = r['職種']
                        fte = 0.0 if "事務" in j else 1.0
                        bs = get_default_salary(j, fte)
                        new_r.append({'氏名':n, '職種':j, '役職':'一般', '人員換算':fte, '基準給与':bs})
                if new_r:
                    st.session_state.staff_master = pd.concat([curr, pd.DataFrame(new_r)], ignore_index=True)
                    save_masters(st.session_state.staff_master, st.session_state.office_master)
                st.success(f"読込完了: {len(df)}件")
            else: st.error("データなし")

# 3. REPORTS
elif page == "REPORTS":
    st.subheader("📊 集計レポート")
    df = st.session_state.master_df
    if not df.empty:
        stf = sorted(df['氏名'].unique())
        sel = st.multiselect("スタッフ絞り込み", stf, default=stf)
        if sel:
            v = df[df['氏名'].isin(sel)].copy()
            t1, t2 = st.tabs(["週次レポート", "月次レポート"])
            with t1:
                v['Week'] = v['訪問日'] - pd.to_timedelta(v['訪問日'].dt.weekday, unit='D')
                p = v.pivot_table(index=['氏名','Week'], columns='カテゴリ', aggfunc='size', fill_value=0)
                p['Total'] = p.sum(axis=1)
                st.dataframe(p.style.background_gradient(cmap='Greens'), use_container_width=True)
                st.download_button("📥 Excel保存 (週次)", to_excel(p, "Weekly"), "weekly_report.xlsx")
            with t2:
                v['Month'] = v['訪問日'].dt.strftime('%Y-%m')
                p = v.pivot_table(index=['氏名','Month'], columns='カテゴリ', aggfunc='size', fill_value=0)
                p['Total'] = p.sum(axis=1)
                st.dataframe(p.style.background_gradient(cmap='Greens'), use_container_width=True)
                st.download_button("📥 Excel保存 (月次)", to_excel(p, "Monthly"), "monthly_report.xlsx")
    else: st.warning("データ未読み込み")

# 4. P/L ANALYZER
elif page == "P/L ANALYZER":
    st.subheader("💰 収支・給与分析")
    df = st.session_state.master_df
    conf = st.session_state.office_master
    if not df.empty:
        df['Month'] = df['訪問日'].dt.strftime('%Y-%m')
        target = df['Month'].max()
        df_tgt = df[df['Month'] == target].copy()
        
        st.markdown(f"**対象月: {target}**")
        
        with st.expander("📝 計算パラメータ調整", expanded=True):
            c1,c2,c3 = st.columns(3)
            saved_c = conf.get('pl_manual', {})
            in_24h = c1.number_input("医療:24H契約数", value=saved_c.get('iryo_24h_contract', 0))
            in_term = c2.number_input("医療:ターミナル件数", value=saved_c.get('terminal_cases', 0))
            in_ot = c3.number_input("全社残業代(円)", value=saved_c.get('ot_pay_total', 0))
            
            addons = st.data_editor(conf.get('manual_addons', []), num_rows="dynamic",
                                  column_config={"name":"項目名","price":"単価","count":"件数"}, use_container_width=True)
            
            if st.button("計算実行 & 保存"):
                conf['pl_manual'] = {'iryo_24h_contract': in_24h, 'terminal_cases': in_term, 'ot_pay_total': in_ot}
                conf['manual_addons'] = addons
                save_masters(st.session_state.staff_master, conf)
                st.rerun()

        rev, exp, rows, details = run_pl_engine(df_tgt, st.session_state.staff_master, conf)
        prof = rev - exp
        
        st.divider()
        k1,k2,k3 = st.columns(3)
        k1.metric("総収益 (Revenue)", f"{rev:,} 円")
        k2.metric("総支出 (Cost)", f"{exp:,} 円")
        k3.metric("営業利益 (Profit)", f"{prof:,} 円", delta_color="normal")
        
        st.markdown("##### 人件費・手当内訳")
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
    else: st.warning("データ未読み込み")

# 5. BI DASHBOARD
elif page == "BI DASHBOARD":
    st.subheader("🚀 経営分析ダッシュボード")
    if not st.session_state.master_df.empty:
        std = st.number_input("月間所定労働時間", 160)
        df = st.session_state.master_df.copy()
        df['Month'] = df['訪問日'].dt.strftime('%Y-%m')
        target = df['Month'].max()
        df = df[df['Month'] == target]
        
        agg = df.groupby(['氏名','職種']).agg(時間=('時間(分)','sum')).reset_index()
        mrg = pd.merge(st.session_state.staff_master, agg, on=['氏名','職種'], how='left').fillna(0)
        
        bi = []
        for _, r in mrg.iterrows():
            if "事務" in r['職種']: continue
            act = ceil_decimal(r['時間']/60, 1)
            req = ceil_decimal(r['人員換算']*std, 1)
            rate = ceil_decimal((act/req)*100, 1) if req>0 else 0
            bi.append({"氏名":r['氏名'], "FTE":r['人員換算'], "実働(H)":act, "稼働率(%)":rate})
        
        c_df = pd.DataFrame(bi)
        st.dataframe(c_df.style.background_gradient(subset=['稼働率(%)'], cmap='Oranges'), use_container_width=True)
        
        chart = alt.Chart(c_df).mark_bar().encode(
            x='氏名', y='稼働率(%)', color='氏名'
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)
    else: st.warning("データ未読み込み")

# 6. SETTINGS
elif page == "SETTINGS":
    st.subheader("🛠️ マスタ設定")
    t1, t2 = st.tabs(["事業所設定", "スタッフ設定"])
    with t1:
        c = st.session_state.office_master
        with st.form("ofc"):
            c1,c2 = st.columns(2)
            ag = c1.selectbox("地域区分", list(AREA_GRADES.keys()), index=list(AREA_GRADES.keys()).index(c['area_grade']))
            ke = c2.number_input("介護緊急時契約数", value=c['kaigo_em_cnt'])
            c3,c4 = st.columns(2)
            ft = c3.selectbox("機能強化型区分", list(IRYO_MANAGE_FEES.keys()), index=list(IRYO_MANAGE_FEES.keys()).index(c['fac_type']))
            ih = c4.radio("24H体制", ["あり","なし"], index=["あり","なし"].index(c['is_24h']))
            if st.form_submit_button("保存"):
                c.update({'area_grade':ag, 'kaigo_em_cnt':ke, 'fac_type':ft, 'is_24h':ih})
                save_masters(st.session_state.staff_master, c)
                st.success("設定を保存しました")
    
    with t2:
        with st.form("edt"):
            ed = st.data_editor(st.session_state.staff_master, num_rows="dynamic", use_container_width=True)
            if st.form_submit_button("保存"):
                for i, r in ed.iterrows():
                    if "事務" in r['職種']: ed.at[i,'人員換算'] = 0.0
                    if r['役職'] in ["管理者","リーダー"] or r['職種'] in ["事務員","マネージャー"]: pass
                    else: ed.at[i,'基準給与'] = get_default_salary(r['職種'], r['人員換算'])
                st.session_state.staff_master = ed
                save_masters(ed, st.session_state.office_master)
                st.success("スタッフ情報を保存しました")
                time.sleep(1); st.rerun()
