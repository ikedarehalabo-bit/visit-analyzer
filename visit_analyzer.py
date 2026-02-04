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
from datetime import datetime, timedelta

# =============================================================================
# 1. 定数・設定パラメータ
# =============================================================================
STAFF_MASTER_FILE = "staff_master.csv"
OFFICE_MASTER_FILE = "office_master.json"

JOB_RANK = {
    "看護師": 1, "准看護師": 1, "保健師": 1,
    "PT": 2, "理学療法士": 2, "OT": 3, "作業療法士": 3, "ST": 4, "言語聴覚士": 4,
    "マネージャー": 80, "事務員": 90, "その他": 99
}

STD_SALARY = {"NURSE": 360000, "REHAB": 270000}

# AI分析用：FTE1.0あたりの理論最大売上 (月間85万円を基準)
THEORETICAL_MAX_REV_PER_FTE = 850000 

KAIGO_UNITS = {20: 313, 30: 470, 40: 470, 60: 821, 90: 1125, "other": 821}
AREA_GRADES = {
    "1級地 (11.40円)": 11.40, "2級地 (11.26円)": 11.26, "3級地 (11.12円)": 11.12,
    "4級地 (10.90円)": 10.90, "5級地 (10.70円)": 10.70, "6級地 (10.42円)": 10.42,
    "7級地 (10.14円)": 10.14, "その他 (10.00円)": 10.00
}

IRYO_BASE = {30: 4250, 60: 5550, 90: 11250, "other": 5550}
IRYO_MANAGE_FEES = {"機能強化型1": 12830, "機能強化型2": 9800, "機能強化型3": 8400, "その他": 7440}

ADDON_PRICES = {
    "iryo_emerg_visit": 2650, "nanbyo_2nd": 4500, "nanbyo_3rd": 8000,
    "iryo_24h_base": 5400, "terminal_base": 25000, "kaigo_emerg_unit": 574
}
PRIVATE_PRICES = {"NURSE_60": 10000, "REHAB_40": 6500}

# =============================================================================
# 2. ユーティリティ関数
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
    norm = normalize_text(job_name)
    for k, v in JOB_RANK.items():
        if k in norm: return v
    return 99

def is_rehab_staff(job_name): return get_job_rank_num(job_name) in [2, 3, 4]
def is_nurse_staff(job_name): return get_job_rank_num(job_name) == 1

def get_default_salary(job_title, fte=1.0):
    rank = get_job_rank_num(job_title)
    if rank == 1: return int(STD_SALARY["NURSE"] * fte)
    elif rank in [2, 3, 4]: return int(STD_SALARY["REHAB"] * fte)
    return 0

def check_flag(text, keywords):
    norm = normalize_text(text)
    return any(k in norm for k in keywords)

def to_excel(df, sheet_name='Sheet1'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True, sheet_name=sheet_name)
    return output.getvalue()

def get_week_label(dt):
    if pd.isna(dt): return ""
    first_day = dt.replace(day=1)
    adjusted_dom = dt.day + first_day.weekday()
    week_num = int(math.ceil(adjusted_dom / 7.0))
    return f"{dt.month}月第{week_num}週"

# =============================================================================
# 3. ファイル解析ロジック (コア機能：自費統合)
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
        except: continue

        for sheet_name, df_raw in sheets:
            df_str = df_raw.fillna("").astype(str)
            if df_str.shape[0] < 6 or df_str.shape[1] < 10: continue
            
            staff_info = df_str.iloc[1, 0].strip() # A2: スタッフ名と職種
            if not staff_info: continue
            
            job_title = "不明"
            m = re.search(r'[（\(](.*?)[）\)]', staff_info)
            if m: job_title = m.group(1).strip()

            START_ROW = 5 
            df_data = df_raw.iloc[START_ROW:].copy()
            
            for _, row in df_data.iterrows():
                d_val = row.iloc[1] # B列: 日付
                if pd.isna(d_val) or str(d_val).strip() == "": continue
                try: 
                    v_date = pd.to_datetime(d_val, errors='coerce')
                    if pd.isna(v_date): continue
                except: continue

                user = str(row.iloc[2]).strip()
                time_str = str(row.iloc[7])
                svc = str(row.iloc[8])
                ins_txt = str(row.iloc[9])

                mins = extract_minutes(time_str)
                f_pvt = "自費" in svc
                f_em = check_flag(svc, ["緊急", "緊"])
                f_nb = "難病複数回" in svc

                # --- カテゴリ決定ロジック ---
                if f_pvt:
                    ins_type = "自費"
                    category_name = "自費" 
                elif "医療" in ins_txt:
                    ins_type = "医療"
                    category_name = f"{mins}分(医療)"
                elif "介護" in ins_txt:
                    ins_type = "介護"
                    category_name = f"{mins}分(介護)"
                else:
                    ins_type = "その他"
                    category_name = f"{mins}分(その他)"

                all_records.append({
                    '氏名': staff_info, '職種': job_title, '訪問日': v_date,
                    '利用者名': user, '時間(分)': mins, '保険': ins_type,
                    'カテゴリ': category_name,
                    'サービス内容': svc, '緊急フラグ': f_em, '難病フラグ': f_nb, '自費フラグ': f_pvt
                })
    return pd.DataFrame(all_records)

# =============================================================================
# 4. マスタデータ管理
# =============================================================================
def load_masters():
    # スタッフマスタ
    cols = ['氏名','職種','役職','人員換算','基準給与','調整額']
    if os.path.exists(STAFF_MASTER_FILE):
        try:
            df_s = pd.read_csv(STAFF_MASTER_FILE)
            if '固定給与' in df_s.columns: df_s.rename(columns={'固定給与':'基準給与'}, inplace=True)
            if '調整額' not in df_s.columns: df_s['調整額'] = 0
        except:
            df_s = pd.DataFrame(columns=cols)
    else:
        df_s = pd.DataFrame(columns=cols)

    # 事業所マスタ
    default_off = {
        "area_grade": "3級地 (11.12円)", "kaigo_em_cnt": 0,
        "fac_type": "機能強化型1", "is_24h": "あり",
        "pl_params": {
            "iryo_24h_contract": 0, "terminal_cases": 0, "ot_pay_total": 0, "sga_total": 0,
            "nurse_em_price": 5000
        },
        "manual_addons": [],
        "incentive_rules": [
            {"target": "リハビリ職", "threshold": 70.0, "price": 4350},
            {"target": "看護職", "threshold": 80.0, "price": 4000}
        ]
    }
    if os.path.exists(OFFICE_MASTER_FILE):
        try:
            with open(OFFICE_MASTER_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                for k,v in default_off.items():
                    if k not in saved: saved[k] = v
                if "pl_manual" in saved: saved["pl_params"].update(saved["pl_manual"])
                return df_s, saved
        except: pass
    
    return df_s, default_off

def save_masters(df_s, dict_o):
    df_s.to_csv(STAFF_MASTER_FILE, index=False)
    with open(OFFICE_MASTER_FILE, 'w', encoding='utf-8') as f:
        json.dump(dict_o, f, ensure_ascii=False, indent=4)

# =============================================================================
# 5. P/L計算エンジン (収入・支出・インセンティブ)
# =============================================================================
def calculate_staff_revenue(df, conf):
    """スタッフごとの売上貢献額算出"""
    area_p = AREA_GRADES.get(conf['area_grade'], 11.12)
    staff_rev = {}
    
    if not df.empty:
        nb_df = df[(df['保険']=='医療') & (df['難病フラグ'])].copy()
        nb_df = nb_df.sort_values(['訪問日','時間(分)'])
        nb_df['seq'] = nb_df.groupby(['訪問日','利用者名']).cumcount() + 1
        df['難病回数'] = 0
        df.loc[nb_df.index, '難病回数'] = nb_df['seq']

    for _, r in df.iterrows():
        name = r['氏名']
        m = r['時間(分)']
        rev = 0
        
        if r['保険'] == '自費':
            if is_nurse_staff(r['職種']): rev = PRIVATE_PRICES["NURSE_60"]
            elif is_rehab_staff(r['職種']): rev = PRIVATE_PRICES["REHAB_40"]
        elif r['保険'] == '介護':
            rev = KAIGO_UNITS.get(m, 821) * area_p
        elif r['保険'] == '医療':
            rank = r['難病回数']
            if rank <= 1: rev = IRYO_BASE.get(m, 5550)
            elif rank == 2: rev = ADDON_PRICES['nanbyo_2nd']
            elif rank >= 3: rev = ADDON_PRICES['nanbyo_3rd']
            
        staff_rev[name] = staff_rev.get(name, 0) + rev
        
    return staff_rev

def run_pl_engine(df, smst, conf):
    area_p = AREA_GRADES.get(conf['area_grade'], 11.12)
    manage_p = IRYO_MANAGE_FEES.get(conf['fac_type'], 7440)
    params = conf.get('pl_params', {})
    
    # 収入
    staff_rev_map = calculate_staff_revenue(df, conf)
    base_rev = sum(staff_rev_map.values())
    
    r_em_iryo = df[(df['保険']=='医療') & (df['緊急フラグ'])].shape[0] * ADDON_PRICES['iryo_emerg_visit']
    users_man = df[(df['保険']=='医療') & (df['利用者名']!='不明')]['利用者名'].nunique()
    r_man = users_man * manage_p
    
    p24 = ADDON_PRICES['iryo_24h_base'] if conf['is_24h'] == "あり" else 0
    r_24 = params.get('iryo_24h_contract', 0) * p24
    r_term = params.get('terminal_cases', 0) * ADDON_PRICES['terminal_base']
    r_add = sum([int(x['price']*x['count']) for x in conf.get('manual_addons', []) if x.get('name')])
    r_k_em = conf['kaigo_em_cnt'] * ADDON_PRICES['kaigo_emerg_unit'] * area_p
    
    total_rev = int(base_rev + r_em_iryo + r_man + r_24 + r_term + r_k_em + r_add)

    # 支出 (変動インセンティブ対応)
    df['cost_min'] = df['時間(分)']
    rehab_40 = (df['保険']=='医療') & (df['職種'].apply(is_rehab_staff)) & (df['時間(分)']==40)
    df.loc[rehab_40, 'cost_min'] = 60 # 特例

    agg = df.groupby(['氏名','職種']).agg(時間=('cost_min','sum'), 緊急=('緊急フラグ','sum')).reset_index()
    merged = pd.merge(smst, agg, on=['氏名','職種'], how='left').fillna(0)
    
    total_exp, details = 0, []
    np = params.get('nurse_em_price', 5000)
    rules = conf.get('incentive_rules', [])

    for _, r in merged.iterrows():
        fix = int(r['基準給与'])
        adj = int(r.get('調整額', 0))
        job = r['職種']
        role = r['役職']
        
        inc = 0
        work_hours = ceil_decimal(r['時間']/60, 1)
        
        if role not in ["管理者", "リーダー"]:
            for rule in rules:
                target_job = rule.get('target', '')
                is_target = False
                if target_job == "全職種": is_target = True
                elif target_job == "リハビリ職" and is_rehab_staff(job): is_target = True
                elif target_job == "看護職" and is_nurse_staff(job): is_target = True
                
                if is_target:
                    th = float(rule.get('threshold', 0))
                    price = float(rule.get('price', 0))
                    if work_hours > th:
                        inc += int(ceil_decimal(work_hours - th, 1) * price)

        em = int(r['緊急'] * np) if is_nurse_staff(job) else 0
        
        gross = fix + inc + em + adj
        cost = int(gross * 1.15) 
        total_exp += cost
        
        details.append({
            "氏名": r['氏名'], "職種": job, "基準給与": fix, 
            "調整額": adj, "インセン": inc, "緊急手当": em, "総コスト": cost
        })
    
    total_exp += params.get('ot_pay_total', 0)
    total_exp += params.get('sga_total', 0)

    return total_rev, total_exp, details, staff_rev_map

# =============================================================================
# 6. UI 実装 (日本語化・完全版)
# =============================================================================
st.set_page_config(page_title="VISIT ANALYZER V12", layout="wide", page_icon="⚡")
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@700;900&family=Noto+Sans+JP:wght@400;700&display=swap');
    .stApp { background-color: #050505; color: #e0e0e0; font-family: 'Noto Sans JP', sans-serif; }
    h1, h2, h3 { font-family: 'Montserrat', 'Noto Sans JP'; color: #fff; text-shadow: 0 0 10px #008080; }
    div[data-testid="stMetricValue"] { color: #00FFFF !important; font-family: 'Montserrat'; }
    section[data-testid="stSidebar"] { background-color: #0a0a0a; border-right: 1px solid #333; }
    .stDataFrame { border: 1px solid #333; }
</style>
""", unsafe_allow_html=True)

if 'master_df' not in st.session_state: st.session_state.master_df = pd.DataFrame()
if 'staff_master' not in st.session_state:
    s, o = load_masters()
    st.session_state.staff_master = s
    st.session_state.office_master = o

with st.sidebar:
    st.title("メニュー")
    page = st.radio("機能選択:", ["ホーム", "データ読込", "集計レポート", "収支・給与分析", "BIダッシュボード", "マスタ設定"])

# --- ページルーティング ---

if page == "ホーム":
    st.title("VISIT ANALYZER V12")
    st.markdown("### 訪問看護 経営・集計統合システム")
    st.info("左側のメニューから機能を選択してください。")
    st.markdown("""
    - **データ読込**: 実績簿(Excel)をドラッグ＆ドロップで取り込みます
    - **集計レポート**: 日次・週次・月次のクロス集計（回数・時間）を行います
    - **収支・給与分析**: インセンティブ計算、人件費率、営業利益を算出します
    - **BIダッシュボード**: 看護・リハ部門別の稼働率やAI訪問効率スコアを表示します
    - **マスタ設定**: 単価設定やスタッフ情報の管理を行います
    """)

elif page == "データ読込":
    st.subheader("実績簿データ読込")
    up = st.file_uploader("実績簿(Excel)をここにドロップ", type=['xlsx'], accept_multiple_files=True)
    if up:
        with st.spinner("データ解析中..."):
            df = parse_files(up)
            if not df.empty:
                st.session_state.master_df = df
                curr = st.session_state.staff_master
                exist = curr['氏名'].tolist()
                new_r = []
                for _, r in df[['氏名','職種']].drop_duplicates().iterrows():
                    n = r['氏名']
                    if n not in exist and n != "不明":
                        j = r['職種']
                        fte = 0.0 if "事務" in j else 1.0
                        bs = get_default_salary(j, fte)
                        new_r.append({'氏名':n, '職種':j, '役職':'一般', '人員換算':fte, '基準給与':bs, '調整額':0})
                if new_r:
                    st.session_state.staff_master = pd.concat([curr, pd.DataFrame(new_r)], ignore_index=True)
                    save_masters(st.session_state.staff_master, st.session_state.office_master)
                st.success(f"読込完了: {len(df)} 件のレコード")

elif page == "集計レポート":
    st.subheader("訪問集計レポート (クロス集計版)")
    df = st.session_state.master_df

    if not df.empty:
        # フィルタリング
        stf = sorted(df['氏名'].unique())
        sel = st.multiselect("スタッフ選択", stf, default=stf)
        
        if sel:
            v = df[df['氏名'].isin(sel)].copy()

            # --- 集計用ヘルパー関数 ---
            def create_pivot(source_df, period_col, period_name):
                # ピボットテーブル作成 (回数と時間を同時に集計)
                p = source_df.pivot_table(
                    index=['氏名', '職種', period_col],
                    columns='カテゴリ',
                    values='時間(分)',
                    aggfunc=['count', 'sum'],
                    fill_value=0
                )
                
                # カラムの整理: 英語のプレフィックスを日本語に置換してフラット化
                new_cols = []
                for t, c in p.columns:
                    prefix = "回数" if t == 'count' else "時間"
                    new_cols.append(f"{prefix}_{c}")
                p.columns = new_cols
                
                # 列の並び替え
                return p.sort_index(axis=1)

            # --- タブによる表示切り替え ---
            t_day, t_week, t_month = st.tabs(["📅 日次レポート", "📊 週次レポート", "📈 月次レポート"])

            # 1. 日次集計
            with t_day:
                v['日次'] = v['訪問日'].dt.strftime('%m/%d(%a)')
                df_day = create_pivot(v, '日次', 'Date')
                st.markdown("##### 日別集計 (回数・時間)")
                st.dataframe(df_day.style.background_gradient(cmap='Blues'), use_container_width=True)
                st.download_button("Excelダウンロード", to_excel(df_day, "Daily"), "report_daily.xlsx")

            # 2. 週次集計
            with t_week:
                v['週次'] = v['訪問日'].apply(get_week_label)
                df_week = create_pivot(v, '週次', 'Week')
                
                # 合計列の計算
                count_cols = [c for c in df_week.columns if c.startswith('回数')]
                sum_cols = [c for c in df_week.columns if c.startswith('時間')]
                
                df_week['【合計】回数'] = df_week[count_cols].sum(axis=1)
                df_week['【合計】時間'] = df_week[sum_cols].sum(axis=1)

                st.markdown("##### 週別集計")
                st.dataframe(df_week.style.background_gradient(cmap='Greens', subset=['【合計】時間']), use_container_width=True)
                st.download_button("Excelダウンロード", to_excel(df_week, "Weekly"), "report_weekly.xlsx")

            # 3. 月次集計
            with t_month:
                v['月次'] = v['訪問日'].dt.strftime('%Y-%m')
                df_month = create_pivot(v, '月次', 'Month')
                
                count_cols = [c for c in df_month.columns if c.startswith('回数')]
                sum_cols = [c for c in df_month.columns if c.startswith('時間')]
                
                df_month['【合計】回数'] = df_month[count_cols].sum(axis=1)
                df_month['【合計】時間'] = df_month[sum_cols].sum(axis=1)
                
                st.markdown("##### 月間集計")
                st.dataframe(df_month.style.background_gradient(cmap='Oranges', subset=['【合計】時間']), use_container_width=True)
                st.download_button("Excelダウンロード", to_excel(df_month, "Monthly"), "report_monthly.xlsx")

    else:
        st.info("データを読み込んでください")

elif page == "収支・給与分析":
    st.subheader("収支・給与シミュレーション")
    df = st.session_state.master_df
    conf = st.session_state.office_master
    smst = st.session_state.staff_master
    
    if not df.empty:
        df['Month'] = df['訪問日'].dt.strftime('%Y-%m')
        tgt = df['Month'].max()
        df_tgt = df[df['Month'] == tgt].copy()
        
        st.markdown(f"**対象月: {tgt}**")
        
        with st.expander("計算パラメータ設定", expanded=False):
            params = conf.get('pl_params', {})
            c1,c2,c3 = st.columns(3)
            p_sga = c1.number_input("販売管理費 (SGA)", value=params.get('sga_total', 0))
            p_ot = c2.number_input("全社残業代", value=params.get('ot_pay_total', 0))
            
            if st.button("パラメータ更新"):
                params['sga_total'] = p_sga
                params['ot_pay_total'] = p_ot
                conf['pl_params'] = params
                save_masters(smst, conf)
                st.rerun()

        rev, exp, rows, _ = run_pl_engine(df_tgt, smst, conf)
        prof = rev - exp
        
        # 経営指標
        labor_cost = sum([r['総コスト'] for r in rows])
        labor_ratio = (labor_cost / rev * 100) if rev > 0 else 0
        profit_margin = (prof / rev * 100) if rev > 0 else 0
        
        

        st.markdown("#### 経営指標 (KPI)")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("人件費率", f"{labor_ratio:.1f} %")
        k2.metric("営業利益率", f"{profit_margin:.1f} %")
        k3.metric("総売上", f"{rev:,}")
        k4.metric("営業利益", f"{prof:,}")
        
        st.markdown("#### 給与・手当明細 (直接編集可)")
        with st.form("salary_edit"):
            row_df = pd.DataFrame(rows)
            edited_df = st.data_editor(
                row_df,
                column_config={
                    "基準給与": st.column_config.NumberColumn(required=True, step=1000),
                    "調整額": st.column_config.NumberColumn(required=True, step=1000),
                    "インセン": st.column_config.NumberColumn(disabled=True),
                    "総コスト": st.column_config.NumberColumn(disabled=True)
                },
                use_container_width=True,
                num_rows="fixed"
            )
            if st.form_submit_button("保存して再計算"):
                for i, r in edited_df.iterrows():
                    name = r['氏名']
                    idx = smst[smst['氏名'] == name].index
                    if not idx.empty:
                        smst.at[idx[0], '基準給与'] = r['基準給与']
                        smst.at[idx[0], '調整額'] = r['調整額']
                st.session_state.staff_master = smst
                save_masters(smst, conf)
                st.success("更新しました")
                time.sleep(0.5); st.rerun()

elif page == "BIダッシュボード":
    st.subheader("AI訪問効率分析")
    if not st.session_state.master_df.empty:
        df = st.session_state.master_df.copy()
        df['Month'] = df['訪問日'].dt.strftime('%Y-%m')
        tgt = df['Month'].max()
        df = df[df['Month'] == tgt]
        
        _, _, _, staff_rev_map = run_pl_engine(df, st.session_state.staff_master, st.session_state.office_master)
        agg = df.groupby(['氏名','職種']).agg(実働時間=('時間(分)','sum')).reset_index()
        mrg = pd.merge(st.session_state.staff_master, agg, on=['氏名','職種'], how='left').fillna(0)
        
        mrg['Dept'] = mrg['職種'].apply(lambda x: 'REHAB' if is_rehab_staff(x) else ('NURSE' if is_nurse_staff(x) else 'OTHER'))
        
        def render_dept_bi(dept_code):
            d = mrg[mrg['Dept'] == dept_code].copy()
            if d.empty: st.info("該当スタッフなし"); return
            
            rows = []
            for _, r in d.iterrows():
                name = r['氏名']
                fte = r['人員換算']
                act_h = ceil_decimal(r['実働時間']/60, 1)
                
                # AI分析: 理論最大値に対する達成度
                actual_rev = staff_rev_map.get(name, 0)
                max_potential = fte * THEORETICAL_MAX_REV_PER_FTE
                eff_score = (actual_rev / max_potential * 100) if max_potential > 0 else 0
                
                rows.append({
                    "氏名": name, "人員換算(FTE)": fte, "実働(H)": act_h, 
                    "売上貢献(円)": int(actual_rev), 
                    "AI効率スコア(%)": ceil_decimal(eff_score, 1)
                })
            
            bdf = pd.DataFrame(rows)
            c1, c2 = st.columns(2)
            c1.metric("部門平均スコア", f"{bdf['AI効率スコア(%)'].mean():.1f} %")
            c2.metric("部門総売上", f"{bdf['売上貢献(円)'].sum():,}")
            
            st.dataframe(bdf.style.background_gradient(subset=['AI効率スコア(%)'], cmap='Blues', vmin=50, vmax=100), use_container_width=True)
            
            ch = alt.Chart(bdf).mark_bar().encode(
                x='氏名', y='AI効率スコア(%)', color=alt.Color('AI効率スコア(%)', scale=alt.Scale(scheme='tealblues'))
            ).properties(height=250)
            st.altair_chart(ch, use_container_width=True)

        t1, t2 = st.tabs(["看護部 (NURSE)", "リハビリ部 (REHAB)"])
        with t1: render_dept_bi('NURSE')
        with t2: render_dept_bi('REHAB')

elif page == "マスタ設定":
    st.subheader("設定・マスタ管理")
    c = st.session_state.office_master
    p = c.setdefault('pl_params', {})
    
    t1, t2 = st.tabs(["事業所・インセンティブ", "スタッフマスタ"])
    with t1:
        with st.form("conf_form"):
            st.markdown("##### 基本情報")
            c1,c2 = st.columns(2)
            ag = c1.selectbox("地域区分", list(AREA_GRADES.keys()), index=list(AREA_GRADES.keys()).index(c['area_grade']))
            ft = c2.selectbox("機能強化型区分", list(IRYO_MANAGE_FEES.keys()), index=list(IRYO_MANAGE_FEES.keys()).index(c['fac_type']))
            
            st.markdown("##### 固定パラメータ")
            c3,c4,c5 = st.columns(3)
            p_sga = c3.number_input("販売管理費", value=p.get('sga_total', 0))
            p_24 = c4.number_input("医療24H契約数", value=p.get('iryo_24h_contract', 0))
            p_tm = c5.number_input("ターミナル件数", value=p.get('terminal_cases', 0))
            
            st.markdown("##### インセンティブ計算ルール")
            st.caption("※対象職種、閾値(時間)、単価(円)を自由に追加できます")
            rules = st.data_editor(
                c.get('incentive_rules', []),
                num_rows="dynamic",
                column_config={
                    "target": st.column_config.SelectboxColumn("対象職種", options=["リハビリ職", "看護職", "全職種"], required=True),
                    "threshold": st.column_config.NumberColumn("閾値時間 (～時間以上)", required=True),
                    "price": st.column_config.NumberColumn("単価 (円)", required=True)
                },
                use_container_width=True
            )
            
            if st.form_submit_button("設定を保存"):
                c.update({'area_grade':ag, 'fac_type':ft, 'incentive_rules':rules})
                c['pl_params'].update({'iryo_24h_contract':p_24, 'terminal_cases':p_tm, 'sga_total':p_sga})
                save_masters(st.session_state.staff_master, c)
                st.success("保存しました")

    with t2:
        with st.form("st_form"):
            edited = st.data_editor(st.session_state.staff_master, num_rows="dynamic", use_container_width=True)
            if st.form_submit_button("スタッフ情報を保存"):
                st.session_state.staff_master = edited
                save_masters(edited, c)
                st.success("保存しました")
