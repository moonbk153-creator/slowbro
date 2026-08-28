import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io 
import pytz
import numpy as np
import re
from sqlalchemy import create_engine, text

# [최적화] 공휴일 로드 및 오류 방지
try:
    import holidays
    HAS_HOLIDAYS = True
    kr_holidays_dict = holidays.KR(years=range(2020, 2035))
    KR_HOLIDAYS = [str(date) for date in kr_holidays_dict.keys()]
except ImportError:
    HAS_HOLIDAYS = False
    KR_HOLIDAYS = []

try: from streamlit_autorefresh import st_autorefresh
except ImportError: st_autorefresh = None

KST = pytz.timezone('Asia/Seoul')
st.set_page_config(page_title="색도 관리 시스템", layout="wide")

if 'show_toast' in st.session_state:
    st.toast(st.session_state['show_toast'], icon="✅")
    del st.session_state['show_toast']

EXCEL_FILE = 'data sheet.xlsx'
EQUIPMENT_LIST = ["버닝", "태환 12kg", "프로밧 25kg", "뷸러 60kg", "뷸러 120kg"]
ADMIN_PASSWORD, ACCESS_PASSWORD = st.secrets["ADMIN_PASSWORD"], st.secrets["APP_PASSWORD"]

# ----------------------------------------------------
# 1. 인증 및 기본 설정
# ----------------------------------------------------
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if not st.session_state['logged_in']:
    if "pw" in st.query_params and st.query_params["pw"] == ACCESS_PASSWORD:
        st.session_state['logged_in'] = True
        st.rerun()
    st.title("🔒 색도 관리 시스템 - 접속 제한")
    input_pw = st.text_input("사내 공용 비밀번호를 입력하세요", type="password")
    if st.button("🔓 접속하기"):
        if input_pw == ACCESS_PASSWORD:
            st.session_state['logged_in'] = True
            st.session_state['show_toast'] = "시스템 접속 성공!"
            st.rerun()
        else: st.error("❌ 비밀번호 불일치")
    st.stop()

def get_now_kst(): return datetime.now(KST)

def safe_date_parse(val):
    v = str(val).strip()
    if v in ['nan', 'None', '', 'NaN', 'NaT']: return ""
    v = v.split(" ")[0].replace("/", "-").replace(".", "-")
    try: return pd.to_datetime(v).strftime("%Y-%m-%d")
    except: return v

# ----------------------------------------------------
# 2. 클라우드 DB 연동 (Supabase PostgreSQL + SQLAlchemy)
# ----------------------------------------------------
@st.cache_resource
def get_engine():
    db_url = st.secrets["DB_URL"].replace("postgres://", "postgresql://")
    return create_engine(db_url, pool_pre_ping=True)

def db_execute(query, params={}):
    with get_engine().begin() as conn:
        conn.execute(text(query), params)

def db_fetchall(query, params={}):
    with get_engine().connect() as conn:
        res = conn.execute(text(query), params)
        return [tuple(r) for r in res]

def db_fetchone(query, params={}):
    with get_engine().connect() as conn:
        res = conn.execute(text(query), params)
        r = res.fetchone()
        return tuple(r) if r else None

def init_db():
    db_execute('''CREATE TABLE IF NOT EXISTS color_records (id SERIAL PRIMARY KEY, timestamp TEXT, production_date TEXT, equipment TEXT, worker TEXT, product_name TEXT, target_value REAL, measured_value REAL, difference REAL, status TEXT, remarks TEXT DEFAULT '', input_amount TEXT DEFAULT '-', checked INTEGER DEFAULT 0)''')
    db_execute('''CREATE TABLE IF NOT EXISTS target_history (id SERIAL PRIMARY KEY, product_name TEXT, target_value REAL, effective_date TEXT)''')
    db_execute('''CREATE TABLE IF NOT EXISTS product_notices (product_name TEXT PRIMARY KEY, notice_text TEXT, start_date TEXT, end_date TEXT)''')
    db_execute('''CREATE TABLE IF NOT EXISTS workers (name TEXT PRIMARY KEY)''')
    
    db_execute("CREATE INDEX IF NOT EXISTS idx_color_prod_date ON color_records(product_name, production_date)")
    db_execute("CREATE INDEX IF NOT EXISTS idx_target_hist ON target_history(product_name, effective_date)")
    
    cnt = db_fetchone("SELECT count(*) FROM workers")[0]
    if cnt == 0:
        for w in ["윤승태", "오세현", "조성윤", "이민형"]: 
            db_execute("INSERT INTO workers (name) VALUES (:w)", {"w": w})

def get_all_workers():
    return [r[0] for r in db_fetchall("SELECT name FROM workers")]

def add_worker(name):
    try: 
        db_execute("INSERT INTO workers (name) VALUES (:name)", {"name": name.strip()})
        return True
    except: return False

def delete_worker(name):
    db_execute("DELETE FROM workers WHERE name = :name", {"name": name})

def update_checked_status(record_ids, status_val):
    for r in record_ids: 
        db_execute("UPDATE color_records SET checked=:status WHERE id=:id", {"status": status_val, "id": r})

def delete_from_db(r_id):
    db_execute("DELETE FROM color_records WHERE id = :id", {"id": r_id})

def update_db(r_id, d_date, eq, wk, p, tgt, meas, diff, st_val, rmks, amt, chk=0):
    db_execute('''UPDATE color_records SET production_date=:pd, equipment=:eq, worker=:wk, product_name=:p, target_value=:tgt, measured_value=:meas, difference=:diff, status=:st_val, remarks=:rmks, input_amount=:amt, checked=:chk WHERE id=:id''', 
              {"pd":d_date, "eq":str(eq).strip(), "wk":str(wk).strip(), "p":str(p).strip(), "tgt":tgt, "meas":meas, "diff":diff, "st_val":st_val, "rmks":rmks, "amt":amt, "chk":chk, "id":r_id})

def get_historical_target(p_name, d_str):
    r = db_fetchone('SELECT target_value FROM target_history WHERE product_name=:p AND effective_date <= :d ORDER BY effective_date DESC, id DESC LIMIT 1', {"p":p_name, "d":d_str})
    return r[0] if r else TARGET_DATA.get(p_name, 0.0)

def save_notice(p, txt, s_date, e_date):
    db_execute('''INSERT INTO product_notices (product_name, notice_text, start_date, end_date) VALUES (:p, :txt, :sd, :ed) ON CONFLICT (product_name) DO UPDATE SET notice_text=EXCLUDED.notice_text, start_date=EXCLUDED.start_date, end_date=EXCLUDED.end_date''', {"p":p, "txt":txt, "sd":s_date, "ed":e_date})

def delete_notice(p):
    db_execute("DELETE FROM product_notices WHERE product_name = :p", {"p": p})

def get_all_active_notices(t_str):
    rows = db_fetchall('SELECT product_name, notice_text FROM product_notices WHERE start_date <= :t AND end_date >= :t', {"t":t_str})
    return {r[0]: r[1] for r in rows}

def get_raw_notice(p):
    return db_fetchone("SELECT notice_text, start_date, end_date FROM product_notices WHERE product_name=:p", {"p":p})

def save_to_db(d_date, eq, wk, p, tgt, meas, diff, st_val, rmks, amt):
    ts = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
    db_execute('''INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount) VALUES (:ts, :pd, :eq, :wk, :p, :tgt, :meas, :diff, :st_val, :rmks, :amt)''', 
              {"ts":ts, "pd":d_date, "eq":str(eq).strip(), "wk":str(wk).strip(), "p":str(p).strip(), "tgt":tgt, "meas":meas, "diff":diff, "st_val":st_val, "rmks":rmks, "amt":amt})

def check_recent_duplicate(d_date, eq, p, meas_val):
    r = db_fetchone('SELECT measured_value, timestamp FROM color_records WHERE production_date=:pd AND equipment=:eq AND product_name=:p ORDER BY id DESC LIMIT 1', {"pd":d_date, "eq":str(eq).strip(), "p":str(p).strip()})
    if r and float(r[0]) == float(meas_val):
        try:
            if (get_now_kst() - datetime.strptime(r[1], "%Y-%m-%d %H:%M:%S")).total_seconds() < 30: return True 
        except: pass
    return False

def get_last_record(p):
    return db_fetchone('SELECT production_date, measured_value, status FROM color_records WHERE product_name = :p ORDER BY production_date DESC, timestamp DESC LIMIT 1', {"p":str(p).strip()})

def get_equipment_last_records(p_name):
    query = """
        WITH RankedRecords AS (SELECT equipment, production_date, measured_value, status, id, ROW_NUMBER() OVER (PARTITION BY equipment ORDER BY production_date DESC, timestamp DESC, id DESC) as rn FROM color_records WHERE product_name = :p), 
        EquipCounts AS (SELECT equipment, COUNT(*) as cnt FROM color_records WHERE product_name = :p GROUP BY equipment)
        SELECT r.equipment, r.production_date, r.measured_value, r.status, c.cnt FROM RankedRecords r JOIN EquipCounts c ON r.equipment = c.equipment WHERE r.rn = 1 ORDER BY c.cnt DESC, r.production_date DESC
    """
    return db_fetchall(query, {"p":str(p_name).strip()})

def auto_fill_input_amount(row):
    eq = str(row['생산설비']).lower().replace(" ", "")
    amt = str(row['투입량']).strip()
    if '버닝' in eq: return amt if amt != "" else "-"
    if amt in ["", "-", "nan", "None"]:
        if "태환" in eq: return "12kg"
        elif "프로밧" in eq: return "25kg"
        elif "60" in eq: return "60kg"
        elif "120" in eq: return "125kg" # [수정] 뷸러120kg -> 125kg 적용
    return amt

@st.cache_data(show_spinner=False, ttl=600)
def load_from_db():
    q = """
    SELECT 
        c.id as 고유번호, c.timestamp as 입력일시, c.production_date as 생산일, 
        c.equipment as 생산설비, COALESCE(c.input_amount, '-') as 투입량, 
        c.worker as 작업자, c.product_name as 제품명, c.measured_value as 측정색도, 
        COALESCE(c.remarks, '') as 특이사항, COALESCE(c.checked, 0) as checked_status, 
        COALESCE(
            (SELECT target_value FROM target_history th WHERE th.product_name = c.product_name AND th.effective_date <= c.production_date ORDER BY th.effective_date DESC LIMIT 1), 
            (SELECT target_value FROM target_history th WHERE th.product_name = c.product_name ORDER BY th.effective_date ASC LIMIT 1), 
            0.0
        ) as 기준색도 
    FROM color_records c
    """
    try: 
        df = pd.read_sql_query(text(q), get_engine())
    except Exception: 
        return pd.DataFrame(columns=['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '투입량', '판정', '확인여부', '특이사항', '입력일시', '고유번호'])

    if df.empty:
        return pd.DataFrame(columns=['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '투입량', '판정', '확인여부', '특이사항', '입력일시', '고유번호'])

    df['생산일'] = df['생산일'].apply(safe_date_parse)
    df['제품명'] = df['제품명'].astype(str).str.strip()
    df['생산설비'] = df['생산설비'].astype(str).str.strip()
    df['작업자'] = df['작업자'].astype(str).str.strip().replace(['nan', 'None', '', 'NaN'], '미입력(과거기록)')
    df['투입량'] = df.apply(auto_fill_input_amount, axis=1)
    df['확인여부'] = df['checked_status'].apply(lambda x: "확인완료 ✅" if x == 1 else "미확인 ❌")

    df['측정색도'] = pd.to_numeric(df['측정색도'], errors='coerce')
    df['기준색도'] = pd.to_numeric(df['기준색도'], errors='coerce')
    df['오차'] = (df['측정색도'] - df['기준색도'])
    df['판정'] = "합격 🟢"
    df.loc[df['오차'].abs() > 2.0, '판정'] = "불합격 🔴"
    df.loc[df['오차'].isna(), '판정'] = "오류"
    
    df['특이사항'] = df['특이사항'].fillna('').astype(str)
    df['특이사항'] = df['특이사항'].str.replace(r'\[설비 첫\s*배치\s*🚀\]\s*', '', regex=True)
    df['특이사항'] = df['특이사항'].str.replace(r'\[마지막 배치\s*🏁\]\s*', '', regex=True)
    df['특이사항'] = df['특이사항'].str.replace(r'\[기준값 변경 후 첫 생산\s*🔔\]\s*', '', regex=True)
    df['특이사항'] = df['특이사항'].str.replace("nan", "", regex=False).str.strip()

    df = df.sort_values(by=['생산일', '입력일시', '고유번호'], ascending=[False, False, False]).reset_index(drop=True)

    try:
        th_df = pd.read_sql_query(text("SELECT product_name, effective_date FROM target_history WHERE effective_date NOT IN ('2000-01-01', '2024-04-11', '')"), get_engine())
        target_change_first_ids = set()
        for _, r in th_df.iterrows():
            sub = df[(df['제품명'] == r['product_name'].strip()) & (df['생산일'] >= r['effective_date'])]
            if not sub.empty: target_change_first_ids.add(sub.iloc[-1]['고유번호'])
        
        if target_change_first_ids:
            mask = df['고유번호'].isin(target_change_first_ids)
            df.loc[mask, '특이사항'] = "[기준값 변경 후 첫 생산 🔔] " + df.loc[mask, '특이사항']
    except: pass

    df['norm_p'] = df['제품명'].str.replace(" ", "").str.lower()
    df['norm_e'] = df['생산설비'].str.replace(" ", "").str.lower()
    
    f_idx = df.groupby(['norm_p', 'norm_e']).tail(1).index
    df.loc[f_idx, '특이사항'] = "[설비 첫 배치 🚀] " + df.loc[f_idx, '특이사항']
    
    df = df.drop(columns=['norm_p', 'norm_e'])
    
    l_idx = df.groupby('생산일').head(1).index
    df.loc[l_idx, '특이사항'] = "[마지막 배치 🏁] " + df.loc[l_idx, '특이사항']
    
    df['특이사항'] = df['특이사항'].str.strip()
    return df[['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '투입량', '판정', '확인여부', '특이사항', '입력일시', '고유번호']]

@st.cache_data(show_spinner=False)
def get_ai_predictions():
    df = load_from_db()
    if df.empty: return []
    
    predict_data = []
    today_d = get_now_kst().date()
    df['생산일_dt'] = pd.to_datetime(df['생산일'], errors='coerce').dt.date
    
    for prod, group in df.groupby('제품명'):
        unique_dates = sorted(group['생산일_dt'].dropna().drop_duplicates().tolist())
        if len(unique_dates) < 2: continue
        
        last_date = unique_dates[-1]
        if (today_d - last_date).days > 120: continue 
        
        intervals = [int(np.busday_count(str(unique_dates[i-1]), str(unique_dates[i]), holidays=KR_HOLIDAYS)) for i in range(1, len(unique_dates))]
        if not intervals: continue
        avg_interval = max(1, sum(intervals) / len(intervals))
        
        next_date_np = np.busday_offset(str(last_date), int(avg_interval), roll='forward', holidays=KR_HOLIDAYS)
        next_date = pd.to_datetime(next_date_np).date()
        d_day = int(np.busday_count(str(today_d), str(next_date), holidays=KR_HOLIDAYS))
        
        if d_day < 0: status_str = f"🚨 긴급 ({-d_day}영업일 지남)"
        elif d_day == 0: status_str = "🔥 오늘 생산 권장"
        elif d_day <= 3: status_str = f"⚠️ D-{d_day} (임박)"
        else: status_str = f"✅ D-{d_day} (여유)"
        
        predict_data.append({
            "제품명": prod, "마지막 생산일": last_date.strftime("%Y-%m-%d"),
            "평균 생산 주기": f"약 {int(avg_interval)}영업일", "다음 예상일": next_date.strftime("%Y-%m-%d"),
            "생산 필요 상태": status_str, "_sort": d_day
        })
    return predict_data

@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w: df.to_excel(w, index=False, sheet_name='기록')
    return output.getvalue()

init_db() 
CURRENT_WORKERS = get_all_workers()

@st.cache_data
def load_tgt():
    rows = db_fetchall("SELECT product_name, target_value FROM target_history ORDER BY effective_date ASC, id ASC")
    if rows:
        return {r[0]: r[1] for r in rows}
    else:
        try: 
            df = pd.read_excel(EXCEL_FILE, usecols="C:D", header=1).dropna()
            targets = {str(r.iloc[0]).strip(): float(r.iloc[1]) if not pd.isna(r.iloc[1]) else 0.0 for i, r in df.iterrows()}
            for p, v in targets.items():
                db_execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (:p, :v, '2000-01-01')", {"p":p, "v":v})
            return targets
        except: 
            return {"(데이터 없음)": 0.0}

TARGET_DATA = load_tgt()
today_str_kst = get_now_kst().strftime("%Y-%m-%d")
ACTIVE_NOTICES = get_all_active_notices(today_str_kst)

# ----------------------------------------------------
# 3. 관리자 전용 메뉴 
# ----------------------------------------------------
@st.dialog("🛠️ 관리자 전용 메뉴", width="large")
def admin_menu_dialog():
    input_pw_admin = st.text_input("🔒 비밀번호를 입력하세요", type="password", key="admin_pw_input")
    
    if input_pw_admin == ADMIN_PASSWORD:
        t1, t2, t3, t4, t5, t6, t7, t8, t9 = st.tabs(["🔍 금일 확인", "📝 수정/삭제", "📂 과거기록 업로드", "📅 제품기준 적용", "📢 공지", "⏳ 미생산", "👥 통계", "🧑‍🔧 데이터 정화", "🔮 AI 예측"])
        
        with t1:
            st.info("오늘 생산된 배치 확인 관리")
            tdf = history_df[history_df['생산일'] == today_str_kst]
            if tdf.empty: st.info("기록 없음")
            else:
                mode = st.radio("보기", ["미확인 ❌", "확인완료 ✅"], horizontal=True, key="admin_view_mode")
                sub_df = tdf[tdf['확인여부'] == mode]
                st.dataframe(sub_df, hide_index=True)
                if mode == "미확인 ❌":
                    to_chk = st.multiselect("확인 처리할 내역", sub_df['고유번호'].tolist(), key="admin_sel_confirm")
                    if st.button("✅ 선택 확인 완료", key="admin_btn_confirm"):
                        update_checked_status(to_chk, 1); st.cache_data.clear(); st.session_state['show_toast'] = "확인 완료!"; st.rerun()
                else:
                    to_unchk = st.multiselect("미확인 복구 내역", sub_df['고유번호'].tolist(), key="admin_sel_unconfirm")
                    if st.button("🔄 선택 복구", key="admin_btn_unconfirm"):
                        update_checked_status(to_unchk, 0); st.cache_data.clear(); st.session_state['show_toast'] = "복구 완료!"; st.rerun()
        with t2:
            col1, col2 = st.columns(2)
            with col1:
                tid = st.number_input("고유번호", min_value=1, key="admin_num_id")
                act = st.radio("작업", ["삭제", "수정"], key="admin_radio_action")
            with col2:
                if act == "삭제" and st.button("🗑️ 데이터 삭제", key="admin_btn_del_record"):
                    delete_from_db(tid); st.cache_data.clear(); st.session_state['show_toast'] = "삭제됨!"; st.rerun()
                elif act == "수정":
                    row = db_fetchone("SELECT product_name, target_value, production_date, equipment, worker, measured_value, remarks, input_amount, COALESCE(checked, 0) FROM color_records WHERE id=:id", {"id":tid})
                    if row:
                        try: def_date = datetime.strptime(row[2], "%Y-%m-%d").date()
                        except: def_date = get_now_kst().date()
                        npd = st.date_input("생산일", value=def_date, key="admin_date_edit").strftime("%Y-%m-%d")
                        
                        opts = list(TARGET_DATA.keys())
                        if row[0] not in opts: opts.append(row[0])
                        
                        nprod = st.selectbox("제품", opts, index=opts.index(row[0]), key="admin_sel_prod")
                        neq = st.selectbox("설비", EQUIPMENT_LIST, index=EQUIPMENT_LIST.index(row[3]) if row[3] in EQUIPMENT_LIST else 0, key="admin_sel_equip")
                        
                        # [수정] 수정 모드에서도 120kg은 125kg으로 표시
                        namt = st.selectbox("투입량", ["1.35kg","2.5kg","3.75kg"], index=["1.35kg","2.5kg","3.75kg"].index(row[7]) if row[7] in ["1.35kg","2.5kg","3.75kg"] else 0, key="admin_sel_amt") if "버닝" in neq else ("12kg" if "태환" in neq else "25kg" if "프로밧" in neq else "60kg" if "60" in neq else "125kg" if "120" in neq else "-")
                        
                        nw = st.selectbox("작업자", CURRENT_WORKERS, index=CURRENT_WORKERS.index(row[4]) if row[4] in CURRENT_WORKERS else 0, key="admin_sel_worker")
                        nm = st.number_input("측정", value=float(row[5]), step=0.1, key="admin_num_meas")
                        nrm = st.text_input("특이사항", value=row[6], key="admin_txt_rmk")
                        
                        if st.button("✏️ 수정 완료", key="admin_btn_edit_record"):
                            tgt = get_historical_target(nprod, npd)
                            diff = round(nm - tgt, 1)
                            stat = "합격 🟢" if abs(diff)<=2.0 else "불합격 🔴"
                            update_db(tid, npd, neq, nw, nprod, tgt, nm, diff, stat, nrm, namt, row[8])
                            st.cache_data.clear(); st.session_state['show_toast'] = "수정됨!"; st.rerun()
        
        with t3:
            st.info("과거 생산/측정 기록이 담긴 엑셀 데이터를 일괄 업로드합니다.")
            up = st.file_uploader("과거 기록 엑셀 업로드", type=['xlsx', 'xls'], key="admin_file_record")
            if up and st.button("🚀 기록 일괄 업로드", key="admin_btn_upload_record"):
                try:
                    df_up = pd.read_excel(up)
                    if all(c in df_up.columns for c in ['생산일', '제품명', '생산설비', '작업자', '측정색도']):
                        for _, r in df_up.iterrows():
                            if str(r['측정색도']).strip() in ['-', '', 'nan', 'None']: continue
                            try: meas = float(str(r['측정색도']).strip())
                            except: continue
                            
                            p_dt = safe_date_parse(r['생산일'])
                            if not p_dt: continue 
                            
                            pd_name, eq, wk = str(r['제품명']).strip(), str(r['생산설비']).strip(), str(r['작업자']).strip()
                            
                            # [수정] 과거 엑셀 업로드 시에도 120이면 무조건 125kg으로 기록
                            am = str(r.get('투입량', '')).strip() if '버닝' in eq.lower() else ("12kg" if "태환" in eq else "25kg" if "프로밧" in eq else "60kg" if "60" in eq else "125kg" if "120" in eq else "-")
                            
                            rm = str(r.get('특이사항', '')).strip()
                            tgt = get_historical_target(pd_name, p_dt)
                            diff = round(meas - tgt, 1)
                            stat = "합격 🟢" if abs(diff)<=2.0 else "불합격 🔴"
                            
                            db_execute('''INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount) VALUES (:ts, :pd, :eq, :wk, :p, :tgt, :meas, :diff, :st, :rm, :amt)''',
                                      {"ts":get_now_kst().strftime("%Y-%m-%d %H:%M:%S"), "pd":p_dt, "eq":eq, "wk":wk, "p":pd_name, "tgt":tgt, "meas":meas, "diff":diff, "st":stat, "rm":rm if rm not in ['nan','None'] else '', "amt":am})
                        st.cache_data.clear(); st.session_state['show_toast'] = "과거 기록 업로드 성공!"; st.rerun()
                    else:
                        st.error("엑셀에 필수 열('생산일', '제품명', '생산설비', '작업자', '측정색도')이 부족합니다.")
                except Exception as e: st.error(f"오류: {e}")
            
            st.markdown("---")
            st.error("🛠️ **DB 고유번호 꼬임 해결 (초기화 및 재정렬)**")
            st.caption("과거 데이터를 나중에 업로드하여 고유번호(순서)가 날짜와 맞지 않게 꼬였을 때, 아래 버튼을 누르면 생산일자 순으로 고유번호를 깔끔하게 재정렬합니다.")
            if st.button("🧹 고유번호 날짜순 전면 재정렬", type="primary", key="admin_btn_reindex"):
                try:
                    df_all = pd.read_sql_query(text("SELECT * FROM color_records"), get_engine())
                    if not df_all.empty:
                        df_all = df_all.sort_values(by=['production_date', 'id'], ascending=[True, True]).reset_index(drop=True)
                        df_all['id'] = df_all.index + 1
                        
                        db_execute("TRUNCATE TABLE color_records RESTART IDENTITY;")
                        df_all.to_sql('color_records', get_engine(), if_exists='append', index=False)
                        
                        st.cache_data.clear()
                        st.session_state['show_toast'] = "고유번호 전면 재정렬 완료! 순서가 정상화되었습니다."
                        st.rerun()
                    else:
                        st.info("데이터가 없습니다.")
                except Exception as e:
                    st.error(f"재정렬 중 오류 발생: {e}")
        
        with t4:
            st.info("제품별 기준 색도 엑셀 파일을 업로드하여 시스템에 즉시 적용합니다.")
            h_up = st.file_uploader("제품 기준값/이력 엑셀 업로드", type=['xlsx','xls'], key="admin_file_history")
            if h_up and st.button("🚀 제품 기준값 전체 적용", key="admin_btn_upload_history"):
                try:
                    df_h = pd.read_excel(h_up)
                    if all(c in df_h.columns for c in ['제품명','적용시작일','기준색도']):
                        db_execute("TRUNCATE TABLE target_history RESTART IDENTITY;")
                        for _, r in df_h.iterrows():
                            dt = safe_date_parse(r['적용시작일']) or '2000-01-01'
                            try: db_execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (:p, :v, :d)", {"p":str(r['제품명']).strip(), "v":float(r['기준색도']), "d":dt})
                            except: pass
                        st.cache_data.clear(); 
                        st.session_state['show_toast'] = "제품 기준값이 시스템 전체에 즉시 적용되었습니다!"
                        st.rerun()
                    else:
                        st.error("엑셀 파일에 '제품명', '적용시작일', '기준색도' 열이 포함되어 있어야 합니다.")
                except Exception as e:
                    st.error(f"적용 중 오류 발생: {e}")
                    
        with t5:
            np_prod = st.selectbox("제품", list(TARGET_DATA.keys()), key="admin_notice_prod")
            rn = get_raw_notice(np_prod)
            ntxt = st.text_area("내용", value=rn[0] if rn else "", key="admin_notice_text")
            c1, c2 = st.columns(2)
            with c1: sd = st.date_input("시작", value=datetime.strptime(rn[1], "%Y-%m-%d").date() if rn else get_now_kst().date(), key="admin_notice_sd")
            with c2: 
                nol = st.checkbox("무기한", value=(rn[2]=="2099-12-31") if rn else False, key="admin_notice_unlimit")
                ed = st.date_input("종료", disabled=nol, key="admin_notice_ed")
            if st.button("📢 공지 등록/수정", type="primary", key="admin_btn_save_notice"):
                save_notice(np_prod, ntxt, sd.strftime("%Y-%m-%d"), "2099-12-31" if nol else ed.strftime("%Y-%m-%d"))
                st.cache_data.clear(); st.session_state['show_toast'] = "공지 등록!"; st.rerun()
            if st.button("🗑️ 공지 삭제", key="admin_btn_del_notice"): delete_notice(np_prod); st.cache_data.clear(); st.rerun()
        with t6:
            inact = []
            td = get_now_kst().date()
            for p in TARGET_DATA.keys():
                lr = get_last_record(p)
                if lr:
                    try:
                        d = (td - datetime.strptime(lr[0], "%Y-%m-%d").date()).days
                        if d >= 120: inact.append({"제품명":p, "최종 생산":lr[0], "경과":f"{d}일"})
                    except: pass
                else: inact.append({"제품명":p, "최종 생산":"없음", "경과":"이력 없음"})
            st.dataframe(pd.DataFrame(inact), use_container_width=True, hide_index=True)
        with t7:
            if not history_df.empty:
                ws = []
                for nm, grp in history_df.groupby('작업자'):
                    tc, fc = len(grp), len(grp[grp['판정'].str.contains("불합격", na=False)])
                    ws.append({"작업자":nm, "총":tc, "합격":tc-fc, "불합격":fc, "불량률(%)":fc/tc*100 if tc>0 else 0, "오차(절대)":grp['오차'].abs().mean()})
                st.dataframe(pd.DataFrame(ws).sort_values(by="총", ascending=False).style.format({"불량률(%)":"{:.1f}%", "오차(절대)":"{:.2f}"}), hide_index=True)
        with t8:
            st.info("작업자 명단 관리 및 과거 엑셀 데이터 명칭 강제 통일화 도구입니다.")
            c_w1, c_w2 = st.columns(2)
            with c_w1:
                nw = st.text_input("새 작업자 이름", key="admin_new_worker")
                if st.button("➕ 작업자 추가", type="primary", key="admin_btn_add_worker") and add_worker(nw): st.cache_data.clear(); st.session_state['show_toast'] = "작업자 추가!"; st.rerun()
            with c_w2:
                if CURRENT_WORKERS:
                    dw = st.selectbox("기존 작업자", CURRENT_WORKERS, key="admin_del_worker_sel")
                    if st.button("➖ 작업자 삭제", key="admin_btn_del_worker"): delete_worker(dw); st.cache_data.clear(); st.session_state['show_toast'] = "작업자 삭제!"; st.rerun()
            
            st.markdown("---")
            st.error("🛠️ **DB 명칭 불일치 해결 (과거 데이터 통합)**")
            st.caption("과거 엑셀로 업로드한 기록의 설비명이나 제품명에 미세한 오타/띄어쓰기가 있어 같은 제품으로 인식되지 않을 때, 시스템 기준으로 강제 통일시킵니다.")
            if st.button("✨ 데이터 명칭 전면 통일화 (첫 배치 오류 완전 해결)", type="primary", use_container_width=True, key="admin_btn_clean_db"):
                db_execute("UPDATE color_records SET worker = TRIM(worker), product_name = TRIM(product_name)")
                db_execute("UPDATE color_records SET equipment = '버닝' WHERE REPLACE(equipment, ' ', '') LIKE :v", {"v": "%버닝%"})
                db_execute("UPDATE color_records SET equipment = '태환 12kg' WHERE REPLACE(equipment, ' ', '') LIKE :v", {"v": "%태환%"})
                db_execute("UPDATE color_records SET equipment = '프로밧 25kg' WHERE REPLACE(equipment, ' ', '') LIKE :v", {"v": "%프로밧%"})
                db_execute("UPDATE color_records SET equipment = '뷸러 60kg' WHERE REPLACE(equipment, ' ', '') LIKE :v", {"v": "%60%"})
                db_execute("UPDATE color_records SET equipment = '뷸러 120kg' WHERE REPLACE(equipment, ' ', '') LIKE :v", {"v": "%120%"})
                st.cache_data.clear(); st.session_state['show_toast'] = "데이터 명칭 100% 통일화 완료!"; st.rerun()
                
            # =========================================================
            # [핵심 추가] 뷸러 120kg 과거 기록 투입량 일괄 수정 버튼
            # =========================================================
            st.markdown("---")
            st.error("🛠️ **투입량 일괄 수정 (뷸러 120kg -> 125kg)**")
            st.caption("과거 기록 중 뷸러 120kg 설비의 투입량을 기존 120kg에서 실제 투입량인 125kg으로 일괄 변경합니다.")
            if st.button("✨ 뷸러 120kg 투입량 '125kg'으로 일괄 수정", type="primary", use_container_width=True, key="admin_btn_fix_125"):
                db_execute("UPDATE color_records SET input_amount = '125kg' WHERE equipment LIKE :v", {"v": "%120%"})
                st.cache_data.clear(); st.session_state['show_toast'] = "투입량 125kg 일괄 수정 완료!"; st.rerun()
            # =========================================================

        with t9:
            st.info("최근 4개월(120일) 이내에 2회 이상 생산된 제품들의 영업일 기준 평균 생산 주기 분석")
            pred_data = get_ai_predictions()
            if pred_data:
                pred_df = pd.DataFrame(pred_data).sort_values('_sort').drop(columns=['_sort'])
                def hl_pred(s):
                    colors = []
                    for v in s:
                        if '긴급' in str(v) or '오늘' in str(v): colors.append('background-color: #FADBD8; color: black; font-weight: bold;')
                        elif '임박' in str(v): colors.append('background-color: #FCF3CF; color: black; font-weight: bold;')
                        else: colors.append('')
                    return colors
                st.dataframe(pred_df.style.apply(hl_pred, subset=['생산 필요 상태']).set_properties(**{'text-align': 'center'}), use_container_width=True, hide_index=True)
            else:
                st.success("데이터가 부족하여 아직 예측할 수 없습니다.")
    elif input_pw_admin != "": st.error("❌ 비밀번호 불일치")

# ----------------------------------------------------
# 4. 메인 화면 출력부
# ----------------------------------------------------
history_df = load_from_db()

c1, c2, c3 = st.columns([7, 1.5, 1])
with c1: st.title("🎨 일일 제품 색도 관리 시스템")
with c2: 
    if st.button("🛠️ 관리자 메뉴", use_container_width=True): admin_menu_dialog()
with c3:
    if st.button("🔒 로그아웃", use_container_width=True):
        st.query_params.clear(); st.session_state['logged_in'] = False; st.rerun()

st.markdown("---")
st.subheader("📝 데이터 등록")

tab_n, tab_q = st.tabs(["📋 일반 데이터 등록", "⚡ 진행 중인 라인 빠른 추가"])

with tab_n:
    with st.container(border=True):
        cs1, cs2, cs3, cs4 = st.columns(4)
        with cs1: prod_date_str = st.date_input("생산일 선택", value=get_now_kst().date(), key="main_date").strftime("%Y-%m-%d")
        with cs2: 
            selected_equipment = st.selectbox("생산 설비 선택", EQUIPMENT_LIST, key="main_equip")
            equip_clean = str(selected_equipment).lower().replace(" ", "")
        with cs3: worker_name = st.selectbox("작업자 선택", CURRENT_WORKERS if CURRENT_WORKERS else [""], key="main_worker")
        with cs4:
            if "버닝" in equip_clean: input_amount_val = st.selectbox("원료 투입량", ["1.35kg", "2.5kg", "3.75kg"], key="main_amt_sel")
            else:
                # [수정] 뷸러 120kg 선택 시 투입량을 125kg으로 자동 고정
                input_amount_val = "12kg" if "태환" in equip_clean else "25kg" if "프로밧" in equip_clean else "60kg" if "60" in equip_clean else "125kg" if "120" in equip_clean else "-"
                st.text_input("투입량 (고정)", input_amount_val, disabled=True, key="main_amt_txt")
                
        col_p1, col_p2 = st.columns([2, 1])
        with col_p1:
            selected_product = st.selectbox(
                "🔍 제품명 검색 및 선택", 
                list(TARGET_DATA.keys()), 
                index=None, 
                placeholder="제품명을 클릭하여 검색하세요 (우측 'X' 버튼으로 즉시 지 지우기)", 
                key="main_prod"
            )
            if selected_product and ACTIVE_NOTICES.get(selected_product): 
                st.warning(f"📢 **전달사항:** {ACTIVE_NOTICES[selected_product]}")
                
        with col_p2:
            target_value = get_historical_target(selected_product, prod_date_str) if selected_product else 0.0
            st.info(f"📌 해당 생산일({prod_date_str}) 기준 색도: **{float(target_value):.1f}**")
        
        if selected_product:
            last_records = get_equipment_last_records(selected_product)
            if last_records:
                valid_records, very_old_records = [], []
                today_date = get_now_kst().date()
                for row in last_records:
                    try: days_passed = (today_date - datetime.strptime(row[1], "%Y-%m-%d").date()).days
                    except: days_passed = 0
                    if days_passed >= 365: very_old_records.append(row[0])
                    else: valid_records.append(row)
                if very_old_records:
                    st.error(f"🚨 **[경고] 1년 이상 장기 미생산 알림!**\n\n다음 설비에서 1년 이상 생산된 적이 없습니다: **{', '.join(very_old_records)}**")
                if valid_records:
                    st.caption("💡 **설비별 최근 생산 이력**")
                    cols = st.columns(len(valid_records[:3]))
                    for idx, row in enumerate(valid_records[:3]):
                        disp_date = f":red[**{row[1]} (4개월 초과)**]" if (today_date - datetime.strptime(row[1], "%Y-%m-%d").date()).days > 120 else row[1]
                        disp_meas = f":red[**{float(row[2]):.1f} (이전 불합격!)**]" if "불합격" in row[3] else f"{float(row[2]):.1f}"
                        with cols[idx]: st.info(f"⚙️ **{row[0]}**\n\n🕒 {disp_date}\n\n📉 {disp_meas}")
            else: st.warning("이전 생산 기록이 없습니다.")
        else:
            st.info("👆 위에서 제품을 선택하시면 과거 설비별 생산 이력이 표시됩니다.")

        cs8, cs9, cs10 = st.columns([2,2,1])
        with cs8: measured_value = st.number_input("측정 색도 입력", value=float(target_value), step=0.1, key="main_meas")
        with cs9: remarks_input = st.text_input("특이사항 (선택사항)", placeholder="메모 입력", key="main_rmk")
        with cs10:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("데이터 등록하기", type="primary", use_container_width=True, key="main_btn_save"):
                if not selected_product: st.warning("⚠️ 제품명을 먼저 선택해주세요!")
                elif not worker_name: st.warning("⚠️ 작업자 오류!")
                elif check_recent_duplicate(prod_date_str, selected_equipment, selected_product, measured_value): st.error("⚠️ 중복 데이터!")
                else:
                    diff = round(measured_value - target_value, 1)
                    save_to_db(prod_date_str, selected_equipment, worker_name, selected_product, target_value, measured_value, diff, "합격 🟢" if abs(diff)<=2.0 else "불합격 🔴", remarks_input, input_amount_val)
                    st.cache_data.clear(); st.session_state['show_toast'] = "정상 등록 완료!"; st.rerun()

with tab_q:
    with st.container(border=True):
        tdf = history_df[history_df['생산일'] == today_str_kst]
        if tdf.empty: st.info("오늘 첫 생산을 일반 탭에서 진행해주세요.")
        else:
            rb = tdf[['제품명','생산설비','투입량','작업자']].drop_duplicates().reset_index(drop=True)
            opts = [f"▶ {r['제품명']} ({r['생산설비']} / {r['투입량']} / {r['작업자']})" for _,r in rb.iterrows()]
            cq1, cq2, cq3, cq4 = st.columns([3,1,1,1])
            with cq1: sb = st.selectbox("이어서 측정할 제품", opts, key="quick_sel")
            if sb:
                idx = opts.index(sb)
                qp, qe, qa, qw = rb.iloc[idx]['제품명'], rb.iloc[idx]['생산설비'], rb.iloc[idx]['투입량'], rb.iloc[idx]['작업자']
                qt = get_historical_target(qp, today_str_kst)
                with cq2: st.text_input("기준", f"{float(qt):.1f}", disabled=True, key="quick_tgt")
                with cq3: qm = st.number_input("측정값", value=float(qt), step=0.1, key="quick_meas")
                with cq4:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🚀 1초 빠른 등록", type="primary", use_container_width=True, key="quick_btn"):
                        if check_recent_duplicate(today_str_kst, qe, qp, qm): st.error("⚠️ 중복")
                        else:
                            diff = round(qm - qt, 1)
                            save_to_db(today_str_kst, qe, qw, qp, qt, qm, diff, "합격 🟢" if abs(diff)<=2.0 else "불합격 🔴", "", qa)
                            st.cache_data.clear(); st.session_state['show_toast'] = "빠른 등록 완료!"; st.rerun()

st.markdown("---")
st.subheader("📊 누적 측정 기록 조회")

if st_autorefresh:
    if st.checkbox("🔄 실시간 모니터링 켜기 (10초)", key="arf_chk"): st_autorefresh(interval=10000)

cf1, cf2, cf3 = st.columns(3)
with cf1: sq = st.text_input("🔍 검색", key="filter_sq").strip()
with cf2: dm = st.radio("📅 기간", ["오늘", "전체", "특정 일자"], horizontal=True, key="filter_dm")
fd_str = ""
with cf3:
    if dm == "특정 일자": fd_str = st.date_input("선택", key="filter_date").strftime("%Y-%m-%d")

ddf = history_df.copy()
if not ddf.empty:
    if sq: ddf = ddf[ddf['제품명'].astype(str).str.contains(sq)]
    if dm == "오늘": ddf = ddf[ddf['생산일'] == today_str_kst]
    elif dm == "특정 일자": ddf = ddf[ddf['생산일'] == fd_str]

    if not ddf.empty:
        eq_map = {'버닝': 0, '태환12kg': 1, '프로밧25kg': 2, '뷸러60kg': 3, '뷸러120kg': 4}
        ddf['s'] = ddf['생산설비'].astype(str).str.replace(" ", "").str.lower().map(lambda x: eq_map.get(x, 5))
        
        ddf = ddf.sort_values(by=['s', '생산일', '고유번호'], ascending=[True, False, True])
        ddf = ddf.drop(columns=['s'])

if not ddf.empty:
    tb = len(ddf)
    mt = "오늘" if dm=="오늘" else fd_str if dm=="특정 일자" else "전체"
    ec = ddf['생산설비'].value_counts()
    pe = [e for e in EQUIPMENT_LIST if e in ec.index]
    
    mc = st.columns(1 + len(pe))
    with mc[0]: st.metric(f"📦 {mt} 배치", f"{tb} 건")
    for i, e in enumerate(pe):
        with mc[i+1]: st.metric(f"⚙️ {e}", f"{ec[e]} 건")
    st.markdown("<br>", unsafe_allow_html=True)

def hl_stat(s): return ['color: white; background-color: #E74C3C; font-weight: bold;' if '불합격' in str(v) else 'color: #27AE60; font-weight: bold;' for v in s]
def hl_eq(s):
    clrs = []
    for v in s:
        c = str(v).replace(" ","").lower()
        if '버닝' in c: clrs.append('background-color: #E1F5FE; color: black; font-weight: bold;') 
        elif '태환' in c: clrs.append('background-color: #FFF3CD; color: black; font-weight: bold;') 
        elif '프로밧' in c: clrs.append('background-color: #FCE4EC; color: black; font-weight: bold;') 
        elif '60' in c: clrs.append('background-color: #E8F5E9; color: black; font-weight: bold;') 
        elif '120' in c: clrs.append('background-color: #D4EFDF; color: black; font-weight: bold;') 
        else: clrs.append('')
    return clrs

if not ddf.empty:
    page_size = 100
    total_pages = max(1, int(np.ceil(len(ddf) / page_size)))
    page = st.number_input("📄 페이지 선택", min_value=1, max_value=total_pages, value=1)
    
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    page_df = ddf.iloc[start_idx:end_idx].copy()
    
    mdf = page_df.drop(columns=['확인여부'], errors='ignore')
    
    sdf = mdf.style.format({"측정색도":"{:.1f}", "기준색도":"{:.1f}", "오차":"{:.1f}"}, na_rep="-") \
                   .apply(hl_eq, subset=['생산설비']) \
                   .apply(hl_stat, subset=['판정']) \
                   .set_properties(subset=['특이사항'], **{'background-color': '#E8DAEF', 'color': 'black', 'font-weight': 'bold'}) \
                   .set_properties(subset=['제품명'], **{'font-weight': 'bold'})
    
    st.markdown(sdf.to_html(), unsafe_allow_html=True)
    
    fn = f"색도측정_{today_str_kst if dm=='오늘' else fd_str if dm=='특정 일자' else '전체'}.xlsx"
    full_mdf = ddf.drop(columns=['확인여부'], errors='ignore')
    st.download_button("📥 엑셀 전체 다운로드", to_excel(full_mdf), fn, key="btn_download_excel")
else: 
    st.info("🔍 일치하는 기록이 없습니다.")
