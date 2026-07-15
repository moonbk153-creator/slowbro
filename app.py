import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import io 
import pytz
import numpy as np

try:
    import holidays
    HAS_HOLIDAYS = True
    KR_HOLIDAYS = holidays.KR(years=range(2020, 2035))
except: 
    HAS_HOLIDAYS = False
    KR_HOLIDAYS = []

try: from streamlit_autorefresh import st_autorefresh
except: st_autorefresh = None

KST = pytz.timezone('Asia/Seoul')
st.set_page_config(page_title="색도 관리 시스템", layout="wide")

EXCEL_FILE, DB_FILE = 'data sheet.xlsx', 'color_management.db'
EQUIPMENT_LIST = ["버닝", "태환 12kg", "프로밧 25kg", "뷸러 60kg", "뷸러 120kg"]
ADMIN_PASSWORD = st.secrets["ADMIN_PASSWORD"]
ACCESS_PASSWORD = st.secrets["APP_PASSWORD"]

if 'show_toast' in st.session_state:
    st.toast(st.session_state['show_toast'], icon="✅")
    del st.session_state['show_toast']
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
            st.rerun()
        else: st.error("❌ 비밀번호 불일치")
    st.stop()

def get_now_kst(): return datetime.now(KST)
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS color_records (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, production_date TEXT, equipment TEXT, worker TEXT, product_name TEXT, target_value REAL, measured_value REAL, difference REAL, status TEXT, remarks TEXT, input_amount TEXT, checked INTEGER DEFAULT 0)''')
    c.execute('''CREATE TABLE IF NOT EXISTS target_history (id INTEGER PRIMARY KEY AUTOINCREMENT, product_name TEXT, target_value REAL, effective_date TEXT)''')
    conn.commit(); conn.close()

def load_from_db():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM color_records ORDER BY id DESC", conn)
    conn.close()
    return df

def save_to_db(d_date, eq, wk, p, tgt, meas, diff, st_val, rmks, amt):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('INSERT INTO color_records (production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', 
              (d_date, eq, wk, p, tgt, meas, diff, st_val, rmks, amt))
    conn.commit(); conn.close()

init_db()
TARGET_DATA = {str(r.iloc[0]): float(r.iloc[1]) for i, r in pd.read_excel(EXCEL_FILE, usecols="C:D", header=1).dropna().iterrows()}
@st.dialog("🛠️ 관리자 전용 메뉴", width="large")
def admin_menu_dialog():
    input_pw_admin = st.text_input("🔒 관리자 비밀번호", type="password")
    if input_pw_admin == ADMIN_PASSWORD:
        st.subheader("🔮 AI 영업일 기준 예측")
        history_df = load_from_db()
        history_df['생산일'] = pd.to_datetime(history_df['production_date']).dt.date
        
        predict_data = []
        today = get_now_kst().date()
        for prod, group in history_df.groupby('product_name'):
            dates = sorted(group['생산일'].drop_duplicates().tolist())
            if len(dates) < 2 or (today - dates[-1]).days > 120: continue
            intervals = [np.busday_count(dates[i-1], dates[i], holidays=KR_HOLIDAYS) for i in range(1, len(dates))]
            avg = np.mean(intervals)
            next_d = np.busday_offset(dates[-1], int(avg), roll='forward', holidays=KR_HOLIDAYS)
            predict_data.append({"제품": prod, "다음 예상일": next_d})
        st.table(pd.DataFrame(predict_data))
        st.title("🎨 색도 관리 시스템")
if st.button("🛠️ 관리자 메뉴"): admin_menu_dialog()

# 입력 및 조회
tab1, tab2 = st.tabs(["📋 등록", "📊 조회"])
with tab1:
    prod = st.selectbox("제품", list(TARGET_DATA.keys()))
    meas = st.number_input("측정 색도", value=0.0)
    if st.button("등록"):
        save_to_db(get_now_kst().strftime("%Y-%m-%d"), "버닝", "윤승태", prod, 0.0, meas, 0.0, "합격", "", "-")
        st.rerun()

with tab2:
    history_df = load_from_db()
    page_size = 50
    page = st.number_input("페이지", 1, max(1, len(history_df)//page_size + 1))
    df = history_df.iloc[(page-1)*page_size : page*page_size].copy()
    
    # 스타일 적용 (HTML)
    def style_df(df):
        return df.style.map(lambda x: 'background-color: #ffcccc' if '불합격' in str(x) else '')
    st.markdown(style_df(df).to_html(), unsafe_allow_html=True)
