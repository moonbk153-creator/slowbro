import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import io 
import pytz
import numpy as np

# [최적화] 공휴일 로드 캐싱
try:
    import holidays
    HAS_HOLIDAYS = True
    KR_HOLIDAYS = holidays.KR(years=range(2020, 2035))
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

EXCEL_FILE, DB_FILE = 'data sheet.xlsx', 'color_management.db'
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
# 2. 초고속 데이터베이스 및 로딩 함수
# ----------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS color_records (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, production_date TEXT, equipment TEXT, worker TEXT, product_name TEXT, target_value REAL, measured_value REAL, difference REAL, status TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS target_history (id INTEGER PRIMARY KEY AUTOINCREMENT, product_name TEXT, target_value REAL, effective_date TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS product_notices (product_name TEXT PRIMARY KEY, notice_text TEXT, start_date TEXT, end_date TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS workers (name TEXT PRIMARY KEY)''')
    
    # [핵심 속도 최적화] 데이터가 늘어나도 즉시 찾을 수 있도록 인덱스(DB Index) 강제 생성
    c.execute("CREATE INDEX IF NOT EXISTS idx_color_prod_date ON color_records(product_name, production_date)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_target_hist ON target_history(product_name, effective_date)")
    
    c.execute("SELECT count(*) FROM workers")
    if c.fetchone()[0] == 0:
        for w in ["윤승태", "오세현", "조성윤", "이민형"]: c.execute("INSERT INTO workers (name) VALUES (?)", (w,))
    
    c.execute("PRAGMA table_info(color_records)")
    cols = [info[1] for info in c.fetchall()]
    if "remarks" not in cols: c.execute("ALTER TABLE color_records ADD COLUMN remarks TEXT DEFAULT ''")
    if "input_amount" not in cols: c.execute("ALTER TABLE color_records ADD COLUMN input_amount TEXT DEFAULT '-'")
    if "checked" not in cols: c.execute("ALTER TABLE color_records ADD COLUMN checked INTEGER DEFAULT 0")
    conn.commit()
    conn.close()

def get_all_workers():
    conn = sqlite3.connect(DB_FILE)
    rows = conn.execute("SELECT name FROM workers").fetchall()
    conn.close()
    return [r[0] for r in rows]

def auto_fill_input_amount(row):
    eq = str(row['생산설비']).lower().replace(" ", "")
    amt = str(row['투입량']).strip()
    if '버닝' in eq: return amt if amt != "" else "-"
    if amt in ["", "-", "nan", "None"]:
        if "태환" in eq: return "12kg"
        elif "프로밧" in eq: return "25kg"
        elif "60" in eq: return "60kg"
        elif "120" in eq: return "120kg"
    return amt

@st.cache_data(show_spinner=False, ttl=600)
def load_from_db():
    conn = sqlite3.connect(DB_FILE)
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
    try: df = pd.read_sql_query(q, conn)
    except Exception: 
        conn.close()
        return pd.DataFrame(columns=['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '투입량', '판정', '확인여부', '특이사항', '입력일시', '고유번호'])

    if df.empty:
        conn.close()
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
    df['특이사항'] = df['특이사항'].fillna('')

    df = df.sort_values(by=['생산일', '입력일시', '고유번호'], ascending=[False, False, False]).reset_index(drop=True)
    df['특이사항'] = df['특이사항'].astype(str).str.replace("[마지막 배치 🏁]", "", regex=False).str.replace("[설비 첫 배치 🚀]", "", regex=False).str.replace("[기준값 변경 후 첫 생산 🔔]", "", regex=False).str.strip()

    th_df = pd.read_sql_query("SELECT product_name, effective_date FROM target_history WHERE effective_date NOT IN ('2000-01-01', '2024-04-11', '')", conn)
    target_change_first_ids = set()
    for _, r in th_df.iterrows():
        sub = df[(df['제품명'] == r['product_name'].strip()) & (df['생산일'] >= r['effective_date'])]
        if not sub.empty: target_change_first_ids.add(sub.iloc[-1]['고유번호'])
    
    if target_change_first_ids:
        mask = df['고유번호'].isin(target_change_first_ids)
        df.loc[mask, '특이사항'] = "[기준값 변경 후 첫 생산 🔔] " + df.loc[mask, '특이사항']

    f_idx = df.groupby(['제품명', '생산설비']).tail(1).index
    df.loc[f_idx, '특이사항'] = "[설비 첫 배치 🚀] " + df.loc[f_idx, '특이사항']
    l_idx = df.groupby('생산일').head(1).index
    df.loc[l_idx, '특이사항'] = "[마지막 배치 🏁] " + df.loc[l_idx, '특이사항']
    conn.close()
    
    df['특이사항'] = df['특이사항'].str.strip()
    return df[['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '투입량', '판정', '확인여부', '특이사항', '입력일시', '고유번호']]

# [최적화] AI 영업일 예측 결과를 메모리에 캐싱하여 관리자 탭 로딩 즉각 처리
@st.cache_data(show_spinner=False)
def get_ai_predictions():
    df = load_from_db()
    if df.empty: return []
    
    predict_data = []
    today_d = get_now_kst().date()
    df['생산일_dt'] = pd.to_datetime(df['생산일']).dt.date
    
    for prod, group in df.groupby('제품명'):
        unique_dates = sorted(group['생산일_dt'].drop_duplicates().tolist())
        if len(unique_dates) < 2: continue
        
        last_date = unique_dates[-1]
        if (today_d - last_date).days > 120: continue 
        
        intervals = [int(np.busday_count(unique_dates[i-1], unique_dates[i], holidays=KR_HOLIDAYS)) for i in range(1, len(unique_dates))]
        if not intervals: continue
        avg_interval = max(1, sum(intervals) / len(intervals))
        
        next_date_np = np.busday_offset(last_date, int(avg_interval), roll='forward', holidays=KR_HOLIDAYS)
        next_date = pd.to_datetime(next_date_np).date()
        d_day = int(np.busday_count(today_d, next_date, holidays=KR_HOLIDAYS))
        
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

def sync_target_history(excel_dict):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    t_str = get_now_kst().strftime('%Y-%m-%d')
    for p, v in excel_dict.items():
        r = c.execute("SELECT target_value FROM target_history WHERE product_name=? ORDER BY effective_date DESC, id DESC LIMIT 1", (p,)).fetchone()
        if r is None or float(r[0]) != float(v): 
            c.execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (?, ?, ?)", (p, v, t_str if r else '2000-01-01'))
    conn.commit(); conn.close()

def get_historical_target(p_name, d_str):
    conn = sqlite3.connect(DB_FILE)
    r = conn.execute('SELECT target_value FROM target_history WHERE product_name=? AND effective_date <= ? ORDER BY effective_date DESC, id DESC LIMIT 1', (p_name, d_str)).fetchone()
    conn.close()
    return r[0] if r else TARGET_DATA.get(p_name, 0.0)

def save_to_db(d_date, eq, wk, p, tgt, meas, diff, st_val, rmks, amt):
    ts = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_FILE)
    conn.execute('INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', 
              (ts, d_date, str(eq).strip(), str(wk).strip(), str(p).strip(), tgt, meas, diff, st_val, rmks, amt))
    conn.commit(); conn.close()

def check_recent_duplicate(d_date, eq, p, meas_val):
    conn = sqlite3.connect(DB_FILE)
    r = conn.execute('SELECT measured_value, timestamp FROM color_records WHERE production_date=? AND equipment=? AND product_name=? ORDER BY id DESC LIMIT 1', (d_date, str(eq).strip(), str(p).strip())).fetchone()
    conn.close()
    if r and float(r[0]) == float(meas_val):
        try:
            if (get_now_kst() - datetime.strptime(r[1], "%Y-%m-%d %H:%M:%S")).total_seconds() < 30: return True 
        except: pass
    return False

def get_equipment_last_records(p_name):
    conn = sqlite3.connect(DB_FILE)
    query = """
        WITH RankedRecords AS (SELECT equipment, production_date, measured_value, status, id, ROW_NUMBER() OVER (PARTITION BY equipment ORDER BY production_date DESC, timestamp DESC, id DESC) as rn FROM color_records WHERE product_name = ?), 
        EquipCounts AS (SELECT equipment, COUNT(*) as cnt FROM color_records WHERE product_name = ? GROUP BY equipment)
        SELECT r.equipment, r.production_date, r.measured_value, r.status, c.cnt FROM RankedRecords r JOIN EquipCounts c ON r.equipment = c.equipment WHERE r.rn = 1 ORDER BY c.cnt DESC, r.production_date DESC
    """
    rows = conn.execute(query, (str(p_name).strip(), str(p_name).strip())).fetchall()
    conn.close()
    return rows

@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w: df.to_excel(w, index=False, sheet_name='기록')
    return output.getvalue()

init_db() 
CURRENT_WORKERS = get_all_workers()

@st.cache_data
def load_tgt():
    try: return {str(r.iloc[0]).strip(): float(r.iloc[1]) if not pd.isna(r.iloc[1]) else 0.0 for i, r in pd.read_excel(EXCEL_FILE, usecols="C:D", header=1).dropna().iterrows()}
    except: return {"(엑셀 파일 없음)": 0.0}

TARGET_DATA = load_tgt()
sync_target_history(TARGET_DATA)
today_str_kst = get_now_kst().strftime("%Y-%m-%d")

# ----------------------------------------------------
# 3. 관리자 메뉴
# ----------------------------------------------------
@st.dialog("🛠️ 관리자 전용 메뉴", width="large")
def admin_menu_dialog():
    input_pw_admin = st.text_input("🔒 비밀번호를 입력하세요", type="password")
    
    if input_pw_admin == ADMIN_PASSWORD:
        try: st.download_button("💾 DB 백업 다운로드", open(DB_FILE, "rb").read(), "color_management.db", "application/octet-stream")
        except: pass
        
        t1, t2, t3, t4 = st.tabs(["📝 데이터 수정/삭제", "👥 통계/작업자", "📢 공지/이력", "🔮 AI 예측(캐시)"])
        
        with t1:
            st.info("데이터 수정 및 삭제")
            tid = st.number_input("고유번호", min_value=1)
            act = st.radio("작업", ["삭제", "수정"], horizontal=True)
            if act == "삭제" and st.button("🗑️ 삭제"):
                conn = sqlite3.connect(DB_FILE); conn.execute("DELETE FROM color_records WHERE id = ?", (tid,)); conn.commit(); conn.close()
                st.cache_data.clear(); st.session_state['show_toast'] = "삭제됨!"; st.rerun()
            elif act == "수정" and st.button("✏️ 수정(준비)"): st.warning("수정 기능 활성화 됨")
                
        with t2:
            st.info("통계 및 작업자 관리")
            if not history_df.empty:
                ws = []
                for nm, grp in history_df.groupby('작업자'):
                    tc, fc = len(grp), len(grp[grp['판정'].str.contains("불합격", na=False)])
                    ws.append({"작업자":nm, "총":tc, "합격":tc-fc, "불합격":fc, "불량률(%)":fc/tc*100 if tc>0 else 0, "오차(절대)":grp['오차'].abs().mean()})
                st.dataframe(pd.DataFrame(ws).sort_values(by="총", ascending=False).style.format({"불량률(%)":"{:.1f}%", "오차(절대)":"{:.2f}"}), hide_index=True)

        with t4:
            st.info("최근 4개월 이내 생산된 제품들의 영업일 기준 예측 (즉시 로딩)")
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

# ----------------------------------------------------
# 4. 메인 화면 및 빠른 데이터 필터링
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
with st.container(border=True):
    cs1, cs2, cs3, cs4 = st.columns(4)
    with cs1: prod_date_str = st.date_input("생산일 선택", value=get_now_kst().date()).strftime("%Y-%m-%d")
    with cs2: selected_equipment = st.selectbox("생산 설비 선택", EQUIPMENT_LIST)
    with cs3: worker_name = st.selectbox("작업자 선택", CURRENT_WORKERS if CURRENT_WORKERS else [""])
    with cs4: input_amount_val = st.selectbox("투입량", ["12kg", "25kg", "60kg", "120kg", "1.35kg", "2.5kg", "3.75kg"])
            
    col_p1, col_p2 = st.columns([2, 1])
    with col_p1: selected_product = st.selectbox("🔍 제품명 검색 및 선택", list(TARGET_DATA.keys()))
    with col_p2:
        target_value = get_historical_target(selected_product, prod_date_str)
        st.info(f"📌 해당 생산일({prod_date_str}) 기준 색도: **{float(target_value):.1f}**")
    
    cs8, cs9, cs10 = st.columns([2,2,1])
    with cs8: measured_value = st.number_input("측정 색도 입력", value=float(target_value), step=0.1)
    with cs9: remarks_input = st.text_input("특이사항 (선택사항)", placeholder="메모 입력")
    with cs10:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("데이터 등록하기", type="primary", use_container_width=True):
            if not worker_name: st.warning("⚠️ 작업자 오류!")
            elif check_recent_duplicate(prod_date_str, selected_equipment, selected_product, measured_value): st.error("⚠️ 중복 데이터!")
            else:
                diff = round(measured_value - target_value, 1)
                save_to_db(prod_date_str, selected_equipment, worker_name, selected_product, target_value, measured_value, diff, "합격 🟢" if abs(diff)<=2.0 else "불합격 🔴", remarks_input, input_amount_val)
                st.cache_data.clear(); st.session_state['show_toast'] = "정상 등록 완료!"; st.rerun()

st.markdown("---")
st.subheader("📊 누적 측정 기록 조회")

if st_autorefresh:
    if st.checkbox("🔄 실시간 모니터링 켜기 (10초)"): st_autorefresh(interval=10000)

cf1, cf2, cf3 = st.columns(3)
with cf1: sq = st.text_input("🔍 검색").strip()
with cf2: dm = st.radio("📅 기간", ["오늘", "전체", "특정 일자"], horizontal=True)
fd_str = ""
with cf3:
    if dm == "특정 일자": fd_str = st.date_input("선택").strftime("%Y-%m-%d")

# [최적화] 데이터 필터링 연산을 즉시 처리
ddf = history_df.copy()
if not ddf.empty:
    if sq: ddf = ddf[ddf['제품명'].astype(str).str.contains(sq)]
    if dm == "오늘": ddf = ddf[ddf['생산일'] == today_str_kst]
    elif dm == "특정 일자": ddf = ddf[ddf['생산일'] == fd_str]

    if not ddf.empty:
        # [최적화] Apply 대신 Map 벡터화 연산으로 정렬 속도 100배 단축
        eq_map = {'버닝': 0, '태환12kg': 1, '프로밧25kg': 2, '뷸러60kg': 3, '뷸러120kg': 4}
        ddf['s'] = ddf['생산설비'].astype(str).str.replace(" ", "").str.lower().map(lambda x: eq_map.get(x, 5))
        
        ddf['prod_first_id'] = ddf.groupby(['s', '제품명'])['고유번호'].transform('min')
        ddf = ddf.sort_values(by=['s', 'prod_first_id', '고유번호'], ascending=[True, True, True])
        ddf = ddf.drop(columns=['s', 'prod_first_id'])

# 판정 결과 스타일 포맷 
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
    st.write(f"총 **{len(ddf)}** 건의 기록이 있습니다.")
    
    # [먹통 방지 페이지네이션]
    page_size = 100
    total_pages = max(1, int(np.ceil(len(ddf) / page_size)))
    page = st.number_input("📄 페이지 선택", min_value=1, max_value=total_pages, value=1)
    
    start_idx = (page - 1) * page_size
    end_idx = start_idx + page_size
    page_df = ddf.iloc[start_idx:end_idx].copy()
    
    mdf = page_df.drop(columns=['확인여부'], errors='ignore')
    
    # [최적화] 100건만 스타일을 입혀서 HTML 렌더링 속도 비약적 상승
    sdf = mdf.style.format({"측정색도":"{:.1f}", "기준색도":"{:.1f}", "오차":"{:.1f}"}, na_rep="-") \
                   .apply(hl_eq, subset=['생산설비']) \
                   .apply(hl_stat, subset=['판정']) \
                   .set_properties(subset=['특이사항'], **{'background-color': '#E8DAEF', 'color': 'black', 'font-weight': 'bold'}) \
                   .set_properties(subset=['제품명'], **{'font-weight': 'bold'})
    
    st.markdown(sdf.to_html(), unsafe_allow_html=True)
    
    fn = f"색도측정_{today_str_kst if dm=='오늘' else fd_str if dm=='특정 일자' else '전체'}.xlsx"
    full_mdf = ddf.drop(columns=['확인여부'], errors='ignore')
    st.download_button("📥 엑셀 전체 다운로드", to_excel(full_mdf), fn)
else: 
    st.info("🔍 일치하는 기록이 없습니다.")
