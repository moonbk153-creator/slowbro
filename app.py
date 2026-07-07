import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import io 
import pytz

# 자동 새로고침 라이브러리 로드 (안전장치 포함)
try:
    from streamlit_autorefresh import st_autorefresh
except ImportError:
    st_autorefresh = None

# 한국 시간대 설정
KST = pytz.timezone('Asia/Seoul')

st.set_page_config(page_title="색도 관리 시스템", layout="wide")

# [UI 개선 1] Toast 알림을 화면이 새로고침된 후에도 띄우기 위한 세션 변수 처리
if 'show_toast' in st.session_state:
    st.toast(st.session_state['show_toast'], icon="✅")
    del st.session_state['show_toast']

EXCEL_FILE = 'data sheet.xlsx'
DB_FILE = 'color_management.db'

# 모든 설비 단위를 소문자 kg으로 통일
EQUIPMENT_LIST = ["버닝", "태환 12kg", "프로밧 25kg", "뷸러 60kg", "뷸러 120kg"]

# [보안] Streamlit Secrets에서 암호를 불러옵니다.
ADMIN_PASSWORD = st.secrets["ADMIN_PASSWORD"]
ACCESS_PASSWORD = st.secrets["APP_PASSWORD"]

# ==========================================
# [보안] 직원용 공용 접속 비밀번호 설정 및 자동 접속
# ==========================================
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    if "pw" in st.query_params and st.query_params["pw"] == ACCESS_PASSWORD:
        st.session_state['logged_in'] = True
        st.rerun()

    st.title("🔒 색도 관리 시스템 - 접속 제한")
    st.markdown("---")
    st.subheader("작업자 전용 인증")
    
    input_pw = st.text_input("사내 공용 비밀번호를 입력하세요", type="password", placeholder="비밀번호 8자리 입력")
    
    if st.button("🔓 시스템 접속하기"):
        if input_pw == ACCESS_PASSWORD:
            st.session_state['logged_in'] = True
            st.session_state['show_toast'] = "시스템 접속 성공!"
            st.rerun()
        else:
            st.error("❌ 비밀번호가 일치하지 않습니다. 다시 확인해 주세요.")
            
    st.info("💡 본 시스템은 외부인의 접근이 제한된 사내 품질 관리 프로그램입니다. 비밀번호 분실 시 관리자에게 문의하세요.")
    st.stop()

# ==========================================
# 1. 데이터베이스(DB) 함수 모음
# ==========================================
def get_now_kst():
    return datetime.now(KST)

def safe_date_parse(val):
    val_str = str(val).strip()
    if val_str in ['nan', 'None', '', 'NaN', 'NaT']: return ""
    val_str = val_str.split(" ")[0].replace("/", "-").replace(".", "-")
    try:
        return pd.to_datetime(val_str).strftime("%Y-%m-%d")
    except:
        return val_str

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS color_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            production_date TEXT,
            equipment TEXT,
            worker TEXT,
            product_name TEXT,
            target_value REAL,
            measured_value REAL,
            difference REAL,
            status TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS target_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_name TEXT,
            target_value REAL,
            effective_date TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS product_notices (
            product_name TEXT PRIMARY KEY,
            notice_text TEXT,
            start_date TEXT,
            end_date TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS workers (
            name TEXT PRIMARY KEY
        )
    ''')
    
    cursor.execute("SELECT count(*) FROM workers")
    if cursor.fetchone()[0] == 0:
        default_workers = ["윤승태", "오세현", "조성윤", "이민형"]
        for w in default_workers:
            cursor.execute("INSERT INTO workers (name) VALUES (?)", (w,))

    cursor.execute("PRAGMA table_info(color_records)")
    columns = [info[1] for info in cursor.fetchall()]

    if "remarks" not in columns:
        cursor.execute("ALTER TABLE color_records ADD COLUMN remarks TEXT DEFAULT ''")
    if "input_amount" not in columns:
        cursor.execute("ALTER TABLE color_records ADD COLUMN input_amount TEXT DEFAULT '-'")
    if "checked" not in columns:
        cursor.execute("ALTER TABLE color_records ADD COLUMN checked INTEGER DEFAULT 0")

    conn.commit()
    conn.close()

def get_all_workers():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM workers")
    rows = cursor.fetchall()
    conn.close()
    return [row[0] for row in rows]

def add_worker(name):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO workers (name) VALUES (?)", (name.strip(),))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    conn.close()
    return success

def delete_worker(name):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM workers WHERE name = ?", (name,))
    conn.commit()
    conn.close()

def update_checked_status(record_ids, status_val):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    for r_id in record_ids:
        cursor.execute("UPDATE color_records SET checked=? WHERE id=?", (status_val, r_id))
    conn.commit()
    conn.close()

def sync_target_history(excel_target_dict):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    today_str = get_now_kst().strftime('%Y-%m-%d')
    
    for prod, val in excel_target_dict.items():
        cursor.execute("SELECT target_value FROM target_history WHERE product_name=? ORDER BY effective_date DESC, id DESC LIMIT 1", (prod,))
        row = cursor.fetchone()
        
        if row is None:
            cursor.execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (?, ?, ?)", (prod, val, '2000-01-01'))
        elif float(row[0]) != float(val):
            cursor.execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (?, ?, ?)", (prod, val, today_str))
            
    conn.commit()
    conn.close()

def get_historical_target(product_name, prod_date_str):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT target_value FROM target_history 
        WHERE product_name=? AND effective_date <= ?
        ORDER BY effective_date DESC, id DESC LIMIT 1
    ''', (product_name, prod_date_str))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else TARGET_DATA.get(product_name, 0.0)

def get_target_history_df(product_name):
    conn = sqlite3.connect(DB_FILE)
    query = '''
        SELECT effective_date, target_value 
        FROM target_history 
        WHERE product_name = ? 
        ORDER BY effective_date DESC, id DESC
    '''
    df = pd.read_sql_query(query, conn, params=(product_name,))
    conn.close()
    return df

def save_notice(product_name, notice_text, start_date, end_date):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT OR REPLACE INTO product_notices (product_name, notice_text, start_date, end_date)
        VALUES (?, ?, ?, ?)
    ''', (product_name, notice_text, start_date, end_date))
    conn.commit()
    conn.close()

def delete_notice(product_name):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM product_notices WHERE product_name = ?", (product_name,))
    conn.commit()
    conn.close()

def get_all_active_notices(today_str):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT product_name, notice_text 
        FROM product_notices 
        WHERE start_date <= ? AND end_date >= ?
    ''', (today_str, today_str))
    rows = cursor.fetchall()
    conn.close()
    return {row[0]: row[1] for row in rows}

def get_raw_notice(product_name):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT notice_text, start_date, end_date FROM product_notices WHERE product_name=?", (product_name,))
    row = cursor.fetchone()
    conn.close()
    return row

def save_to_db(prod_date, equipment, worker, product, target, measured, diff, status, remarks, input_amount="-"):
    timestamp = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (timestamp, prod_date, str(equipment).strip(), str(worker).strip(), str(product).strip(), target, measured, diff, status, remarks, input_amount))
    conn.commit()
    conn.close()

def check_recent_duplicate(prod_date, equipment, product, measured_val):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT measured_value, timestamp 
        FROM color_records 
        WHERE production_date=? AND equipment=? AND product_name=?
        ORDER BY id DESC LIMIT 1
    ''', (prod_date, str(equipment).strip(), str(product).strip()))
    row = cursor.fetchone()
    conn.close()
    
    if row:
        last_val, last_time_str = row
        if float(last_val) == float(measured_val):
            try:
                last_time = datetime.strptime(last_time_str, "%Y-%m-%d %H:%M:%S")
                current_time_str = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
                current_time = datetime.strptime(current_time_str, "%Y-%m-%d %H:%M:%S")
                if (current_time - last_time).total_seconds() < 30: 
                    return True 
            except:
                pass
    return False

def auto_fill_input_amount(row):
    equip = str(row['생산설비']).lower().replace(" ", "")
    current_amt = str(row['투입량']).strip()
    if '버닝' in equip: return current_amt if current_amt != "" else "-"
    if current_amt in ["", "-", "nan", "None"]:
        if "태환" in equip: return "12kg"
        elif "프로밧" in equip: return "25kg"
        elif "60" in equip: return "60kg"
        elif "120" in equip: return "120kg"
    return current_amt

@st.cache_data(show_spinner=False)
def load_from_db():
    conn = sqlite3.connect(DB_FILE)
    query = """
        SELECT 
            c.id as 고유번호, 
            c.timestamp as 입력일시, 
            c.production_date as 생산일, 
            c.equipment as 생산설비, 
            COALESCE(c.input_amount, '-') as 투입량,
            c.worker as 작업자, 
            c.product_name as 제품명, 
            c.measured_value as 측정색도, 
            COALESCE(c.remarks, '') as 특이사항,
            COALESCE(c.checked, 0) as checked_status,
            COALESCE(
                (SELECT target_value FROM target_history th 
                 WHERE th.product_name = c.product_name AND th.effective_date <= c.production_date 
                 ORDER BY th.effective_date DESC LIMIT 1),
                (SELECT target_value FROM target_history th 
                 WHERE th.product_name = c.product_name 
                 ORDER BY th.effective_date ASC LIMIT 1),
                0.0
            ) as 기준색도
        FROM color_records c
    """
    try:
        df = pd.read_sql_query(query, conn)
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

    df['특이사항'] = df['특이사항'].astype(str).str.replace("[마지막 배치 🏁]", "", regex=False)\
                                                 .str.replace("[설비 첫 배치 🚀]", "", regex=False)\
                                                 .str.replace("[기준값 변경 후 첫 생산 🔔]", "", regex=False).str.strip()

    th_df = pd.read_sql_query("SELECT product_name, effective_date FROM target_history WHERE effective_date NOT IN ('2000-01-01', '2024-04-11', '')", conn)
    target_change_first_ids = set()
    for _, row in th_df.iterrows():
        subset = df[(df['제품명'] == row['product_name'].strip()) & (df['생산일'] >= row['effective_date'])]
        if not subset.empty:
            target_change_first_ids.add(subset.iloc[-1]['고유번호'])
    
    if target_change_first_ids:
        tc_mask = df['고유번호'].isin(target_change_first_ids)
        df.loc[tc_mask, '특이사항'] = "[기준값 변경 후 첫 생산 🔔] " + df.loc[tc_mask, '특이사항']

    first_batch_indices = df.groupby(['제품명', '생산설비']).tail(1).index
    df.loc[first_batch_indices, '특이사항'] = "[설비 첫 배치 🚀] " + df.loc[first_batch_indices, '특이사항']

    last_batch_indices = df.groupby('생산일').head(1).index
    df.loc[last_batch_indices, '특이사항'] = "[마지막 배치 🏁] " + df.loc[last_batch_indices, '특이사항']

    conn.close()
    
    df['특이사항'] = df['특이사항'].str.strip()
    desired_order = ['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '투입량', '판정', '확인여부', '특이사항', '입력일시', '고유번호']
    return df[desired_order]

def delete_from_db(record_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM color_records WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()

def update_db(record_id, new_date, new_equip, new_worker, new_product, new_target, new_measured, new_diff, new_status, new_remarks, new_input_amount, new_checked=0):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE color_records 
        SET production_date=?, equipment=?, worker=?, product_name=?, target_value=?, measured_value=?, difference=?, status=?, remarks=?, input_amount=?, checked=?
        WHERE id=?
    ''', (new_date, str(new_equip).strip(), str(new_worker).strip(), str(new_product).strip(), new_target, new_measured, new_diff, new_status, new_remarks, new_input_amount, new_checked, record_id))
    conn.commit()
    conn.close()

def get_last_record(product_name):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT production_date, measured_value, status 
        FROM color_records 
        WHERE product_name = ? 
        ORDER BY production_date DESC, timestamp DESC LIMIT 1
    ''', (str(product_name).strip(),))
    row = cursor.fetchone()
    conn.close()
    return row

def get_equipment_last_records(product_name):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    query = """
        WITH RankedRecords AS (
            SELECT 
                equipment, production_date, measured_value, status, id,
                ROW_NUMBER() OVER (PARTITION BY equipment ORDER BY production_date DESC, timestamp DESC, id DESC) as rn
            FROM color_records
            WHERE product_name = ?
        ),
        EquipCounts AS (
            SELECT equipment, COUNT(*) as cnt
            FROM color_records
            WHERE product_name = ?
            GROUP BY equipment
        )
        SELECT r.equipment, r.production_date, r.measured_value, r.status, c.cnt
        FROM RankedRecords r
        JOIN EquipCounts c ON r.equipment = c.equipment
        WHERE r.rn = 1
        ORDER BY c.cnt DESC, r.production_date DESC
    """
    cursor.execute(query, (str(product_name).strip(), str(product_name).strip()))
    rows = cursor.fetchall()
    conn.close()
    return rows

@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='색도측정기록')
    return output.getvalue()

# ==========================================
# 2. 엑셀 기준표 로드 및 초기화
# ==========================================
init_db() 
CURRENT_WORKERS = get_all_workers()

@st.cache_data
def load_target_data():
    try:
        df = pd.read_excel(EXCEL_FILE, usecols="C:D", header=1)
        df = df.dropna(subset=[df.columns[0]]) 
        
        target_dict = {}
        for index, row in df.iterrows():
            prod = str(row.iloc[0]).strip()
            target_dict[prod] = float(row.iloc[1]) if not pd.isna(row.iloc[1]) else 0.0
                
        return target_dict
    except FileNotFoundError:
        st.error(f"⚠️ '{EXCEL_FILE}' 파일을 찾을 수 없습니다.")
        return {"(엑셀 파일 없음)": 0.0}

TARGET_DATA = load_target_data()
sync_target_history(TARGET_DATA)

today_str_kst = get_now_kst().strftime("%Y-%m-%d")
ACTIVE_NOTICES = get_all_active_notices(today_str_kst)

# ==========================================
# [UI 개선 5] 모달 다이얼로그(st.dialog)를 활용한 관리자 메뉴
# ==========================================
@st.dialog("🛠️ 관리자 전용 메뉴", width="large")
def admin_menu_dialog():
    input_password = st.text_input("🔒 관리자 비밀번호를 입력하세요", type="password", key="admin_pw_input")
    
    if input_password == ADMIN_PASSWORD:
        st.success("✅ 관리자 인증 완료")
        
        try:
            with open(DB_FILE, "rb") as f:
                db_bytes = f.read()
            st.download_button(
                label="💾 데이터베이스 원본 파일(.db) 백업 다운로드",
                data=db_bytes,
                file_name="color_management.db",
                mime="application/octet-stream",
                type="secondary"
            )
        except Exception:
            pass
        
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "🔍 금일 확인", "📝 수정/삭제", "📂 엑셀 업로드", "📅 기준값 이력", 
            "📢 공지 관리", "⏳ 장기 미생산", "👥 통계", "🧑‍🔧 명단 관리"
        ])
        
        with tab1:
            st.info("오늘 생산된 배치 데이터를 모니터링하여 관리자가 최종 확인한 건과 미확인 내역을 분리 관리합니다.")
            today_batches_df = history_df[history_df['생산일'].astype(str).str.contains(today_str_kst, na=False)]
            
            if today_batches_df.empty:
                st.info("📅 금일 등록된 생산 기록이 없습니다.")
            else:
                view_mode = st.radio("분리 조회", ["미확인 ❌", "확인완료 ✅"], horizontal=True, key="admin_view_mode")
                target_df = today_batches_df[today_batches_df['확인여부'] == view_mode]
                
                if target_df.empty:
                    st.success(f"해당 상태({view_mode})의 내역이 없습니다.")
                else:
                    st.dataframe(target_df, use_container_width=True, hide_index=True)
                    
                    if view_mode == "미확인 ❌":
                        to_confirm = st.multiselect("✅ 확인 완료 처리할 배치 선택", target_df['고유번호'].tolist())
                        if st.button("선택 건 확인 완료", type="primary", use_container_width=True):
                            if to_confirm:
                                update_checked_status(to_confirm, 1)
                                st.cache_data.clear()
                                st.session_state['show_toast'] = f"{len(to_confirm)}건 확인 완료 처리됨!"
                                st.rerun()
                    else:
                        to_unconfirm = st.multiselect("🔄 미확인으로 되돌릴 배치 선택", target_df['고유번호'].tolist())
                        if st.button("선택 건 미확인 복구", use_container_width=True):
                            if to_unconfirm:
                                update_checked_status(to_unconfirm, 0)
                                st.cache_data.clear()
                                st.session_state['show_toast'] = f"{len(to_unconfirm)}건 미확인으로 되돌림!"
                                st.rerun()

        with tab2:
            st.write("표의 **'고유번호'**를 확인한 후 작업을 진행하세요.")
            col_admin1, col_admin2 = st.columns(2)
            with col_admin1:
                target_id = st.number_input("대상 고유번호 입력", min_value=1, step=1)
                action = st.radio("작업 선택", ["데이터 삭제", "데이터 수정"])
                
            with col_admin2:
                if action == "데이터 삭제":
                    st.warning("삭제된 데이터는 복구할 수 없습니다.")
                    if st.button("🗑️ 선택한 데이터 삭제"):
                        delete_from_db(target_id)
                        st.cache_data.clear()
                        st.session_state['show_toast'] = "데이터 삭제됨!"
                        st.rerun()
                        
                elif action == "데이터 수정":
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute("PRAGMA table_info(color_records)")
                    columns_info = [info[1] for info in c.fetchall()]
                    chk_col = "checked" if "checked" in columns_info else "0"
                    
                    c.execute(f"SELECT product_name, target_value, production_date, equipment, worker, measured_value, remarks, input_amount, COALESCE({chk_col}, 0) FROM color_records WHERE id=?", (target_id,))
                    row = c.fetchone()
                    conn.close()
                    
                    if row:
                        p_name, t_val, p_date, equip, workr, m_val, rmks, i_amt, chk_stat = row
                        
                        try: curr_date = datetime.strptime(p_date, "%Y-%m-%d").date()
                        except: curr_date = get_now_kst().date()
                            
                        new_p_date = st.date_input("수정할 생산일 지정", value=curr_date)
                        new_p_date_str = new_p_date.strftime("%Y-%m-%d")
                        
                        prod_options = list(TARGET_DATA.keys())
                        if p_name in prod_options: p_idx = prod_options.index(p_name)
                        else:
                            prod_options = [p_name] + prod_options
                            p_idx = 0
                        new_p_name = st.selectbox("수정할 제품명", prod_options, index=p_idx)
                        
                        actual_t_val = get_historical_target(new_p_name, new_p_date_str)
                        st.info(f"📌 지정 날짜 기준색도: **{float(actual_t_val):.1f}**")
                        
                        try: equip_index = EQUIPMENT_LIST.index(equip)
                        except: equip_index = 0
                            
                        new_equip = st.selectbox("수정할 설비", EQUIPMENT_LIST, index=equip_index)
                        
                        new_equip_clean = str(new_equip).lower().replace(" ", "")
                        if "버닝" in new_equip_clean:
                            try: amt_idx = ["1.35kg", "2.5kg", "3.75kg"].index(i_amt)
                            except: amt_idx = 0
                            new_amt = st.selectbox("수정할 투입량", ["1.35kg", "2.5kg", "3.75kg"], index=amt_idx)
                        else:
                            if "태환" in new_equip_clean: new_amt = "12kg"
                            elif "프로밧" in new_equip_clean: new_amt = "25kg"
                            elif "60" in new_equip_clean: new_amt = "60kg"
                            elif "120" in new_equip_clean: new_amt = "120kg"
                            else: new_amt = "-"
                            st.text_input("수정할 투입량 (자동고정)", value=new_amt, disabled=True)

                        if workr in CURRENT_WORKERS: w_options, w_idx = CURRENT_WORKERS, CURRENT_WORKERS.index(workr)
                        else: w_options, w_idx = CURRENT_WORKERS + [workr], len(CURRENT_WORKERS)
                        new_worker = st.selectbox("수정할 작업자", w_options, index=w_idx)

                        new_m_val = st.number_input("수정할 측정색도", value=float(m_val), step=0.1, format="%.1f")
                        new_rmks = st.text_input("수정할 특이사항", value=rmks if rmks else "")
                        
                        if st.button("✏️ 선택한 데이터 수정"):
                            new_diff = round(new_m_val - actual_t_val, 1)
                            new_status = "합격 🟢" if abs(new_diff) <= 2.0 else "불합격 🔴"
                            update_db(target_id, new_p_date_str, new_equip, new_worker, new_p_name, actual_t_val, new_m_val, new_diff, new_status, new_rmks, new_amt, chk_stat)
                            st.cache_data.clear()
                            st.session_state['show_toast'] = "데이터가 수정되었습니다!"
                            st.rerun()
                    else:
                        st.error("해당 고유번호가 없습니다.")
            
            st.markdown("---")
            if st.button("🧹 DB 전체 텍스트 공백 영구 정화"):
                conn = sqlite3.connect(DB_FILE)
                cursor = conn.cursor()
                cursor.execute("UPDATE color_records SET worker = TRIM(worker), equipment = TRIM(equipment), product_name = TRIM(product_name)")
                conn.commit()
                conn.close()
                st.cache_data.clear()
                st.session_state['show_toast'] = "DB 정화 완료!"
                st.rerun()
        
        with tab3:
            uploaded_file = st.file_uploader("과거 측정 기록 엑셀 파일 선택", type=['xlsx', 'xls'], key="record_upload")
            if uploaded_file is not None and st.button("🚀 업로드 실행"):
                try:
                    df_upload = pd.read_excel(uploaded_file)
                    required_cols = ['생산일', '제품명', '생산설비', '작업자', '측정색도']
                    if all(col in df_upload.columns for col in required_cols):
                        for index, row in df_upload.iterrows():
                            measured_str = str(row['측정색도']).strip()
                            if measured_str in ['-', '', 'nan', 'None']: continue
                            try: measured = float(measured_str)
                            except ValueError: continue
                            prod_date_str = safe_date_parse(row['생산일'])
                            product, equip, worker = str(row['제품명']).strip(), str(row['생산설비']).strip(), str(row['작업자']).strip()
                            e_clean = equip.lower().replace(" ", "")
                            if '버닝' in e_clean: upload_amt = str(row['투입량']).strip() if '투입량' in df_upload.columns and not pd.isna(row.get('투입량')) else "1.35kg"
                            elif '태환' in e_clean: upload_amt = "12kg"
                            elif '프로밧' in e_clean: upload_amt = "25kg"
                            elif '60' in e_clean: upload_amt = "60kg"
                            elif '120' in e_clean: upload_amt = "120kg"
                            else: upload_amt = "-"
                            upload_remarks = str(row['특이사항']).strip() if '특이사항' in df_upload.columns and not pd.isna(row.get('특이사항')) else ""
                            if upload_remarks in ['nan', 'None']: upload_remarks = ""
                            target = get_historical_target(product, prod_date_str)
                            diff = round(measured - target, 1)
                            status = "합격 🟢" if abs(diff) <= 2.0 else "불합격 🔴"
                            current_time = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
                            
                            conn = sqlite3.connect(DB_FILE)
                            c = conn.cursor()
                            c.execute('INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', 
                                      (current_time, prod_date_str, equip, worker, product, target, measured, diff, status, upload_remarks, upload_amt))
                            conn.commit()
                            conn.close()
                        st.cache_data.clear()
                        st.session_state['show_toast'] = "엑셀 업로드 성공!"
                        st.rerun()
                    else: st.error("❌ 엑셀 파일에 필수 열이 부족합니다.")
                except Exception as e: st.error(f"오류: {e}")

        with tab4:
            history_file = st.file_uploader("기준값 이력 엑셀 파일 선택", type=['xlsx', 'xls'], key="history_upload")
            if history_file is not None and st.button("🚀 기준값 이력 반영"):
                try:
                    df_history = pd.read_excel(history_file)
                    req_hist_cols = ['제품명', '적용시작일', '기준색도']
                    if all(col in df_history.columns for col in req_hist_cols):
                        conn = sqlite3.connect(DB_FILE)
                        c = conn.cursor()
                        c.execute("DELETE FROM target_history")
                        for index, row in df_history.iterrows():
                            product = str(row['제품명']).strip()
                            eff_date_str = safe_date_parse(row['적용시작일']) or '2000-01-01'
                            try:
                                target_val = float(row['기준색도'])
                                c.execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (?, ?, ?)", (product, target_val, eff_date_str))
                            except ValueError: continue 
                        conn.commit()
                        conn.close()
                        st.cache_data.clear()
                        st.session_state['show_toast'] = "기준값 이력 세팅 완료!"
                        st.rerun()
                    else: st.error("❌ 필수 열 부족")
                except Exception as e: st.error(f"오류: {e}")

        with tab5:
            notice_prod = st.selectbox("제품 선택", list(TARGET_DATA.keys()), key="notice_prod")
            raw_notice = get_raw_notice(notice_prod)
            curr_text = raw_notice[0] if raw_notice else ""
            curr_start = raw_notice[1] if raw_notice else get_now_kst().strftime("%Y-%m-%d")
            curr_end = raw_notice[2] if raw_notice else "2099-12-31"
            is_unlimited = (curr_end == "2099-12-31")
            
            try: start_dt = datetime.strptime(curr_start, "%Y-%m-%d").date()
            except: start_dt = get_now_kst().date()
            try: end_dt = datetime.strptime(curr_end, "%Y-%m-%d").date() if not is_unlimited else get_now_kst().date() + timedelta(days=7)
            except: end_dt = get_now_kst().date() + timedelta(days=7)

            notice_text = st.text_area("공지 내용", value=curr_text)
            col_n1, col_n2 = st.columns(2)
            with col_n1: start_date = st.date_input("시작일", value=start_dt)
            with col_n2:
                no_limit = st.checkbox("무기한", value=is_unlimited)
                end_date = st.date_input("종료일", value=end_dt, disabled=no_limit)
                
            if st.button("📢 공지 등록/수정", type="primary"):
                end_str = "2099-12-31" if no_limit else end_date.strftime("%Y-%m-%d")
                save_notice(notice_prod, notice_text, start_date.strftime("%Y-%m-%d"), end_str)
                st.cache_data.clear()
                st.session_state['show_toast'] = "공지사항 등록 완료!"
                st.rerun()
            if st.button("🗑️ 공지 삭제"):
                delete_notice(notice_prod)
                st.cache_data.clear()
                st.session_state['show_toast'] = "공지사항 삭제 완료!"
                st.rerun()
                    
        with tab6:
            today_date = get_now_kst().date()
            inactive_products = []
            for prod in TARGET_DATA.keys():
                last_record = get_last_record(prod)
                if last_record:
                    last_date_str = last_record[0]
                    try:
                        elapsed_days = (today_date - datetime.strptime(last_date_str, "%Y-%m-%d").date()).days
                        if elapsed_days >= 120:
                            inactive_products.append({"제품명": prod, "최종 생산일": last_date_str, "경과일": f"{elapsed_days}일"})
                    except: pass
                else: inactive_products.append({"제품명": prod, "최종 생산일": "기록 없음", "경과일": "이력 없음"})
            
            if inactive_products:
                st.dataframe(pd.DataFrame(inactive_products), use_container_width=True, hide_index=True)
            else: st.success("🎉 장기 방치된 미생산 제품이 없습니다!")
                
        with tab7:
            if not history_df.empty:
                worker_stats = []
                for w_name, group in history_df.groupby('작업자'):
                    total_cnt, fail_cnt = len(group), len(group[group['판정'].str.contains("불합격", na=False)])
                    worker_stats.append({
                        "작업자": w_name, "총 생산(배치)": total_cnt, "합격 🟢": total_cnt - fail_cnt,
                        "불합격 🔴": fail_cnt, "불량률(%)": (fail_cnt / total_cnt) * 100 if total_cnt > 0 else 0,
                        "평균 오차(절대값)": group['오차'].abs().mean()
                    })
                stats_df = pd.DataFrame(worker_stats).sort_values(by="총 생산(배치)", ascending=False)
                st.dataframe(stats_df.style.format({"불량률(%)": "{:.1f}%", "평균 오차(절대값)": "{:.2f}"}), use_container_width=True, hide_index=True)
            else: st.info("기록이 없습니다.")
                
        with tab8:
            st.write(f"**현재 등록된 작업자:** {', '.join(CURRENT_WORKERS)}")
            new_w = st.text_input("추가할 작업자 이름 입력")
            if st.button("작업자 추가", type="primary"):
                if add_worker(new_w):
                    st.cache_data.clear()
                    st.session_state['show_toast'] = "작업자 추가 완료!"
                    st.rerun()
            if CURRENT_WORKERS:
                del_w = st.selectbox("삭제할 작업자", CURRENT_WORKERS)
                if st.button("선택한 작업자 삭제"):
                    delete_worker(del_w)
                    st.cache_data.clear()
                    st.session_state['show_toast'] = "작업자 삭제 완료!"
                    st.rerun()
    elif input_password != "":
        st.error("비밀번호가 일치하지 않습니다.")

# ==========================================
# 3. 메인 화면 구성
# ==========================================
history_df = load_from_db() # 캐싱된 데이터 미리 로드

col_title, col_admin, col_logout = st.columns([7, 1.5, 1])
with col_title:
    st.title("🎨 일일 제품 색도 관리 시스템")
with col_admin:
    if st.button("🛠️ 관리자 메뉴", use_container_width=True):
        admin_menu_dialog()
with col_logout:
    if st.button("🔒 로그아웃", use_container_width=True):
        st.query_params.clear()
        st.session_state['logged_in'] = False
        st.rerun()
st.markdown("---")

# ==========================================
# [UI 개선 4] 입력 동선을 메인 화면 최상단으로 통합 및 분리
# ==========================================
st.subheader("📝 데이터 등록")
tab_std, tab_quick = st.tabs(["📋 일반 데이터 등록", "⚡ 진행 중인 라인 빠른 추가"])

with tab_std:
    with st.container(border=True):
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            prod_date_input = st.date_input("생산일 선택", value=get_now_kst().date())
            prod_date_str = prod_date_input.strftime("%Y-%m-%d")
        with col_s2:
            selected_equipment = st.selectbox("생산 설비 선택", EQUIPMENT_LIST)
            equip_clean = str(selected_equipment).lower().replace(" ", "")
        with col_s3:
            if not CURRENT_WORKERS: st.error("작업자 없음")
            worker_name = st.selectbox("작업자 선택", CURRENT_WORKERS if CURRENT_WORKERS else [""])
        with col_s4:
            if "버닝" in equip_clean:
                input_amount_val = st.selectbox("투입량 선택", ["1.35kg", "2.5kg", "3.75kg"])
            else:
                input_amount_val = "12kg" if "태환" in equip_clean else "25kg" if "프로밧" in equip_clean else "60kg" if "60" in equip_clean else "120kg" if "120" in equip_clean else "-"
                st.text_input("투입량 (고정)", value=input_amount_val, disabled=True)
        
        st.markdown("---")
        col_s5, col_s6, col_s7 = st.columns([2, 1, 1])
        with col_s5:
            selected_product = st.selectbox("🔍 제품명 검색 및 선택", list(TARGET_DATA.keys()))
        with col_s6:
            target_value = get_historical_target(selected_product, prod_date_str)
            st.text_input("기준 색도", value=f"{float(target_value):.1f}", disabled=True)
        with col_s7:
            measured_value = st.number_input("측정 색도 입력", value=float(target_value), step=0.1, format="%.1f")
        
        if ACTIVE_NOTICES.get(selected_product):
            st.warning(f"📢 **[작업자 전달사항]** {ACTIVE_NOTICES[selected_product]}", icon="🚨")
            
        col_s8, col_s9 = st.columns([3, 1])
        with col_s8:
            remarks_input = st.text_input("특이사항 (선택사항)", placeholder="간단한 메모 입력")
        with col_s9:
            st.markdown("<br>", unsafe_allow_html=True) # 줄맞춤
            if st.button("데이터 등록하기", type="primary", use_container_width=True):
                if not worker_name:
                    st.warning("⚠️ 작업자 이름을 지정해 주세요!")
                elif check_recent_duplicate(prod_date_str, selected_equipment, selected_product, measured_value):
                    st.error("⚠️ 방금 동일한 측정값이 등록되었습니다. 중복 방지 대기 중.")
                else:
                    difference = round(measured_value - target_value, 1)
                    status = "합격 🟢" if abs(difference) <= 2.0 else "불합격 🔴"
                    save_to_db(prod_date_str, selected_equipment, worker_name, selected_product, target_value, measured_value, difference, status, remarks_input, input_amount_val)
                    st.cache_data.clear()
                    st.session_state['show_toast'] = f"{selected_product} 데이터가 정상적으로 등록되었습니다!"
                    st.rerun()

with tab_quick:
    with st.container(border=True):
        today_str_kst = get_now_kst().strftime("%Y-%m-%d")
        today_batches_df = history_df[history_df['생산일'] == today_str_kst]
        
        if today_batches_df.empty:
            st.info("💡 오늘 아직 생산된 기록이 없습니다. '일반 데이터 등록' 탭을 먼저 사용해 주세요.")
        else:
            recent_batches = today_batches_df[['제품명', '생산설비', '투입량', '작업자']].drop_duplicates().reset_index(drop=True)
            batch_options = [f"▶ {row['제품명']} (설비: {row['생산설비']} / 투입량: {row['투입량']} / 작업자: {row['작업자']})" for idx, row in recent_batches.iterrows()]
            
            col_q1, col_q2, col_q3, col_q4 = st.columns([3, 1, 1, 1])
            with col_q1:
                selected_batch_str = st.selectbox("이어서 측정할 제품 선택", batch_options, key="quick_batch_select")
            
            if selected_batch_str:
                selected_idx = batch_options.index(selected_batch_str)
                quick_prod = recent_batches.iloc[selected_idx]['제품명']
                quick_equip = recent_batches.iloc[selected_idx]['생산설비']
                quick_amt = recent_batches.iloc[selected_idx]['투입량']
                quick_worker = recent_batches.iloc[selected_idx]['작업자']
                quick_target = get_historical_target(quick_prod, today_str_kst)
                
                with col_q2:
                    st.text_input("기준 색도", value=f"{float(quick_target):.1f}", disabled=True, key="quick_tgt")
                with col_q3:
                    quick_measured = st.number_input("새 측정값", value=float(quick_target), step=0.1, format="%.1f", key="quick_val")
                with col_q4:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🚀 1초 빠른 등록", use_container_width=True, type="primary"):
                        if check_recent_duplicate(today_str_kst, quick_equip, quick_prod, quick_measured):
                            st.error("⚠️ 방금 동일한 측정값이 등록되었습니다.")
                        else:
                            diff = round(quick_measured - quick_target, 1)
                            status = "합격 🟢" if abs(diff) <= 2.0 else "불합격 🔴"
                            save_to_db(today_str_kst, quick_equip, quick_worker, quick_prod, quick_target, quick_measured, diff, status, "", quick_amt)
                            st.cache_data.clear()
                            st.session_state['show_toast'] = "🚀 빠른 등록이 완료되었습니다!"
                            st.rerun()

st.markdown("---")

# ==========================================
# 5. 화면 구성: 데이터 조회 표 (조회 및 필터)
# ==========================================
st.subheader("📊 누적 측정 기록 조회")

if st_autorefresh is not None:
    auto_refresh = st.checkbox("🔄 실시간 모니터링 켜기 (10초마다 화면 자동 새로고침)", key="chk_autorefresh")
    if auto_refresh:
        st_autorefresh(interval=10000, limit=None, key="main_auto_refresh")

col_filter1, col_filter2, col_filter3 = st.columns(3)
with col_filter1:
    search_query = st.text_input("🔍 제품명 검색 (전체 조회는 빈칸)", key="input_search").strip()
with col_filter2:
    date_filter_mode = st.radio("📅 조회 기간 선택", ["오늘(Today)", "전체 기간", "특정 일자 지정"], index=0, horizontal=True, key="radio_date_mode")

filter_date_str = ""
with col_filter3:
    if date_filter_mode == "특정 일자 지정":
        filter_date = st.date_input("조회할 생산일 선택", value=get_now_kst().date(), key="date_picker")
        filter_date_str = filter_date.strftime("%Y-%m-%d")

def get_equip_sort_order(val):
    val_clean = str(val).lower().replace(" ", "")
    if '버닝' in val_clean: return 0
    elif '태환' in val_clean: return 1
    elif '프로밧' in val_clean: return 2
    elif '뷸러60' in val_clean or ('뷸러' in val_clean and '60' in val_clean): return 3
    elif '뷸러120' in val_clean or ('뷸러' in val_clean and '120' in val_clean): return 4
    else: return 5

display_df = history_df.copy()

if not display_df.empty:
    if search_query:
        display_df = display_df[display_df['제품명'].astype(str).str.contains(search_query, na=False)]
    
    if date_filter_mode == "오늘(Today)":
        display_df = display_df[display_df['생산일'] == today_str_kst]
        export_file_name = f"색도측정기록_{today_str_kst}.xlsx"
    elif date_filter_mode == "특정 일자 지정":
        display_df = display_df[display_df['생산일'] == filter_date_str]
        export_file_name = f"색도측정기록_{filter_date_str}.xlsx"
    else:
        export_file_name = "색도측정기록_전체기간누적.xlsx"

    if not display_df.empty:
        display_df['정렬순서'] = display_df['생산설비'].apply(get_equip_sort_order)
        display_df = display_df.sort_values(by=['정렬순서', '고유번호'], ascending=[True, False])
        display_df = display_df.drop(columns=['정렬순서'])

total_batches = len(display_df)

if date_filter_mode == "오늘(Today)": metric_title = "오늘 총 생산"
elif date_filter_mode == "특정 일자 지정": metric_title = f"{filter_date_str} 총 생산"
else: metric_title = f"전체 기간 누적" if not search_query else f"'{search_query}' 검색"

equip_counts = display_df['생산설비'].value_counts() if not display_df.empty else pd.Series()
present_equips = [eq for eq in EQUIPMENT_LIST if eq in equip_counts.index]

metric_cols = st.columns(1 + len(present_equips))
with metric_cols[0]: st.metric(label=f"📦 {metric_title} 배치 수", value=f"{total_batches} 건")
for idx, eq in enumerate(present_equips):
    with metric_cols[idx + 1]: st.metric(label=f"⚙️ {eq}", value=f"{equip_counts[eq]} 건")
        
st.markdown("<br>", unsafe_allow_html=True)

# [UI 개선 3] 판정 결과의 조건부 서식 강화 (불합격시 붉은색 강렬한 표시)
def highlight_status(s):
    return ['color: white; background-color: #E74C3C; font-weight: bold;' if '불합격' in str(v) else 'color: #27AE60; font-weight: bold;' for v in s]

def highlight_equipment(s):
    colors = []
    for val in s:
        val_clean = str(val).lower().replace(" ", "")
        if '버닝' in val_clean: colors.append('background-color: #E1F5FE; color: black; font-weight: bold;') 
        elif '태환' in val_clean: colors.append('background-color: #FFF3CD; color: black; font-weight: bold;') 
        elif '프로밧' in val_clean: colors.append('background-color: #FCE4EC; color: black; font-weight: bold;') 
        elif '뷸러60' in val_clean or ('뷸러' in val_clean and '60' in val_clean): colors.append('background-color: #E8F5E9; color: black; font-weight: bold;') 
        elif '뷸러120' in val_clean or ('뷸러' in val_clean and '120' in val_clean): colors.append('background-color: #D4EFDF; color: black; font-weight: bold;') 
        else: colors.append('')
    return colors

if not display_df.empty:
    main_display_df = display_df.drop(columns=['확인여부'], errors='ignore')
    
    styled_df = main_display_df.style.format(
        {"측정색도": "{:.1f}", "기준색도": "{:.1f}", "오차": "{:.1f}"}, 
        na_rep="-"
    ).apply(
        highlight_equipment, subset=['생산설비']
    ).apply(
        highlight_status, subset=['판정']
    ).set_properties(
        subset=['특이사항'], 
        **{'background-color': '#E8DAEF', 'color': 'black', 'font-weight': 'bold'}
    ).set_properties(
        subset=['제품명'], 
        **{'font-weight': 'bold'}
    )

    st.dataframe(
        styled_df, 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "특이사항": st.column_config.TextColumn("특이사항", width="large"),
            "제품명": st.column_config.TextColumn("제품명", width="medium"),
            "작업자": st.column_config.TextColumn("작업자", width="small"),
            "투입량": st.column_config.TextColumn("투입량", width="small")
        }
    )
    
    st.markdown("<br>", unsafe_allow_html=True)
    excel_data = to_excel(main_display_df)
    st.download_button(label="📥 현재 화면의 표를 엑셀 파일로 다운로드", data=excel_data, file_name=export_file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info(f"🔍 현재 선택하신 조건({date_filter_mode})에 일치하는 기록이 없습니다.")
