import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import io 
import pytz # 한국 시간 처리를 위해 추가

# 한국 시간대 설정
KST = pytz.timezone('Asia/Seoul')

st.set_page_config(page_title="색도 관리 시스템", layout="wide")

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
            st.success("인증에 성공했습니다! 시스템을 로딩합니다.")
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
    
    try:
        cursor.execute("ALTER TABLE color_records ADD COLUMN remarks TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass 
        
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

    # [신규 업데이트] 투입량 데이터 저장을 위한 DB 컬럼 자동 생성 안전장치
    try:
        cursor.execute("ALTER TABLE color_records ADD COLUMN input_amount TEXT DEFAULT '-'")
    except sqlite3.OperationalError:
        pass

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

# [수정] save_to_db 함수에 투입량(input_amount) 인자 주입 유도
def save_to_db(prod_date, equipment, worker, product, target, measured, diff, status, remarks, input_amount="-"):
    timestamp = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (timestamp, prod_date, equipment, worker, product, target, measured, diff, status, remarks, input_amount))
    conn.commit()
    conn.close()

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
            c.remarks as 특이사항,
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
        ORDER BY c.production_date DESC, c.timestamp DESC, c.id DESC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    if df.empty:
        return pd.DataFrame(columns=['생산일', '제품명', '생산설비', '투입량', '측정색도', '오차', '기준색도', '작업자', '판정', '특이사항', '입력일시', '고유번호'])

    df['측정색도'] = pd.to_numeric(df['측정색도'], errors='coerce')
    df['기준색도'] = pd.to_numeric(df['기준색도'], errors='coerce')
    df['오차'] = (df['측정색도'] - df['기준색도'])
    
    df['판정'] = "합격 🟢"
    df.loc[df['오차'].abs() > 2.0, '판정'] = "불합격 🔴"
    df.loc[df['오차'].isna(), '판정'] = "오류"
    
    df['특이사항'] = df['특이사항'].fillna('')
    
    desired_order = ['생산일', '제품명', '생산설비', '투입량', '측정색도', '오차', '기준색도', '작업자', '판정', '특이사항', '입력일시', '고유번호']
    return df[desired_order]

def delete_from_db(record_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM color_records WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()

# [수정] update_db 함수에 new_input_amount 매개변수 적용
def update_db(record_id, new_date, new_equip, new_worker, new_measured, new_diff, new_status, new_remarks, new_input_amount):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE color_records 
        SET production_date=?, equipment=?, worker=?, measured_value=?, difference=?, status=?, remarks=?, input_amount=?
        WHERE id=?
    ''', (new_date, new_equip, new_worker, new_measured, new_diff, new_status, new_remarks, new_input_amount, record_id))
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
    ''', (product_name,))
    row = cursor.fetchone()
    conn.close()
    return row

@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='색도측정기록')
    return output.getvalue()

# ==========================================
# 2. 엑셀 기준표 로드
# ==========================================
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

init_db()
TARGET_DATA = load_target_data()
sync_target_history(TARGET_DATA)

today_str_kst = get_now_kst().strftime("%Y-%m-%d")
ACTIVE_NOTICES = get_all_active_notices(today_str_kst)

# ==========================================
# 3. 화면 구성: 왼쪽 사이드바 (데이터 입력)
# ==========================================
col_title, col_logout = st.columns([9, 1])
with col_title:
    st.title("🎨 일일 제품 색도 관리 시스템")
with col_logout:
    if st.button("🔒 로그아웃"):
        st.query_params.clear()
        st.session_state['logged_in'] = False
        st.rerun()
st.markdown("---")

st.sidebar.header("📝 새로운 데이터 입력")

production_date_input = st.sidebar.date_input("생산일 선택", value=get_now_kst().date())
prod_date_str = production_date_input.strftime("%Y-%m-%d")

selected_equipment = st.sidebar.selectbox("생산 설비 선택", EQUIPMENT_LIST)

# ---------------------------------------------------------
# [신규 기능] 버닝 생산 설비일 때만 원료 투입량 선택 selectbox 표출
# ---------------------------------------------------------
input_amount_val = "-"
if "버닝" in str(selected_equipment).lower().replace(" ", ""):
    input_amount_val = st.sidebar.selectbox("원료(생두) 투입량 선택", ["1.35kg", "2.5kg", "3.75kg"])
# ---------------------------------------------------------

worker_name = st.sidebar.text_input("작업자 이름", placeholder="예) 문병국")

st.sidebar.markdown("---")
selected_product = st.sidebar.selectbox("🔍 제품명 검색 및 선택 (클릭 후 타이핑)", list(TARGET_DATA.keys()))

if ACTIVE_NOTICES.get(selected_product):
    st.sidebar.warning(f"📢 **[작업자 전달사항]**\n\n{ACTIVE_NOTICES[selected_product]}", icon="🚨")
    st.toast(f"**{selected_product}** 전달사항이 있습니다! 사이드바를 확인하세요.", icon="🚨")
else:
    st.sidebar.info("📢 **공지사항 없음**")

target_value = get_historical_target(selected_product, prod_date_str)
st.sidebar.info(f"📌 해당 생산일({prod_date_str})의 기준 색도: **{float(target_value):.1f}**")

target_history_df = get_target_history_df(selected_product)
real_changes = target_history_df[~target_history_df['effective_date'].isin(['2000-01-01', '2024-04-11', ''])]

if real_changes.empty:
    with st.sidebar.expander("📖 기준값 변경 이력 (변경 없음)"):
        st.caption("✨ 이 제품은 최초 등록 이후 기준 색도가 변경된 적이 없습니다.")
else:
    with st.sidebar.expander("📖 이 제품의 기준값 변경 이력 (이력 존재)"):
        target_history_df['effective_date'] = target_history_df['effective_date'].replace('2000-01-01', '최초 등록')
        target_history_df['effective_date'] = target_history_df['effective_date'].replace('2024-04-11', '최초 등록')
        target_history_df.columns = ['적용 시작일', '기준색도']
        st.dataframe(target_history_df.style.format({"기준색도": "{:.1f}"}), hide_index=True, use_container_width=True)

last_record = get_last_record(selected_product)
if last_record:
    last_date_str, last_measured, last_status = last_record
    
    try:
        last_date_obj = datetime.strptime(last_date_str, "%Y-%m-%d").date()
        is_old = (get_now_kst().date() - last_date_obj).days > 120
    except:
        is_old = False
        
    try:
        last_measured_fmt = f"{float(last_measured):.1f}"
    except:
        last_measured_fmt = str(last_measured)

    display_date = f":red[**{last_date_str} (4개월 초과!)**]" if is_old else last_date_str
    display_measured = f":red[**{last_measured_fmt} (이전 불합격!)**]" if "불합격" in last_status else str(last_measured_fmt)
    
    st.sidebar.info(f"🕒 **이전 최종 생산일:** {display_date}\n\n📉 **이전 측정 색도:** {display_measured}")
else:
    st.sidebar.warning("이전 생산 기록이 없습니다 (최초 입력).")

st.sidebar.markdown("---")

measured_value = st.sidebar.number_input("측정 색도 입력", value=float(target_value), step=0.1, format="%.1f")
remarks_input = st.sidebar.text_input("특이사항 (선택사항)", placeholder="간단한 메모 입력")

if st.sidebar.button("데이터 등록하기"):
    if worker_name.strip() == "":
        st.sidebar.warning("⚠️ 작업자 이름을 입력해 주세요!")
    else:
        difference = round(measured_value - target_value, 1)
        status = "합격 🟢" if abs(difference) <= 2.0 else "불합격 🔴"
        
        # [수정] 데이터 등록 시 input_amount_val 변수를 함께 기록
        save_to_db(prod_date_str, selected_equipment, worker_name, selected_product, target_value, measured_value, difference, status, remarks_input, input_amount_val)
        
        st.cache_data.clear()
        st.success(f"정상적으로 기록되었습니다.")
        st.rerun()

# ==========================================
# 4. 화면 구성: 메인 화면 (조회 및 필터)
# ==========================================
st.subheader("📊 누적 측정 기록 조회")

history_df = load_from_db()

col_filter1, col_filter2, col_filter3 = st.columns(3)
with col_filter1:
    search_query = st.text_input("🔍 제품명 검색 (전체 조회는 빈칸)")
with col_filter2:
    date_filter_mode = st.radio("📅 조회 기간 선택", ["오늘(Today)", "전체 기간", "특정 일자 지정"], index=0, horizontal=True)
with col_filter3:
    if date_filter_mode == "특정 일자 지정":
        filter_date = st.date_input("조회할 생산일 선택", value=get_now_kst().date())

# 설비 정렬 함수 (대소문자/띄어쓰기 무시)
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
        display_df = display_df[display_df['제품명'].str.contains(search_query, na=False)]
    
    if date_filter_mode == "오늘(Today)":
        display_df = display_df[display_df['생산일'] == today_str_kst]
        display_df['정렬순서'] = display_df['생산설비'].apply(get_equip_sort_order)
        display_df = display_df.sort_values(by=['정렬순서', '고유번호'], ascending=[True, True])
        display_df = display_df.drop(columns=['정렬순서'])
        export_file_name = f"색도측정기록_{today_str_kst}.xlsx"
        
    elif date_filter_mode == "특정 일자 지정":
        filter_date_str = filter_date.strftime("%Y-%m-%d")
        display_df = display_df[display_df['생산일'] == filter_date_str]
        display_df['정렬순서'] = display_df['생산설비'].apply(get_equip_sort_order)
        display_df = display_df.sort_values(by=['정렬순서', '고유번호'], ascending=[True, True])
        display_df = display_df.drop(columns=['정렬순서'])
        export_file_name = f"색도측정기록_{filter_date_str}.xlsx"
        
    else:
        export_file_name = "색도측정기록_전체기간누적.xlsx"

    if not display_df.empty:
        display_df['특이사항'] = display_df['특이사항'].astype(str).apply(
            lambda x: x.replace("[마지막 배치 🏁]", "").strip()
        )
        
        idx_latest = display_df.groupby('생산일')['고유번호'].idxmax()
        for idx in idx_latest:
            current_remark = display_df.loc[idx, '특이사항']
            if current_remark:
                display_df.loc[idx, '특이사항'] = f"[마지막 배치 🏁] {current_remark}"
            else:
                display_df.loc[idx, '특이사항'] = "[마지막 배치 🏁]"

if not (date_filter_mode == "전체 기간" and not search_query):
    total_batches = len(display_df)
    
    if date_filter_mode == "오늘(Today)":
        metric_title = "오늘 총 생산"
    elif date_filter_mode == "특정 일자 지정":
        metric_title = f"{filter_date_str} 총 생산"
    else:
        metric_title = f"'{search_query}' 총 생산"

    st.metric(label=f"📦 {metric_title} 배치 수", value=f"{total_batches} 건")
    st.markdown("<br>", unsafe_allow_html=True)

# 글자색 검정 통일 및 표 배경 색상 맵핑 함수
def highlight_equipment(s):
    colors = []
    for val in s:
        val_clean = str(val).lower().replace(" ", "")
        if '버닝' in val_clean:
            colors.append('background-color: #E1F5FE; color: black; font-weight: bold;') 
        elif '태환' in val_clean:
            colors.append('background-color: #FFF3CD; color: black; font-weight: bold;') 
        elif '프로밧' in val_clean:
            colors.append('background-color: #FCE4EC; color: black; font-weight: bold;') 
        elif '뷸러60' in val_clean or ('뷸러' in val_clean and '60' in val_clean):
            colors.append('background-color: #E8F5E9; color: black; font-weight: bold;') 
        elif '뷸러120' in val_clean or ('뷸러' in val_clean and '120' in val_clean):
            colors.append('background-color: #D4EFDF; color: black; font-weight: bold;') 
        else:
            colors.append('')
    return colors

if not display_df.empty:
    styled_df = display_df.style.format(
        {"측정색도": "{:.1f}", "기준색도": "{:.1f}", "오차": "{:.1f}"}, 
        na_rep="-"
    ).apply(
        highlight_equipment, subset=['생산설비']
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
            "투입량": st.column_config.TextColumn("투입량", width="small") # 투입량 컬럼 세팅 추가
        }
    )
    
    # ⚡ 진행 중인 제품 빠른 추가 (Fast Track)
    if date_filter_mode == "오늘(Today)":
        st.markdown("---")
        st.markdown("### ⚡ 진행 중인 라인 빠른 추가")
        st.info("오늘 이미 생산한 기록이 있는 제품은 특이사항이 없다면 아래에서 **'측정 색도'**만 입력하여 즉시 추가 기록됩니다.")
        
        recent_batches = display_df[['제품명', '생산설비', '작업자']].drop_duplicates().reset_index(drop=True)
        batch_options = [f"▶ {row['제품명']} (설비: {row['생산설비']} / 작업자: {row['작업자']})" for idx, row in recent_batches.iterrows()]
        
        selected_batch_str = st.selectbox("이어서 측정할 제품을 선택하세요", batch_options)
        
        if selected_batch_str:
            selected_idx = batch_options.index(selected_batch_str)
            quick_prod = recent_batches.iloc[selected_idx]['제품명']
            quick_equip = recent_batches.iloc[selected_idx]['생산설비']
            quick_worker = recent_batches.iloc[selected_idx]['작업자']
            quick_target = get_historical_target(quick_prod, today_str_kst)
            
            # [신규 기능] 빠른 추가 라인에서도
