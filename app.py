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

def save_to_db(prod_date, equipment, worker, product, target, measured, diff, status, remarks):
    timestamp = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (timestamp, prod_date, equipment, worker, product, target, measured, diff, status, remarks))
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
        return pd.DataFrame(columns=['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '판정', '특이사항', '입력일시', '고유번호'])

    df['측정색도'] = pd.to_numeric(df['측정색도'], errors='coerce')
    df['기준색도'] = pd.to_numeric(df['기준색도'], errors='coerce')
    df['오차'] = (df['측정색도'] - df['기준색도'])
    
    df['판정'] = "합격 🟢"
    df.loc[df['오차'].abs() > 2.0, '판정'] = "불합격 🔴"
    df.loc[df['오차'].isna(), '판정'] = "오류"
    
    df['특이사항'] = df['특이사항'].fillna('')
    
    desired_order = ['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '판정', '특이사항', '입력일시', '고유번호']
    return df[desired_order]

def delete_from_db(record_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM color_records WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()

def update_db(record_id, new_date, new_equip, new_worker, new_measured, new_diff, new_status, new_remarks):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE color_records 
        SET production_date=?, equipment=?, worker=?, measured_value=?, difference=?, status=?, remarks=?
        WHERE id=?
    ''', (new_date, new_equip, new_worker, new_measured, new_diff, new_status, new_remarks, record_id))
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
        
        save_to_db(prod_date_str, selected_equipment, worker_name, selected_product, target_value, measured_value, difference, status, remarks_input)
        
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

    # ---------------------------------------------------------
    # [핵심 자동화 기능 완벽 수정] 데일리별(생산일 + 제품명 기준) 자동 플래그 부여
    # ---------------------------------------------------------
    if not display_df.empty:
        # 무작정 한 개만 표시되는 문제를 해결하기 위해 그룹 기준에 ['생산일', '제품명']을 적용
        idx_latest = display_df.groupby(['생산일', '제품명'])['고유번호'].idxmax()
        for idx in idx_latest:
            current_remark = display_df.loc[idx, '특이사항']
            if current_remark:
                display_df.loc[idx, '특이사항'] = f"[마지막 배치 🏁] {current_remark}"
            else:
                display_df.loc[idx, '특이사항'] = "[마지막 배치 🏁]"
    # ---------------------------------------------------------

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
            "작업자": st.column_config.TextColumn("작업자", width="small")
        }
    )
    
    # ⚡ 진행 중인 제품 빠른 추가 (Fast Track)
    if date_filter_mode == "오늘(Today)":
        st.markdown("---")
        st.markdown("### ⚡ 진행 중인 라인 빠른 추가")
        st.info("오늘 이미 생산한 기록이 있는 제품은 특이사항이 없다면 아래에서 **'측정 색도'**만 입력하여 즉시 추가 기록됩니다.")
        
        recent_batches = display_df[['제품명', '생산설비', '작업자']].drop_duplicates().reset_index(drop=True)
        batch_options = [f"▶ {row['제품명']} (설비: {row['생산설비']} / 작업자: {row['작업자']})" for idx, row in recent_batches.iterrows()]
        
        col_q1, col_q2, col_q3, col_q4 = st.columns([3, 1, 1, 1])
        
        with col_q1:
            selected_batch_str = st.selectbox("이어서 측정할 제품을 선택하세요", batch_options)
            
        if selected_batch_str:
            selected_idx = batch_options.index(selected_batch_str)
            quick_prod = recent_batches.iloc[selected_idx]['제품명']
            quick_equip = recent_batches.iloc[selected_idx]['생산설비']
            quick_worker = recent_batches.iloc[selected_idx]['작업자']
            quick_target = get_historical_target(quick_prod, today_str_kst)
            
            with col_q2:
                st.text_input("기준 색도", value=f"{float(quick_target):.1f}", disabled=True)
            with col_q3:
                quick_measured = st.number_input("새 측정값", value=float(quick_target), step=0.1, format="%.1f", key="quick_val")
            with col_q4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🚀 1초 등록", use_container_width=True, type="primary"):
                    diff = round(quick_measured - quick_target, 1)
                    status = "합격 🟢" if abs(diff) <= 2.0 else "불합격 🔴"
                    
                    save_to_db(today_str_kst, quick_equip, quick_worker, quick_prod, quick_target, quick_measured, diff, status, "")
                    st.cache_data.clear()
                    st.success("빠른 등록이 완료되었습니다!")
                    st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    excel_data = to_excel(display_df)
    st.download_button(label="📥 현재 화면의 표를 엑셀 파일로 다운로드", data=excel_data, file_name=export_file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info(f"🔍 현재 선택하신 조건({date_filter_mode})에 일치하는 기록이 없습니다.")

st.markdown("---")

# ==========================================
# 5. 화면 구성: 관리자 도구
# ==========================================
with st.expander("🛠️ 관리자 전용 메뉴 (데이터 수정/삭제 및 관리)"):
    input_password = st.text_input("🔒 관리자 비밀번호를 입력하세요", type="password")
    
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
            st.markdown("<small>※ 추후 계정 이전 시 인수인계용 파일로 사용됩니다.</small><br>", unsafe_allow_html=True)
        except Exception as e:
            pass
        
        tab1, tab2, tab3, tab4 = st.tabs(["개별 데이터 수정/삭제", "📂 과거 엑셀 업로드", "📅 기준값 이력 업로드", "📢 제품 공지사항 관리"])
        
        with tab1:
            st.write("위 표의 **'고유번호'**를 확인한 후 작업을 진행하세요.")
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
                        st.success(f"고유번호 {target_id} 삭제됨.")
                        st.rerun()
                        
                elif action == "데이터 수정":
                    conn = sqlite3.connect(DB_FILE)
                    c = conn.cursor()
                    c.execute("SELECT product_name, target_value, production_date, equipment, worker, measured_value, remarks FROM color_records WHERE id=?", (target_id,))
                    row = c.fetchone()
                    conn.close()
                    
                    if row:
                        p_name, t_val, p_date, equip, workr, m_val, rmks = row
                        
                        try: curr_date = datetime.strptime(p_date, "%Y-%m-%d").date()
                        except: curr_date = get_now_kst().date()
                            
                        new_p_date = st.date_input("수정할 생산일 지정", value=curr_date)
                        new_p_date_str = new_p_date.strftime("%Y-%m-%d")
                        
                        actual_t_val = get_historical_target(p_name, new_p_date_str)
                        st.info(f"선택 제품: **{p_name}** (지정한 날짜의 기준색도: {float(actual_t_val):.1f})")
                        
                        try: equip_index = EQUIPMENT_LIST.index(equip)
                        except: equip_index = 0
                            
                        new_equip = st.selectbox("수정할 설비 지정", EQUIPMENT_LIST, index=equip_index)
                        new_worker = st.text_input("수정할 작업자 이름", value=workr)
                        new_m_val = st.number_input("수정할 측정색도 지정", value=float(m_val), step=0.1, format="%.1f")
                        new_rmks = st.text_input("수정할 특이사항", value=rmks if rmks else "")
                        
                        if st.button("✏️ 선택한 데이터 수정"):
                            new_diff = round(new_m_val - actual_t_val, 1)
                            new_status = "합격 🟢" if abs(new_diff) <= 2.0 else "불합격 🔴"
                            update_db(target_id, new_p_date_str, new_equip, new_worker, new_m_val, new_diff, new_status, new_remarks)
                            st.cache_data.clear()
                            st.success(f"고유번호 {target_id} 수정됨.")
                            st.rerun()
                    else:
                        st.error("해당 고유번호가 없습니다.")
            
            st.markdown("---")
            st.write("🔧 **[시스템 유지보수]** 과거 데이터의 대문자 'KG'를 소문자 'kg'로 일괄 변경합니다.")
            if st.button("🚀 전체 과거 데이터 'kg' 일괄 변환 실행"):
                conn = sqlite3.connect(DB_FILE)
                cursor = conn.cursor()
                cursor.execute("UPDATE color_records SET equipment = REPLACE(equipment, 'KG', 'kg') WHERE equipment LIKE '%KG%'")
                updated_rows = cursor.rowcount
                conn.commit()
                conn.close()
                st.cache_data.clear()
                st.success(f"🎉 총 {updated_rows}건의 과거 데이터가 성공적으로 소문자 'kg'로 변환되었습니다!")
                st.rerun()
        
        with tab2:
            st.write("과거에 측정했던 엑셀 기록을 업로드하면 DB에 일괄 저장됩니다.")
            st.info("💡 **필수 열 이름:** `생산일`, `제품명`, `생산설비`, `작업자`, `측정색도`")
            
            uploaded_file = st.file_uploader("과거 측정 기록 엑셀 파일 선택", type=['xlsx', 'xls'], key="record_upload")
            
            if uploaded_file is not None:
                if st.button("🚀 측정 기록 일괄 업로드 실행"):
                    try:
                        df_upload = pd.read_excel(uploaded_file)
                        required_cols = ['생산일', '제품명', '생산설비', '작업자', '측정색도']
                        
                        if all(col in df_upload.columns for col in required_cols):
                            success_count = 0
                            skip_count = 0 
                            
                            for index, row in df_upload.iterrows():
                                measured_str = str(row['측정색도']).strip()
                                if measured_str in ['-', '', 'nan', 'None']:
                                    skip_count += 1
                                    continue
                                
                                try: measured = float(measured_str)
                                except ValueError: skip_count += 1; continue

                                try: prod_date_str = pd.to_datetime(row['생산일']).strftime("%Y-%m-%d")
                                except: prod_date_str = str(row['생산일'])[:10]
                                
                                product = str(row['제품명']).strip()
                                equip = str(row['생산설비'])
                                worker = str(row['작업자'])
                                
                                upload_remarks = ""
                                if '특이사항' in df_upload.columns and not pd.isna(row.get('특이사항')):
                                    upload_remarks = str(row['특이사항']).strip()
                                    if upload_remarks in ['nan', 'None']: upload_remarks = ""
                                
                                target = get_historical_target(product, prod_date_str)
                                diff = round(measured - target, 1)
                                status = "합격 🟢" if abs(diff) <= 2.0 else "불합격 🔴"
                                current_time = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
                                
                                conn = sqlite3.connect(DB_FILE)
                                cursor = conn.cursor()
                                cursor.execute('''
                                    INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (current_time, prod_date_str, equip, worker, product, target, measured, diff, status, upload_remarks))
                                conn.commit()
                                conn.close()
                                
                                success_count += 1
                                
                            if skip_count > 0:
                                st.warning(f"⚠️ 측정값이 비어있거나 기호('-')로 된 {skip_count}건은 제외되었습니다.")
                            
                            st.cache_data.clear()
                            st.success(f"🎉 총 {success_count}건 데이터 저장 성공!")
                            st.rerun()
                        else:
                            st.error("❌ 엑셀 파일에 필수 열이 부족합니다.")
                    except Exception as e:
                        st.error(f"오류 발생: {e}")

        with tab3:
            st.write("과거에 제품의 기준값이 변경되었던 이력(History)을 엑셀로 일괄 주입합니다.")
            history_file = st.file_uploader("기준값 이력 관리 엑셀 파일 선택", type=['xlsx', 'xls'], key="history_upload")
            
            if history_file is not None:
                if st.button("🚀 기준값 이력 일괄 반영"):
                    try:
                        df_history = pd.read_excel(history_file)
                        req_hist_cols = ['제품명', '적용시작일', '기준색도']
                        
                        if all(col in df_history.columns for col in req_hist_cols):
                            conn = sqlite3.connect(DB_FILE)
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM target_history")
                            
                            hist_count = 0
                            for index, row in df_history.iterrows():
                                product = str(row['제품명']).strip()
                                eff_date = str(row['적용시작일']).strip()
                                
                                if pd.isna(row['적용시작일']) or eff_date in ['nan', 'None', '', 'NaT']:
                                    eff_date_str = '2000-01-01'
                                else:
                                    try: eff_date_str = pd.to_datetime(row['적용시작일']).strftime("%Y-%m-%d")
                                    except: eff_date_str = eff_date[:10]
                                    
                                try:
                                    target_val = float(row['기준색도'])
                                    cursor.execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (?, ?, ?)", (product, target_val, eff_date_str))
                                    hist_count += 1
                                except ValueError:
                                    continue 
                                    
                            conn.commit()
                            conn.close()
                            
                            st.cache_data.clear()
                            st.success(f"🎉 총 {hist_count}건 이력 장부 세팅 완료!")
                            st.rerun()
                        else:
                            st.error("❌ 필수 열이 부족합니다.")
                    except Exception as e:
                        st.error(f"오류 발생: {e}")
                        
        with tab4:
            st.info("💡 엑셀 업로드 없이 여기서 특정 제품의 전달사항(공지)을 직접 띄우고 기간을 설정할 수 있습니다.")
            notice_prod = st.selectbox("공지사항을 설정할 제품 선택", list(TARGET_DATA.keys()), key="notice_prod")
            
            raw_notice = get_raw_notice(notice_prod)
            curr_text = raw_notice[0] if raw_notice else ""
            curr_start = raw_notice[1] if raw_notice else get_now_kst().strftime("%Y-%m-%d")
            curr_end = raw_notice[2] if raw_notice else "2099-12-31"
            
            is_unlimited = (curr_end == "2099-12-31")
            
            try: start_dt = datetime.strptime(curr_start, "%Y-%m-%d").date()
            except: start_dt = get_now_kst().date()
                
            try: end_dt = datetime.strptime(curr_end, "%Y-%m-%d").date() if not is_unlimited else get_now_kst().date() + timedelta(days=7)
            except: end_dt = get_now_kst().date() + timedelta(days=7)

            notice_text = st.text_area("작업자에게 보여줄 공지 내용", value=curr_text, placeholder="예) 오늘 생두 수분량이 높으니 주의!")
            
            col_n1, col_n2 = st.columns(2)
            with col_n1:
                start_date = st.date_input("공지 시작일", value=start_dt)
            with col_n2:
                no_limit = st.checkbox("무기한 (기한 없음) - 계속 띄워두기", value=is_unlimited)
                end_date = st.date_input("공지 종료일", value=end_dt, disabled=no_limit)
                
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("📢 공지 등록 / 수정하기", type="primary", use_container_width=True):
                    if notice_text.strip() == "":
                        st.warning("공지 내용을 입력해 주세요.")
                    else:
                        end_str = "2099-12-31" if no_limit else end_date.strftime("%Y-%m-%d")
                        save_notice(notice_prod, notice_text, start_date.strftime("%Y-%m-%d"), end_str)
                        st.cache_data.clear()
                        st.success(f"[{notice_prod}] 공지사항이 등록되었습니다!")
                        st.rerun()
                        
            with col_b2:
                if st.button("🗑️ 이 제품의 공지 삭제", use_container_width=True):
                    delete_notice(notice_prod)
                    st.cache_data.clear()
                    st.success("공지가 성공적으로 삭제되었습니다.")
                    st.rerun()

    elif input_password != "":
        st.error("비밀번호가 일치하지 않습니다.")
