import streamlit as st
import pandas as pd
import sqlite3
import tempfile
import zipfile
from datetime import datetime, timedelta
import io 
import pytz
import numpy as np
import re
from pathlib import Path
from sqlalchemy import create_engine, text

KST = pytz.timezone('Asia/Seoul')

# [최적화] 공휴일 로드 및 오류 방지
try:
    import holidays
    HAS_HOLIDAYS = True
    holiday_end_year = datetime.now(KST).year + 10
    kr_holidays_dict = holidays.KR(years=range(2000, holiday_end_year + 1))
    KR_HOLIDAYS = [str(date) for date in kr_holidays_dict.keys()]
except ImportError:
    HAS_HOLIDAYS = False
    KR_HOLIDAYS = []

try: from streamlit_autorefresh import st_autorefresh
except ImportError: st_autorefresh = None

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
# 2. Neon PostgreSQL 데이터베이스
# 화면 코드가 기존 방식으로 SQL을 부를 수 있도록 연결 도우미를 사용합니다.
@st.cache_resource
def get_db_engine():
    try:
        database_url = st.secrets["connections"]["neon"]["url"]
    except Exception as e:
        raise RuntimeError(
            "Streamlit Secrets에 [connections.neon] 아래 url을 설정해야 합니다."
        ) from e

    if not str(database_url).startswith(("postgresql://", "postgresql+psycopg2://")):
        raise RuntimeError("Neon 연결 주소는 postgresql:// 로 시작해야 합니다.")

    return create_engine(
        database_url,
        pool_pre_ping=True,
        pool_recycle=300,
    )


class DatabaseConnection:
    """기존 화면의 execute/fetch/commit 코드를 PostgreSQL 연결로 이어줍니다."""

    def __init__(self):
        self.connection = get_db_engine().connect()

    @staticmethod
    def _prepare(sql, params=()):
        sql = str(sql).strip()

        # SQLite의 잠금 시작 명령을 PostgreSQL 트랜잭션 잠금으로 바꿉니다.
        if sql.upper() == "BEGIN IMMEDIATE":
            return None, None, True

        # PostgreSQL은 INSERT OR IGNORE 대신 ON CONFLICT를 사용합니다.
        sql = re.sub(
            r"^INSERT\s+OR\s+IGNORE\s+INTO\s+",
            "INSERT INTO ",
            sql,
            flags=re.IGNORECASE,
        )
        if "ON CONFLICT" not in sql.upper() and sql.upper().startswith("INSERT INTO EXCLUDED_PRODUCTS"):
            sql += " ON CONFLICT (product_name) DO NOTHING"

        # 물음표 자리를 SQLAlchemy가 이해하는 이름 있는 자리표시자로 바꿉니다.
        if isinstance(params, dict):
            return text(sql), params, False

        values = tuple(params) if params is not None else ()
        index = 0

        def replace_placeholder(_match):
            nonlocal index
            placeholder = f":p{index}"
            index += 1
            return placeholder

        converted_sql = re.sub(r"\?", replace_placeholder, sql)
        if index != len(values):
            raise ValueError("SQL 입력값 개수가 자리표시자 개수와 다릅니다.")
        bind_values = {f"p{i}": value for i, value in enumerate(values)}
        return text(converted_sql), bind_values, False

    def execute(self, sql, params=()):
        statement, bind_values, is_lock = self._prepare(sql, params)
        if is_lock:
            # 입력 저장이나 기준값 전체 교체가 동시에 실행되지 않게 짧게 잠급니다.
            return self.connection.execute(
                text("SELECT pg_advisory_xact_lock(893746221)")
            )
        return self.connection.execute(statement, bind_values)

    def executemany(self, sql, rows):
        rows = list(rows)
        if not rows:
            return None
        statement, _, is_lock = self._prepare(sql, rows[0])
        if is_lock:
            self.execute("BEGIN IMMEDIATE")
            return None

        converted_rows = []
        for row in rows:
            _, bind_values, _ = self._prepare(sql, row)
            converted_rows.append(bind_values)
        return self.connection.execute(statement, converted_rows)

    def commit(self):
        self.connection.commit()

    def rollback(self):
        self.connection.rollback()

    def close(self):
        self.connection.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if exc_type:
            self.rollback()
        self.close()


def get_db_conn():
    return DatabaseConnection()


def init_db():
    conn = get_db_conn()
    try:
        conn.execute("""CREATE TABLE IF NOT EXISTS color_records (
            id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            timestamp TEXT,
            production_date TEXT,
            equipment TEXT,
            worker TEXT,
            product_name TEXT,
            target_value DOUBLE PRECISION,
            measured_value DOUBLE PRECISION,
            difference DOUBLE PRECISION,
            status TEXT,
            remarks TEXT DEFAULT '',
            input_amount TEXT DEFAULT '-',
            checked INTEGER DEFAULT 0
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS target_history (
            id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            product_name TEXT,
            target_value DOUBLE PRECISION,
            effective_date TEXT
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS product_notices (
            product_name TEXT PRIMARY KEY,
            notice_text TEXT,
            start_date TEXT,
            end_date TEXT
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS workers (
            name TEXT PRIMARY KEY
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS excluded_products (
            product_name TEXT PRIMARY KEY
        )""")

        conn.execute("ALTER TABLE color_records ADD COLUMN IF NOT EXISTS remarks TEXT DEFAULT ''")
        conn.execute("ALTER TABLE color_records ADD COLUMN IF NOT EXISTS input_amount TEXT DEFAULT '-'")
        conn.execute("ALTER TABLE color_records ADD COLUMN IF NOT EXISTS checked INTEGER DEFAULT 0")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_color_prod_date ON color_records(product_name, production_date)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_target_hist ON target_history(product_name, effective_date)")

        worker_count = conn.execute("SELECT COUNT(*) FROM workers").fetchone()[0]
        if worker_count == 0:
            conn.executemany(
                "INSERT INTO workers (name) VALUES (?) ON CONFLICT (name) DO NOTHING",
                [(w,) for w in ["윤승태", "문지원", "조성윤", "이태원"]],
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_all_workers():
    conn = get_db_conn()
    try:
        rows = conn.execute("SELECT name FROM workers ORDER BY name").fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def add_worker(name):
    conn = get_db_conn()
    try:
        result = conn.execute(
            "INSERT INTO workers (name) VALUES (?) ON CONFLICT (name) DO NOTHING",
            (name.strip(),),
        )
        conn.commit()
        return result.rowcount > 0
    except Exception:
        conn.rollback()
        return False
    finally:
        conn.close()


def delete_worker(name):
    conn = get_db_conn()
    try:
        conn.execute("DELETE FROM workers WHERE name = ?", (name,))
        conn.commit()
    finally:
        conn.close()


def get_excluded_products():
    conn = get_db_conn()
    try:
        rows = conn.execute("SELECT product_name FROM excluded_products ORDER BY product_name").fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def add_excluded_product(name):
    conn = get_db_conn()
    try:
        conn.execute("INSERT INTO excluded_products (product_name) VALUES (?)", (name.strip(),))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def remove_excluded_product(name):
    conn = get_db_conn()
    try:
        conn.execute("DELETE FROM excluded_products WHERE product_name = ?", (name.strip(),))
        conn.commit()
    finally:
        conn.close()


def update_checked_status(record_ids, status_val):
    conn = get_db_conn()
    try:
        for record_id in record_ids:
            conn.execute("UPDATE color_records SET checked=? WHERE id=?", (status_val, record_id))
        conn.commit()
    finally:
        conn.close()


def delete_from_db(record_id):
    conn = get_db_conn()
    try:
        conn.execute("DELETE FROM color_records WHERE id = ?", (record_id,))
        conn.commit()
    finally:
        conn.close()


def update_db(record_id, d_date, eq, wk, product, target, measured, diff, status, remarks, amount, checked=0):
    conn = get_db_conn()
    try:
        conn.execute(
            """UPDATE color_records
               SET production_date=?, equipment=?, worker=?, product_name=?, target_value=?,
                   measured_value=?, difference=?, status=?, remarks=?, input_amount=?, checked=?
               WHERE id=?""",
            (d_date, str(eq).strip(), str(wk).strip(), str(product).strip(), target,
             measured, diff, status, remarks, amount, checked, record_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_historical_target(product_name, date_text):
    conn = get_db_conn()
    try:
        row = conn.execute(
            """SELECT target_value FROM target_history
               WHERE product_name=? AND effective_date <= ?
               ORDER BY effective_date DESC, id DESC LIMIT 1""",
            (product_name, date_text),
        ).fetchone()
        return row[0] if row else None
    finally:
        conn.close()


def save_notice(product, notice_text, start_date, end_date):
    conn = get_db_conn()
    try:
        conn.execute(
            """INSERT INTO product_notices (product_name, notice_text, start_date, end_date)
               VALUES (?, ?, ?, ?)
               ON CONFLICT (product_name) DO UPDATE SET
                 notice_text=EXCLUDED.notice_text,
                 start_date=EXCLUDED.start_date,
                 end_date=EXCLUDED.end_date""",
            (product, notice_text, start_date, end_date),
        )
        conn.commit()
    finally:
        conn.close()


def delete_notice(product):
    conn = get_db_conn()
    try:
        conn.execute("DELETE FROM product_notices WHERE product_name = ?", (product,))
        conn.commit()
    finally:
        conn.close()


def get_all_active_notices(date_text):
    conn = get_db_conn()
    try:
        rows = conn.execute(
            "SELECT product_name, notice_text FROM product_notices WHERE start_date <= ? AND end_date >= ?",
            (date_text, date_text),
        ).fetchall()
        return {r[0]: r[1] for r in rows}
    finally:
        conn.close()


def get_raw_notice(product):
    conn = get_db_conn()
    try:
        return conn.execute(
            "SELECT notice_text, start_date, end_date FROM product_notices WHERE product_name=?",
            (product,),
        ).fetchone()
    finally:
        conn.close()


def _is_recent_duplicate(conn, date_text, equipment, product, measured_value):
    row = conn.execute(
        """SELECT measured_value, timestamp FROM color_records
           WHERE production_date=? AND equipment=? AND product_name=?
           ORDER BY id DESC LIMIT 1""",
        (date_text, str(equipment).strip(), str(product).strip()),
    ).fetchone()
    if not row or row[0] is None:
        return False
    try:
        if float(row[0]) != float(measured_value):
            return False
        stored_at = datetime.fromisoformat(row[1])
        if stored_at.tzinfo is None:
            stored_at = KST.localize(stored_at)
        else:
            stored_at = stored_at.astimezone(KST)
        age_seconds = (get_now_kst() - stored_at).total_seconds()
        return 0 <= age_seconds < 30
    except (TypeError, ValueError, OverflowError):
        return False


def save_to_db(date_text, equipment, worker, product, target, measured, difference, status, remarks, amount):
    conn = get_db_conn()
    try:
        conn.execute("BEGIN IMMEDIATE")
        if _is_recent_duplicate(conn, date_text, equipment, product, measured):
            conn.rollback()
            return False

        timestamp = get_now_kst().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            """INSERT INTO color_records
               (timestamp, production_date, equipment, worker, product_name, target_value,
                measured_value, difference, status, remarks, input_amount)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (timestamp, date_text, str(equipment).strip(), str(worker).strip(),
             str(product).strip(), target, measured, difference, status, remarks, amount),
        )
        conn.execute("DELETE FROM excluded_products WHERE product_name = ?", (str(product).strip(),))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def check_recent_duplicate(date_text, equipment, product, measured_value):
    conn = get_db_conn()
    try:
        return _is_recent_duplicate(conn, date_text, equipment, product, measured_value)
    finally:
        conn.close()


def get_last_record(product):
    conn = get_db_conn()
    try:
        return conn.execute(
            """SELECT production_date, measured_value, status FROM color_records
               WHERE product_name = ?
               ORDER BY production_date DESC, timestamp DESC LIMIT 1""",
            (str(product).strip(),),
        ).fetchone()
    finally:
        conn.close()


def get_equipment_last_records(product_name):
    conn = get_db_conn()
    try:
        query = """
            WITH RankedRecords AS (
                SELECT equipment, production_date, measured_value, status, id,
                       ROW_NUMBER() OVER (
                           PARTITION BY equipment
                           ORDER BY production_date DESC, timestamp DESC, id DESC
                       ) AS rn
                FROM color_records WHERE product_name = ?
            ),
            EquipCounts AS (
                SELECT equipment, COUNT(*) AS cnt
                FROM color_records WHERE product_name = ? GROUP BY equipment
            )
            SELECT r.equipment, r.production_date, r.measured_value, r.status, c.cnt
            FROM RankedRecords r JOIN EquipCounts c ON r.equipment = c.equipment
            WHERE r.rn = 1 ORDER BY c.cnt DESC, r.production_date DESC
        """
        return conn.execute(query, (str(product_name).strip(), str(product_name).strip())).fetchall()
    finally:
        conn.close()


def auto_fill_input_amount(row):
    eq = str(row['생산설비']).lower().replace(" ", "")
    amt = str(row['투입량']).strip()
    if '버닝' in eq: return amt if amt != "" else "-"
    if amt in ["", "-", "nan", "None"]:
        if "태환" in eq: return "12kg"
        elif "프로밧" in eq: return "25kg"
        elif "60" in eq: return "60kg"
        elif "120" in eq: return "125kg"
    return amt

@st.cache_data(show_spinner=False, ttl=600)
def load_from_db():
    conn = get_db_conn()
    try:
        q = """
        SELECT 
            c.id as 고유번호, c.timestamp as 입력일시, c.production_date as 생산일, 
            c.equipment as 생산설비, COALESCE(c.input_amount, '-') as 투입량, 
            c.worker as 작업자, c.product_name as 제품명, c.measured_value as 측정색도, 
            COALESCE(c.remarks, '') as 특이사항, COALESCE(c.checked, 0) as checked_status, 
            COALESCE(
                (SELECT target_value FROM target_history th WHERE th.product_name = c.product_name AND th.effective_date <= c.production_date ORDER BY th.effective_date DESC, th.id DESC LIMIT 1),
                c.target_value
            ) as 기준색도
        FROM color_records c
        """
        try: df = pd.read_sql_query(q, conn.connection)
        except Exception as e:
            st.error(f"Neon에서 생산 기록을 읽지 못했습니다: {type(e).__name__}: {e}")
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
        df.loc[df['기준색도'].isna(), '판정'] = "기준 없음 ⚪"
        
        # 가짜 태그 흔적 텍스트 완벽 정화(정규식)
        df['특이사항'] = df['특이사항'].fillna('').astype(str)
        df['특이사항'] = df['특이사항'].str.replace(r'\[설비 첫\s*배치\s*🚀\]\s*', '', regex=True)
        df['특이사항'] = df['특이사항'].str.replace(r'\[마지막 배치\s*🏁\]\s*', '', regex=True)
        df['특이사항'] = df['특이사항'].str.replace(r'\[기준값 변경 후 첫 생산\s*🔔\]\s*', '', regex=True)
        df['특이사항'] = df['특이사항'].str.replace("nan", "", regex=False).str.strip()

        # 내부 연산용 시간 역순 정렬 (최신이 상단, 과거가 하단)
        df = df.sort_values(by=['생산일', '입력일시', '고유번호'], ascending=[False, False, False]).reset_index(drop=True)

        th_df = pd.read_sql_query("SELECT product_name, effective_date FROM target_history WHERE effective_date NOT IN ('2000-01-01', '2024-04-11', '')", conn.connection)
        target_change_first_ids = set()
        for _, r in th_df.iterrows():
            sub = df[(df['제품명'] == r['product_name'].strip()) & (df['생산일'] >= r['effective_date'])]
            if not sub.empty: target_change_first_ids.add(sub.iloc[-1]['고유번호'])
        
        if target_change_first_ids:
            mask = df['고유번호'].isin(target_change_first_ids)
            df.loc[mask, '특이사항'] = "[기준값 변경 후 첫 생산 🔔] " + df.loc[mask, '특이사항']

        # 제품명과 설비명 띄어쓰기 무시하고 역사상 첫 배치 찾기
        df['norm_p'] = df['제품명'].str.replace(" ", "").str.lower()
        df['norm_e'] = df['생산설비'].str.replace(" ", "").str.lower()
        
        f_idx = df.groupby(['norm_p', 'norm_e']).tail(1).index
        df.loc[f_idx, '특이사항'] = "[설비 첫 배치 🚀] " + df.loc[f_idx, '특이사항']
        
        df = df.drop(columns=['norm_p', 'norm_e'])
        
        # 설비와 무관하게 "공장 전체 그날(당일)의 제일 마지막 단일 배치"에만 표시
        l_idx = df.groupby('생산일').head(1).index
        df.loc[l_idx, '특이사항'] = "[마지막 배치 🏁] " + df.loc[l_idx, '특이사항']
        
        df['특이사항'] = df['특이사항'].str.strip()
        return df[['생산일', '제품명', '생산설비', '측정색도', '오차', '기준색도', '작업자', '투입량', '판정', '확인여부', '특이사항', '입력일시', '고유번호']]
    finally:
        conn.close()

@st.cache_data(show_spinner=False, ttl=600)
def get_ai_predictions(today_str):
    if not HAS_HOLIDAYS:
        return []
    df = load_from_db()
    if df.empty: return []
    
    excluded_products = get_excluded_products()
    predict_data = []
    today_d = datetime.strptime(today_str, "%Y-%m-%d").date()
    df['생산일_dt'] = pd.to_datetime(df['생산일'], errors='coerce').dt.date
    
    for prod, group in df.groupby('제품명'):
        # 단종(예측 제외) 목록에 포함된 제품은 AI 예측에서 숨김
        if prod in excluded_products: 
            continue
            
        unique_dates = sorted(d for d in group['생산일_dt'].dropna().drop_duplicates().tolist() if d <= today_d)
        recent_dates = [d for d in unique_dates if (today_d - d).days <= 120]
        if len(recent_dates) < 2: continue

        last_date = recent_dates[-1]
        intervals = [int(np.busday_count(str(recent_dates[i-1]), str(recent_dates[i]), holidays=KR_HOLIDAYS)) for i in range(1, len(recent_dates))]
        if not intervals: continue
        # 최근 120일의 간격 중앙값을 사용해 오래된 공백이나 한 번의 큰 지연 영향을 줄입니다.
        avg_interval = max(1, float(np.median(intervals)))
        
        next_date_np = np.busday_offset(str(last_date), int(np.floor(avg_interval + 0.5)), roll='forward', holidays=KR_HOLIDAYS)
        next_date = pd.to_datetime(next_date_np).date()
        d_day = int(np.busday_count(str(today_d), str(next_date), holidays=KR_HOLIDAYS))
        
        if d_day < 0: status_str = f"🚨 긴급 ({-d_day}영업일 지남)"
        elif d_day == 0: status_str = "🔥 오늘 생산 권장"
        elif d_day <= 3: status_str = f"⚠️ D-{d_day} (임박)"
        else: status_str = f"✅ D-{d_day} (여유)"
        
        predict_data.append({
            "제품명": prod, "마지막 생산일": last_date.strftime("%Y-%m-%d"),
            "최근 생산 주기(중앙값)": f"약 {int(np.floor(avg_interval + 0.5))}영업일", "다음 예상일": next_date.strftime("%Y-%m-%d"),
            "생산 필요 상태": status_str, "_sort": d_day
        })
    return predict_data

def make_database_backup():
    """Neon의 표 5개를 CSV로 묶어 내려받을 수 있게 합니다."""
    tables = ["color_records", "target_history", "product_notices", "workers", "excluded_products"]
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        conn = get_db_conn()
        try:
            for table in tables:
                df = pd.read_sql_query(
                    f"SELECT * FROM {table}",
                    conn.connection,
                )
                archive.writestr(f"{table}.csv", df.to_csv(index=False).encode("utf-8-sig"))
        finally:
            conn.close()
    memory_file.seek(0)
    return memory_file.getvalue()


def migrate_sqlite_database(uploaded_file):
    """예전 앱에서 내려받은 SQLite 백업을 빈 Neon 데이터베이스로 옮깁니다."""
    temp_path = None
    source = None
    conn = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".db") as temp_file:
            temp_file.write(uploaded_file.getvalue())
            temp_path = temp_file.name

        source = sqlite3.connect(temp_path)
        source.row_factory = sqlite3.Row
        available_tables = {
            row["name"]
            for row in source.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }
        if "color_records" not in available_tables:
            raise ValueError("올린 파일 안에 color_records 기록표가 없습니다.")

        def read_table(table_name):
            if table_name not in available_tables:
                return [], []
            columns = [row["name"] for row in source.execute(f"PRAGMA table_info({table_name})")]
            rows = [dict(row) for row in source.execute(f"SELECT * FROM {table_name}")]
            return columns, rows

        record_columns, records = read_table("color_records")
        required_columns = {"production_date", "equipment", "worker", "product_name", "measured_value"}
        if not required_columns.issubset(set(record_columns)):
            missing = ", ".join(sorted(required_columns - set(record_columns)))
            raise ValueError(f"color_records에 필요한 열이 없습니다: {missing}")

        _, targets = read_table("target_history")
        _, notices = read_table("product_notices")
        _, workers = read_table("workers")
        _, excluded = read_table("excluded_products")

        conn = get_db_conn()
        conn.execute("BEGIN IMMEDIATE")
        # 기존 Neon 기록은 유지하고, 백업과 완전히 같은 기록만 건너뜁니다.
        def normalize_for_compare(value):
            if value is None:
                return None
            if isinstance(value, (int, float)):
                return ("number", round(float(value), 8))
            if isinstance(value, str):
                return value.strip()
            return value

        record_fields = [
            "timestamp", "production_date", "equipment", "worker", "product_name",
            "target_value", "measured_value", "difference", "status", "remarks", "input_amount", "checked",
        ]
        existing_record_rows = conn.execute(
            """SELECT timestamp, production_date, equipment, worker, product_name,
                      target_value, measured_value, difference, status, remarks, input_amount, checked
               FROM color_records"""
        ).fetchall()
        seen_record_signatures = {
            tuple(normalize_for_compare(value) for value in row)
            for row in existing_record_rows
        }
        record_rows = []
        for row in records:
            values = dict(row)
            values.setdefault("timestamp", "")
            values.setdefault("target_value", None)
            values.setdefault("difference", None)
            values.setdefault("status", "")
            values.setdefault("remarks", "")
            values.setdefault("input_amount", "-")
            values.setdefault("checked", 0)
            record_values = tuple(values.get(field) for field in record_fields)
            signature = tuple(normalize_for_compare(value) for value in record_values)
            if signature not in seen_record_signatures:
                record_rows.append(record_values)
                seen_record_signatures.add(signature)
        if record_rows:
            conn.executemany(
                """INSERT INTO color_records
                   (timestamp, production_date, equipment, worker, product_name, target_value,
                    measured_value, difference, status, remarks, input_amount, checked)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                record_rows,
            )

        existing_target_rows = conn.execute(
            "SELECT product_name, target_value, effective_date FROM target_history"
        ).fetchall()
        seen_target_signatures = {
            tuple(normalize_for_compare(value) for value in row)
            for row in existing_target_rows
        }
        target_rows = []
        for row in targets:
            if {"product_name", "target_value", "effective_date"}.issubset(row):
                target_values = (
                    row.get("product_name"),
                    row.get("target_value"),
                    row.get("effective_date"),
                )
                signature = tuple(normalize_for_compare(value) for value in target_values)
                if signature not in seen_target_signatures:
                    target_rows.append(target_values)
                    seen_target_signatures.add(signature)
        if target_rows:
            conn.executemany(
                """INSERT INTO target_history (product_name, target_value, effective_date)
                   VALUES (?, ?, ?)""",
                target_rows,
            )

        notice_rows = [
            (row.get("product_name"), row.get("notice_text"), row.get("start_date"), row.get("end_date"))
            for row in notices if row.get("product_name")
        ]
        if notice_rows:
            conn.executemany(
                """INSERT INTO product_notices (product_name, notice_text, start_date, end_date)
                   VALUES (?, ?, ?, ?)
                   ON CONFLICT (product_name) DO UPDATE SET
                     notice_text=EXCLUDED.notice_text,
                     start_date=EXCLUDED.start_date,
                     end_date=EXCLUDED.end_date""",
                notice_rows,
            )

        worker_rows = [(row.get("name"),) for row in workers if row.get("name")]
        if not worker_rows:
            worker_rows = [(name,) for name in ["윤승태", "문지원", "조성윤", "이태원"]]
        if worker_rows:
            conn.executemany(
                "INSERT INTO workers (name) VALUES (?) ON CONFLICT (name) DO NOTHING",
                worker_rows,
            )

        excluded_rows = [
            (row.get("product_name"),)
            for row in excluded if row.get("product_name")
        ]
        if excluded_rows:
            conn.executemany(
                "INSERT INTO excluded_products (product_name) VALUES (?) ON CONFLICT (product_name) DO NOTHING",
                excluded_rows,
            )

        # 가져온 예전 번호 다음부터 새 번호를 만들도록 자동 번호 값을 맞춥니다.
        conn.execute(
            """SELECT setval(
                   pg_get_serial_sequence('color_records', 'id'),
                   COALESCE((SELECT MAX(id) FROM color_records), 1),
                   EXISTS(SELECT 1 FROM color_records)
               )"""
        )
        conn.execute(
            """SELECT setval(
                   pg_get_serial_sequence('target_history', 'id'),
                   COALESCE((SELECT MAX(id) FROM target_history), 1),
                   EXISTS(SELECT 1 FROM target_history)
               )"""
        )
        conn.commit()
        skipped_duplicates = len(records) - len(record_rows)
        return len(record_rows), len(target_rows), skipped_duplicates
    except Exception:
        if conn is not None:
            conn.rollback()
        raise
    finally:
        if conn is not None:
            conn.close()
        if source is not None:
            source.close()
        if temp_path and Path(temp_path).exists():
            Path(temp_path).unlink()


@st.cache_data
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w: df.to_excel(w, index=False, sheet_name='기록')
    return output.getvalue()

# Neon 연결을 확인하고 필요한 표를 준비합니다.
try:
    init_db()
    check_conn = get_db_conn()
    check_conn.execute("SELECT 1")
    check_conn.close()
except Exception as e:
    st.title("🛑 Neon 데이터베이스에 연결할 수 없습니다")
    st.error("Streamlit Secrets의 Neon 연결 주소와 requirements.txt 설정을 확인해주세요.")
    st.caption(f"{type(e).__name__}: {e}")
    st.stop()

CURRENT_WORKERS = get_all_workers()

@st.cache_data
def load_tgt():
    conn = get_db_conn()
    try:
        rows = conn.execute("SELECT product_name, target_value FROM target_history ORDER BY effective_date ASC, id ASC").fetchall()
        if rows:
            return {r[0]: r[1] for r in rows}
        else:
            try: 
                df = pd.read_excel(EXCEL_FILE, usecols="C:D", header=1).dropna()
                targets = {str(r.iloc[0]).strip(): float(r.iloc[1]) if not pd.isna(r.iloc[1]) else 0.0 for i, r in df.iterrows()}
                for p, v in targets.items():
                    conn.execute("INSERT INTO target_history (product_name, target_value, effective_date) VALUES (?, ?, ?)", (p, v, '2000-01-01'))
                conn.commit()
                return targets
            except: 
                return {"(데이터 없음)": 0.0}
    finally:
        conn.close()

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
        with st.expander("💾 PostgreSQL 백업", expanded=False):
            st.caption("백업은 버튼을 눌렀을 때 생성합니다. 기록이 많으면 잠시 걸릴 수 있습니다.")
            if st.button("백업 파일 준비 / 새로 만들기", key="admin_btn_prepare_backup"):
                try:
                    with st.spinner("Neon 기록을 백업 파일로 만들고 있습니다. 잠시 기다려주세요."):
                        st.session_state["neon_backup_bytes"] = make_database_backup()
                    st.success("백업 파일을 준비했습니다. 아래 버튼으로 컴퓨터에 저장하세요.")
                except Exception as e:
                    st.error(f"백업 파일을 만들지 못했습니다: {e}")

            if st.session_state.get("neon_backup_bytes"):
                st.download_button(
                    "⬇️ 준비된 백업 파일 다운로드",
                    st.session_state["neon_backup_bytes"],
                    "color_management_backup.zip",
                    "application/zip",
                    key="admin_btn_download_backup",
                )
        
        t1, t2, t3, t4, t5, t6, t7, t8, t9, t10 = st.tabs(["🔍 금일 확인", "📝 수정/삭제", "📂 과거기록 업로드", "📅 제품기준 적용", "📢 공지", "⏳ 미생산", "👥 통계", "🧑‍🔧 데이터 정화", "🔮 AI 예측", "🗃️ SQLite 이전"])
        
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
                    conn = get_db_conn()
                    try:
                        row = conn.execute(
                            """SELECT product_name, target_value, production_date, equipment, worker,
                                      measured_value, remarks, input_amount, COALESCE(checked, 0)
                               FROM color_records WHERE id=?""",
                            (tid,),
                        ).fetchone()
                    finally:
                        conn.close()
                        
                    if row:
                        try: def_date = datetime.strptime(row[2], "%Y-%m-%d").date()
                        except: def_date = get_now_kst().date()
                        npd = st.date_input("생산일", value=def_date, key="admin_date_edit").strftime("%Y-%m-%d")
                        
                        opts = list(TARGET_DATA.keys())
                        if row[0] not in opts: opts.append(row[0])
                        
                        nprod = st.selectbox("제품", opts, index=opts.index(row[0]), key="admin_sel_prod")
                        neq = st.selectbox("설비", EQUIPMENT_LIST, index=EQUIPMENT_LIST.index(row[3]) if row[3] in EQUIPMENT_LIST else 0, key="admin_sel_equip")
                        namt = st.selectbox("투입량", ["1.35kg","2.5kg","3.75kg"], index=["1.35kg","2.5kg","3.75kg"].index(row[7]) if row[7] in ["1.35kg","2.5kg","3.75kg"] else 0, key="admin_sel_amt") if "버닝" in neq else ("12kg" if "태환" in neq else "25kg" if "프로밧" in neq else "60kg" if "60" in neq else "125kg" if "120" in neq else "-")
                        nw = st.selectbox("작업자", CURRENT_WORKERS, index=CURRENT_WORKERS.index(row[4]) if row[4] in CURRENT_WORKERS else 0, key="admin_sel_worker")
                        nm = st.number_input("측정", value=float(row[5]), step=0.1, key="admin_num_meas")
                        nrm = st.text_input("특이사항", value=row[6], key="admin_txt_rmk")
                        
                        if st.button("✏️ 수정 완료", key="admin_btn_edit_record"):
                            tgt = get_historical_target(nprod, npd)
                            if tgt is None:
                                st.error("선택한 날짜에 적용되는 기준값이 없어 수정할 수 없습니다. 기준 이력을 먼저 등록해주세요.")
                            else:
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
                        conn = get_db_conn()
                        try:
                            for _, r in df_up.iterrows():
                                if str(r['측정색도']).strip() in ['-', '', 'nan', 'None']: continue
                                try: meas = float(str(r['측정색도']).strip())
                                except: continue
                                
                                p_dt = safe_date_parse(r['생산일'])
                                if not p_dt: continue 
                                
                                pd_name, eq, wk = str(r['제품명']).strip(), str(r['생산설비']).strip(), str(r['작업자']).strip()
                                am = str(r.get('투입량', '')).strip() if '버닝' in eq.lower() else ("12kg" if "태환" in eq else "25kg" if "프로밧" in eq else "60kg" if "60" in eq else "125kg" if "120" in eq else "-")
                                rm = str(r.get('특이사항', '')).strip()
                                tgt = get_historical_target(pd_name, p_dt)
                                diff = round(meas - tgt, 1) if tgt is not None else None
                                stat = ("합격 🟢" if abs(diff)<=2.0 else "불합격 🔴") if diff is not None else "기준 없음 ⚪"
                                
                                # DB 입력
                                conn.execute('INSERT INTO color_records (timestamp, production_date, equipment, worker, product_name, target_value, measured_value, difference, status, remarks, input_amount) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', (get_now_kst().strftime("%Y-%m-%d %H:%M:%S"), p_dt, eq, wk, pd_name, tgt, meas, diff, stat, rm if rm not in ['nan','None'] else '', am))
                                
                                # [자동 해제 로직] 과거 기록을 엑셀로 추가해도 시스템이 인지하여 단종 목록에서 자동 삭제
                                conn.execute('DELETE FROM excluded_products WHERE product_name = ?', (pd_name,))
                                
                            conn.commit()
                            st.cache_data.clear(); st.session_state['show_toast'] = "과거 기록 업로드 성공!"; st.rerun()
                        finally:
                            conn.close()
                    else:
                        st.error("엑셀에 필수 열('생산일', '제품명', '생산설비', '작업자', '측정색도')이 부족합니다.")
                except Exception as e: st.error(f"오류: {e}")
            
            st.markdown("---")
            st.error("🛠️ **DB 고유번호 꼬임 해결 (초기화 및 재정렬)**")
            st.caption("과거 데이터를 나중에 업로드하여 고유번호(순서)가 날짜와 맞지 않게 꼬였을 때, 아래 버튼을 누르면 생산일자 순으로 고유번호를 깔끔하게 재정렬합니다.")
            if st.button("🧹 고유번호 날짜순 전면 재정렬", type="primary", key="admin_btn_reindex"):
                conn = get_db_conn()
                try:
                    conn.execute("BEGIN IMMEDIATE")
                    old_ids = [r[0] for r in conn.execute("SELECT id FROM color_records ORDER BY production_date ASC, id ASC").fetchall()]
                    if not old_ids:
                        conn.rollback()
                        st.info("데이터가 없습니다.")
                    else:
                        # 기존 테이블을 삭제하지 않고, 한 트랜잭션 안에서 번호만 바꿉니다.
                        conn.execute("UPDATE color_records SET id = -id")
                        conn.executemany(
                            "UPDATE color_records SET id = ? WHERE id = ?",
                            [(new_id, -old_id) for new_id, old_id in enumerate(old_ids, start=1)]
                        )
                        # 번호를 다시 매긴 뒤, 다음 새 기록이 겹치지 않도록 자동 번호도 맞춥니다.
                        conn.execute(
                            """SELECT setval(
                                   pg_get_serial_sequence('color_records', 'id'),
                                   COALESCE((SELECT MAX(id) FROM color_records), 1),
                                   EXISTS(SELECT 1 FROM color_records)
                               )"""
                        )
                        conn.commit()
                        st.cache_data.clear()
                        st.session_state['show_toast'] = "고유번호 전면 재정렬 완료! 순서가 정상화되었습니다."
                        st.rerun()
                except Exception as e:
                    conn.rollback()
                    st.error(f"재정렬 중 오류 발생: {e}")
                finally:
                    conn.close()
        
        with t4:
            st.info("제품별 기준 색도 엑셀 파일을 업로드하여 시스템에 즉시 적용합니다.")
            h_up = st.file_uploader("제품 기준값/이력 엑셀 업로드", type=['xlsx','xls'], key="admin_file_history")
            if h_up and st.button("🚀 제품 기준값 전체 적용", key="admin_btn_upload_history"):
                try:
                    df_h = pd.read_excel(h_up)
                    if all(c in df_h.columns for c in ['제품명','적용시작일','기준색도']):
                        prepared_rows = []
                        validation_errors = []
                        seen_keys = set()
                        for row_num, (_, r) in enumerate(df_h.iterrows(), start=2):
                            product = str(r['제품명']).strip()
                            if product.lower() in ('', 'nan', 'none'):
                                validation_errors.append(f"{row_num}행: 제품명이 비어 있습니다.")
                                continue

                            date_text = safe_date_parse(r['적용시작일'])
                            try:
                                effective_date = datetime.strptime(date_text, "%Y-%m-%d").strftime("%Y-%m-%d")
                            except (TypeError, ValueError):
                                validation_errors.append(f"{row_num}행: 적용시작일을 날짜로 읽을 수 없습니다.")
                                continue

                            try:
                                target = float(r['기준색도'])
                                if not np.isfinite(target):
                                    raise ValueError("숫자가 아닙니다")
                            except (TypeError, ValueError):
                                validation_errors.append(f"{row_num}행: 기준색도는 유효한 숫자여야 합니다.")
                                continue

                            key = (product, effective_date)
                            if key in seen_keys:
                                validation_errors.append(f"{row_num}행: 같은 제품과 적용일이 중복됩니다.")
                                continue
                            seen_keys.add(key)
                            prepared_rows.append((product, target, effective_date))

                        if not prepared_rows:
                            validation_errors.append("적용할 기준값 행이 없습니다.")

                        if validation_errors:
                            st.error("기준값을 적용하지 않았습니다. 파일을 수정한 뒤 다시 올려주세요.\n\n" + "\n".join(validation_errors[:10]))
                        else:
                            conn = get_db_conn()
                            try:
                                conn.execute("BEGIN IMMEDIATE")
                                conn.execute("DELETE FROM target_history")
                                conn.executemany(
                                    "INSERT INTO target_history (product_name, target_value, effective_date) VALUES (?, ?, ?)",
                                    prepared_rows
                                )
                                conn.commit()
                            except Exception:
                                conn.rollback()
                                raise
                            finally:
                                conn.close()
                            st.cache_data.clear()
                            st.session_state['show_toast'] = "제품 기준값이 시스템 전체에 즉시 적용되었습니다!"
                            st.rerun()
                    else:
                        st.error("엑셀 파일에 '제품명', '적용시작일', '기준색도' 열이 포함되어 있어야 합니다.")
                except Exception as e:
                    st.error(f"적용 중 오류 발생: {e}")
                    
        with t5:
            if ACTIVE_NOTICES:
                st.markdown("#### 📋 현재 적용 중인 공지 목록")
                notice_list = [{"제품명": k, "공지 내용": v} for k, v in ACTIVE_NOTICES.items()]
                st.dataframe(pd.DataFrame(notice_list), use_container_width=True, hide_index=True)
                st.markdown("---")
            else:
                st.info("현재 활성화된 공지가 없습니다.")
            
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
            latest_by_product = {}
            if not history_df.empty:
                latest_rows = history_df.drop_duplicates(subset=["제품명"], keep="first")
                latest_by_product = {
                    str(row["제품명"]): row
                    for _, row in latest_rows.iterrows()
                }
            for p in TARGET_DATA.keys():
                latest_row = latest_by_product.get(p)
                if latest_row is not None:
                    try:
                        last_date = safe_date_parse(latest_row["생산일"])
                        d = (td - datetime.strptime(last_date, "%Y-%m-%d").date()).days
                        if d >= 120: inact.append({"제품명":p, "최종 생산":last_date, "경과":f"{d}일"})
                    except: pass
                else: inact.append({"제품명":p, "최종 생산":"없음", "경과":"이력 없음"})
            st.dataframe(pd.DataFrame(inact), use_container_width=True, hide_index=True)
            
        with t7:
            if not history_df.empty:
                ws = []
                for nm, grp in history_df.groupby('작업자'):
                    if nm == '미입력(과거기록)':
                        disp_nm = nm
                        sort_prio = 1
                    elif nm not in CURRENT_WORKERS:
                        disp_nm = f"{nm} (퇴사)"
                        sort_prio = 2
                    else:
                        disp_nm = nm
                        sort_prio = 0
                        
                    tc = len(grp)
                    pc = int(grp['판정'].eq("합격 🟢").sum())
                    fc = int(grp['판정'].eq("불합격 🔴").sum())
                    unknown = tc - pc - fc
                    known_count = pc + fc
                    ws.append({"sort_prio": sort_prio, "작업자": disp_nm, "총":tc, "합격":pc, "불합격":fc, "미판정":unknown, "불량률(%)":fc/known_count*100 if known_count>0 else 0, "오차(절대)":grp['오차'].abs().mean()})
                
                stat_df = pd.DataFrame(ws).sort_values(by=["sort_prio", "총"], ascending=[True, False]).drop(columns=["sort_prio"])
                st.dataframe(stat_df.style.format({"불량률(%)":"{:.1f}%", "오차(절대)":"{:.2f}"}), hide_index=True)

        with t8:
            st.info("작업자 관리 및 DB 일괄 정화 도구입니다.")
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
                conn = get_db_conn()
                try:
                    conn.execute("UPDATE color_records SET worker = TRIM(worker), product_name = TRIM(product_name)")
                    conn.execute("UPDATE color_records SET equipment = '버닝' WHERE REPLACE(equipment, ' ', '') LIKE '%버닝%'")
                    conn.execute("UPDATE color_records SET equipment = '태환 12kg' WHERE REPLACE(equipment, ' ', '') LIKE '%태환%'")
                    conn.execute("UPDATE color_records SET equipment = '프로밧 25kg' WHERE REPLACE(equipment, ' ', '') LIKE '%프로밧%'")
                    conn.execute("UPDATE color_records SET equipment = '뷸러 60kg' WHERE REPLACE(equipment, ' ', '') LIKE '%60%'")
                    conn.execute("UPDATE color_records SET equipment = '뷸러 120kg' WHERE REPLACE(equipment, ' ', '') LIKE '%120%'")
                    conn.commit()
                finally:
                    conn.close()
                st.cache_data.clear(); st.session_state['show_toast'] = "데이터 명칭 100% 통일화 완료!"; st.rerun()
                
            st.markdown("---")
            st.error("🛠️ **뷸러 120kg 투입량 일괄 수정 (120kg → 125kg)**")
            st.caption("과거에 '120kg'으로 잘못 기록된 뷸러 120kg 설비의 투입량을 실제에 맞게 '125kg'으로 일괄 변경합니다.")
            if st.button("🚀 뷸러 120kg 투입량 125kg으로 일괄 변경", type="primary", use_container_width=True, key="admin_btn_update_125"):
                conn = get_db_conn()
                try:
                    conn.execute("UPDATE color_records SET input_amount = '125kg' WHERE equipment LIKE '%120%' AND input_amount = '120kg'")
                    conn.commit()
                finally:
                    conn.close()
                st.cache_data.clear()
                st.session_state['show_toast'] = "125kg 일괄 변경 완료!"
                st.rerun()
                
        with t10:
            st.info("예전 앱의 SQLite 백업 파일을 Neon PostgreSQL로 한 번 옮기는 메뉴입니다.")
            st.warning(
                "먼저 기존 앱의 관리자 메뉴에서 color_management.db 백업을 내려받아 두세요. "
                "새 앱을 배포한 뒤에는 아래에서 그 .db 파일을 올려 이관합니다."
            )
            st.caption(
                "Neon의 기존 기록은 보존합니다. 백업과 내용이 완전히 같은 기록은 다시 넣지 않고, "
                "새로 옮기는 기록에는 Neon에서 새 고유번호를 부여합니다."
            )
            count_conn = get_db_conn()
            try:
                current_neon_count = count_conn.execute("SELECT COUNT(*) FROM color_records").fetchone()[0]
            finally:
                count_conn.close()
            st.info(f"현재 Neon에 저장된 생산 기록: {current_neon_count:,}건")
            sqlite_upload = st.file_uploader(
                "예전 앱에서 받은 SQLite 백업 파일(.db)",
                type=["db", "sqlite", "sqlite3"],
                key="admin_sqlite_migration_file",
            )
            st.caption("먼저 위에서 .db 파일을 선택하세요. 파일을 선택하기 전에는 아래 버튼이 비활성화됩니다.")
            if st.button(
                "SQLite 기록을 Neon으로 옮기기",
                type="primary",
                key="admin_btn_sqlite_migration",
                disabled=sqlite_upload is None,
            ):
                try:
                    with st.spinner("기록을 옮기는 중입니다. 창을 닫지 마세요."):
                        moved_records, moved_targets, skipped_duplicates = migrate_sqlite_database(sqlite_upload)
                    st.cache_data.clear()
                    st.session_state["show_toast"] = (
                        f"이관 완료: 새 기록 {moved_records:,}건, 기준값 이력 {moved_targets:,}건, "
                        f"중복 건너뜀 {skipped_duplicates:,}건"
                    )
                    st.rerun()
                except Exception as e:
                    st.error(f"이관하지 못했습니다: {e}")

        with t9:
            st.info("최근 4개월(120일) 이내에 2회 이상 생산된 제품의 생산 간격 중간값을 이용한 영업일 예측")
            
            with st.expander("🚫 단종/생산종료 제품 예측 제외 관리", expanded=False):
                st.caption("계약 종료 등으로 더 이상 생산하지 않는 제품을 AI 예측 목록에서 숨깁니다. (생산 재개 시 자동으로 복구됩니다.)")
                ex_prods = get_excluded_products()
                # 모든 제품 목록 확보 (타겟 데이터 + 과거 기록)
                all_prods = sorted(list(set(TARGET_DATA.keys()).union(set(history_df['제품명'].unique()))))
                
                c_ex1, c_ex2 = st.columns(2)
                with c_ex1:
                    to_exclude = st.selectbox("숨길 제품 선택", [p for p in all_prods if p not in ex_prods], key="sel_ex_prod")
                    if st.button("➕ 수동 제외 목록에 추가", key="btn_add_ex"):
                        if to_exclude:
                            add_excluded_product(to_exclude)
                            st.cache_data.clear()
                            st.session_state['show_toast'] = f"{to_exclude} 제품이 예측에서 제외되었습니다."
                            st.rerun()
                with c_ex2:
                    if ex_prods:
                        to_restore = st.selectbox("제외된 제품 목록 (복구 시 선택)", ex_prods, key="sel_res_prod")
                        if st.button("🔄 수동으로 제외 해제", key="btn_remove_ex"):
                            if to_restore:
                                remove_excluded_product(to_restore)
                                st.cache_data.clear()
                                st.session_state['show_toast'] = f"{to_restore} 제품이 예측에 복구되었습니다."
                                st.rerun()
                    else:
                        st.info("현재 예측 제외된 제품이 없습니다.")
            
            if not HAS_HOLIDAYS:
                st.warning("한국 공휴일 정보를 읽을 수 없어 영업일 기준 예측을 표시하지 않습니다. 'holidays' 패키지를 설치해주세요.")
            else:
                pred_data = get_ai_predictions(today_str_kst)
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
                    st.success("데이터가 부족하거나 모든 제품이 제외되어 예측할 수 없습니다.")
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
                input_amount_val = "12kg" if "태환" in equip_clean else "25kg" if "프로밧" in equip_clean else "60kg" if "60" in equip_clean else "125kg" if "120" in equip_clean else "-"
                st.text_input("투입량 (고정)", input_amount_val, disabled=True, key="main_amt_txt")
                
        col_p1, col_p2 = st.columns([2, 1])
        with col_p1:
            selected_product = st.selectbox(
                "🔍 제품명 검색 및 선택", 
                list(TARGET_DATA.keys()), 
                index=None, 
                placeholder="제품명을 클릭하여 검색하세요 (우측 'X' 버튼으로 즉시 지우기)", 
                key="main_prod"
            )
            if selected_product and ACTIVE_NOTICES.get(selected_product): 
                st.warning(f"📢 **전달사항:** {ACTIVE_NOTICES[selected_product]}")
                
        with col_p2:
            target_value = get_historical_target(selected_product, prod_date_str) if selected_product else None
            if target_value is None:
                st.warning(f"📌 {prod_date_str}에 적용되는 기준값 이력이 없습니다.")
            else:
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
        with cs8: measured_value = st.number_input("측정 색도 입력", value=float(target_value) if target_value is not None else 0.0, step=0.1, key="main_meas")
        with cs9: remarks_input = st.text_input("특이사항 (선택사항)", placeholder="메모 입력", key="main_rmk")
        with cs10:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("데이터 등록하기", type="primary", use_container_width=True, key="main_btn_save"):
                if not selected_product: st.warning("⚠️ 제품명을 먼저 선택해주세요!")
                elif not worker_name: st.warning("⚠️ 작업자 오류!")
                elif target_value is None: st.error("⚠️ 선택한 생산일에 적용되는 기준값이 없어 등록할 수 없습니다.")
                elif check_recent_duplicate(prod_date_str, selected_equipment, selected_product, measured_value): st.error("⚠️ 중복 데이터!")
                else:
                    diff = round(measured_value - target_value, 1)
                    saved = save_to_db(prod_date_str, selected_equipment, worker_name, selected_product, target_value, measured_value, diff, "합격 🟢" if abs(diff)<=2.0 else "불합격 🔴", remarks_input, input_amount_val)
                    if not saved:
                        st.error("⚠️ 같은 기록이 방금 등록되었습니다.")
                    else:
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
                if qt is None:
                    st.warning("오늘 적용되는 기준값 이력이 없어 빠른 등록을 할 수 없습니다.")
                else:
                    with cq2: st.text_input("기준", f"{float(qt):.1f}", disabled=True, key="quick_tgt")
                    with cq3: qm = st.number_input("측정값", value=float(qt), step=0.1, key="quick_meas")
                    with cq4:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("🚀 1초 빠른 등록", type="primary", use_container_width=True, key="quick_btn"):
                            if check_recent_duplicate(today_str_kst, qe, qp, qm): st.error("⚠️ 중복")
                            else:
                                diff = round(qm - qt, 1)
                                saved = save_to_db(today_str_kst, qe, qw, qp, qt, qm, diff, "합격 🟢" if abs(diff)<=2.0 else "불합격 🔴", "", qa)
                                if not saved:
                                    st.error("⚠️ 같은 기록이 방금 등록되었습니다.")
                                else:
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
        # [정렬 정상화] 설비 묶음(버닝->태환...) + 시간 역순(최신 생산일 먼저) + 당일 내 먼저 생산한 제품 그룹 + 번호 정순
        eq_map = {'버닝': 0, '태환12kg': 1, '프로밧25kg': 2, '뷸러60kg': 3, '뷸러120kg': 4}
        ddf['s'] = ddf['생산설비'].astype(str).str.replace(" ", "").str.lower().map(lambda x: eq_map.get(x, 5))
        
        # [신규] 당일(생산일) & 설비 내에서 해당 제품이 '가장 먼저 등록된 고유번호'를 찾아 같은 제품끼리 묶어줌
        ddf['prod_first_id'] = ddf.groupby(['s', '생산일', '제품명'])['고유번호'].transform('min')
        
        # 1.설비순 -> 2.최신날짜순 -> 3.먼저 입력된 제품그룹 -> 4.해당 제품 내 고유번호순(실제 시간순)
        ddf = ddf.sort_values(by=['s', '생산일', 'prod_first_id', '고유번호'], ascending=[True, False, True, True])
        ddf = ddf.drop(columns=['s', 'prod_first_id'])

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

def hl_stat(s):
    colors = []
    for value in s:
        label = str(value)
        if '불합격' in label:
            colors.append('color: white; background-color: #E74C3C; font-weight: bold;')
        elif label.startswith('합격'):
            colors.append('color: #27AE60; font-weight: bold;')
        else:
            colors.append('color: #555; background-color: #F2F3F4;')
    return colors
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
