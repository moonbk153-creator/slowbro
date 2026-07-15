import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
import io 
import pytz
import numpy as np

# 한국 공휴일 라이브러리 예외 처리
try:
    import holidays
    HAS_HOLIDAYS = True
except ImportError:
    HAS_HOLIDAYS = False

try: import holidays; HAS_HOLIDAYS = True
except: HAS_HOLIDAYS = False
try: from streamlit_autorefresh import st_autorefresh
except ImportError: st_autorefresh = None
except: st_autorefresh = None

KST = pytz.timezone('Asia/Seoul')
st.set_page_config(page_title="색도 관리 시스템", layout="wide")
@@ -696,18 +692,21 @@
if not ddf.empty:
mdf = ddf.drop(columns=['확인여부'], errors='ignore')

    # 1. 디자인이 적용된 Styler 객체 생성
    sdf = mdf.style.format({"측정색도":"{:.1f}", "기준색도":"{:.1f}", "오차":"{:.1f}"}, na_rep="-") \
                   .apply(hl_eq, subset=['생산설비']) \
                   .apply(hl_stat, subset=['판정']) \
                   .set_properties(subset=['특이사항'], **{'background-color': '#E8DAEF', 'color': 'black', 'font-weight': 'bold'}) \
                   .set_properties(subset=['제품명'], **{'font-weight': 'bold'})
    
    # 2. st.dataframe 대신 st.table을 사용하거나, 
    # 스타일 적용된 데이터는 st.write를 통해 표시하는 것이 오류 방지에 훨씬 강력합니다.
    st.write(sdf) 
    # [수정된 출력부] Styler 객체 없이 데이터 프레임을 그대로 출력합니다.
    # 스타일 적용 대신 가독성을 위해 판정 컬럼을 확인하기 쉽게 표시합니다.
    st.dataframe(
        mdf, 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "판정": st.column_config.TextColumn("판정", help="🟢합격/🔴불합격"),
            "측정색도": st.column_config.NumberColumn("측정색도", format="%.1f"),
            "기준색도": st.column_config.NumberColumn("기준색도", format="%.1f"),
            "오차": st.column_config.NumberColumn("오차", format="%.1f"),
        }
    )

fn = f"색도측정_{today_str_kst if dm=='오늘' else fd_str if dm=='특정 일자' else '전체'}.xlsx"
st.download_button("📥 엑셀 다운로드", to_excel(mdf), fn, key="btn_download_excel")
else: 
st.info("🔍 일치하는 기록이 없습니다.")
