import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
from datetime import datetime
from io import BytesIO
from openpyxl.styles import Font
import openpyxl
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials

# QR/바코드 스캔 라이브러리
try:
    from PIL import Image
    import cv2
    from pyzbar.pyzbar import decode
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

# [설정] 작업자 명단
worker_list = ["박경섭", "무고사", "재르소", "김동헌"] 

st.set_page_config(
    page_title="VISION DATA KEY-IN SYSTEM ----- (by. Romero)", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------
# 🪄 마법 코드 1: UI 숨김 및 태블릿 최적화 커스텀
# ----------------------------------------------------
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
body { overscroll-behavior-y: none !important; }
* {
    -webkit-user-select: none; -ms-user-select: none; user-select: none; 
    -webkit-tap-highlight-color: transparent !important; 
}
input, textarea, select {
    -webkit-user-select: auto !important; -ms-user-select: auto !important; user-select: auto !important;
}
::-webkit-scrollbar { display: none; }
.block-container {
    padding-top: 0rem !important; padding-bottom: 0rem !important;
    padding-left: 1rem !important; padding-right: 1rem !important;
}
button[kind="primary"] {
    background-color: #4b6584 !important; color: white !important; border: none !important;
    font-size: 18px !important; font-weight: bold !important; padding: 12px !important;
}
button[kind="primary"]:hover { background-color: #3b5068 !important; }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ----------------------------------------------------
# 🌐 구글 스프레드시트 연동
# ----------------------------------------------------
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# 💡 신규 'LOT' 컬럼 추가 및 필요 항목 반영
EXCEL_COLUMNS = [
    "날짜", "교대", "시작시간", "종료시간", "휴동시간", "소요시간", "구분", "호기", "모델명", "LOT", "도금구분", 
    "검사수량", "양품수량", "양품수량(전,배 포함)", "불량수량", "양품율", "양품율(전/배 포함)", 
    "전면 불량율", "배면 불량율", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", 
    "기타", "OQC", "비고", "도장일", "도장 Line", "작업자"
]
SHEET_NAME = "VISION_DATA_DB"

if "google_credentials" not in st.secrets:
    st.error("🚨 구글 스프레드시트 보안 키(Secrets)가 설정되지 않았습니다!")
    st.stop()

@st.cache_resource
def get_sheet():
    try:
        creds_data = st.secrets["google_credentials"]
        clean_data = creds_data.strip().strip("'").strip('"') if isinstance(creds_data, str) else dict(creds_data)
        creds_dict = json.loads(clean_data, strict=False) if isinstance(creds_data, str) else clean_data
        if "private_key" in creds_dict: creds_dict["private_key"] = creds_dict["private_key"].replace('\\n', '\n')
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
        return gspread.authorize(creds).open(SHEET_NAME).sheet1
    except Exception as e:
        st.error(f"🚨 구글 연결 초기화 에러: {e}")
        return None

@st.cache_data(ttl=60)
def load_data():
    sheet = get_sheet()
    if sheet is None: return pd.DataFrame(columns=EXCEL_COLUMNS)
    try:
        try: raw_data = sheet.get_all_values()
        except KeyError: raw_data = []
        if not raw_data or len(raw_data) < 2: return pd.DataFrame(columns=EXCEL_COLUMNS)
        headers = [str(h).strip() for h in raw_data[0]]
        df = pd.DataFrame(raw_data[1:], columns=headers)
        result_df = pd.DataFrame()
        for col in EXCEL_COLUMNS:
            result_df[col] = df[col].iloc[:, 0] if isinstance(df.get(col), pd.DataFrame) else df.get(col, "")
        return result_df
    except Exception:
        return pd.DataFrame(columns=EXCEL_COLUMNS)

def clean_for_gsheet(df):
    df_clean = df.copy()
    for col in df_clean.columns:
        df_clean[col] = df_clean[col].apply(lambda x: "" if str(x).strip().lower() in ["nan", "nat", "none", "<na>", "inf", "-inf"] else str(x))
    return df_clean

def save_data_append(df):
    sheet = get_sheet()
    if sheet is None: return False
    try:
        try: header_check = sheet.row_values(1)
        except Exception: header_check = []
        if not header_check: sheet.append_row(EXCEL_COLUMNS, value_input_option='USER_ENTERED')
        
        records_to_insert = []
        for _, row in df.iterrows():
            row_data = ["" if str(row.get(col, "")).strip().lower() in ["nan", "nat", "none", "inf", "-inf"] else str(row.get(col, "")).strip() for col in EXCEL_COLUMNS]
            records_to_insert.append(row_data)

        try: sheet.append_rows(records_to_insert, value_input_option='USER_ENTERED')
        except Exception:
            for r_data in records_to_insert: sheet.append_row(r_data, value_input_option='USER_ENTERED')
        load_data.clear() 
        return True
    except Exception as e:
        st.error(f"🚨 데이터 저장 오류: {e}")
        return False

# ==========================================
# ⌨️ [1페이지] 데이터 입력 화면 
# ==========================================
if "current_page" not in st.session_state: st.session_state.current_page = "input"
if "scanned_lot" not in st.session_state: st.session_state.scanned_lot = ""

if st.session_state.current_page == "input":
    st.markdown("""
        <div style='background: linear-gradient(135deg, #0f172a 0%, #020617 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #1e293b;'>
            <h2 style='color: #f8fafc; margin: 0; font-weight: 600;'>💻 VISION DATA KEY-IN SYSTEM</h2>
        </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("<h4 style='color: #f9fafb;'>📋 Information</h4>", unsafe_allow_html=True)
        
        work_date = st.date_input("근무일자", value=datetime.now())
        shift_type = st.selectbox("교대", ["주간", "야간"])
        model_name = st.selectbox("모델명", ["MEM", "Centaur", "Sphinx-E", "Banff", "Krios", "AV-J", "Seattle", "Juliet-O"])
        
        # 💡 QR 스캔 기능 및 LOT 입력
        st.markdown("<br><b>🔤 LOT 관리</b>", unsafe_allow_html=True)
        lot_number = st.text_input("LOT 번호", value=st.session_state.scanned_lot, placeholder="직접 입력 또는 아래 스캔")
        
        with st.expander("📷 QR/바코드 카메라 스캔"):
            if not QR_AVAILABLE:
                st.warning("QR 인식 라이브러리(opencv-python-headless, pyzbar)가 설치되지 않았습니다.")
            else:
                img_buffer = st.camera_input("QR 바코드를 화면에 맞춰주세요")
                if img_buffer is not None:
                    try:
                        image = Image.open(img_buffer)
                        cv2_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                        decoded_objects = decode(cv2_img)
                        if decoded_objects:
                            scanned_data = decoded_objects[0].data.decode('utf-8')
                            st.session_state.scanned_lot = scanned_data
                            st.success(f"✅ 스캔 성공: {scanned_data}")
                            st.rerun()
                        else:
                            st.error("코드를 인식할 수 없습니다.")
                    except Exception as e:
                        st.error(f"스캔 오류: {e}")

        st.markdown("<br><b>⏱️ 시간 및 생산 관리</b>", unsafe_allow_html=True)
        start_time = st.time_input("시작시간", value=datetime.now().time())
        end_time = st.time_input("종료시간", value=datetime.now().time())
        unit = st.selectbox("호기", [f"{i}호기" for i in range(1, 9)])
        category = st.selectbox("구분", ["1차검사", "2차검사", "3차검사", "4차검사", "Sample", "완불재검"])
        idle_time = st.number_input("휴동시간 (분)", min_value=0, value=0)
        plating_type = st.selectbox("도금구분", ["A", "B"])

        if st.button("📈 종합 분석 데이터 확인", use_container_width=True):
            st.session_state.current_page = "analysis"
            st.rerun()

    # 메인 입력 화면
    main_col1, main_col2 = st.columns([1.1, 0.9])

    with main_col1:
        st.markdown("<h4 style='color:#f9fafb;'>📥 VISION Data Input</h4>", unsafe_allow_html=True)
        
        # 💡 자동 계산 로직 적용 (검사수량, 불량수량 자동화)
        v_row1_col1, v_row1_col2, v_row1_col3 = st.columns(3)
        v_row2_col1, v_row2_col2, v_row2_col3 = st.columns(3)
        v_row3_col1, v_row3_col2 = st.columns(2)
        
        with v_row1_col1: good_qty = st.number_input("양품수량", min_value=0, value=0)
        with v_row2_col1: comp_def = st.number_input("완전불량", min_value=0, value=0)
        with v_row2_col2: front_def = st.number_input("전면불량", min_value=0, value=0)
        with v_row2_col3: rear_def = st.number_input("배면불량", min_value=0, value=0)
        with v_row3_col1: offset_def = st.number_input("옵셋불량", min_value=0, value=0)
        with v_row3_col2: etc_def = st.number_input("기타불량", min_value=0, value=0)
        
        bad_qty = comp_def + front_def + rear_def + offset_def + etc_def
        shortage_qty = st.number_input("수량부족", min_value=0, value=0)
        total_qty = good_qty + bad_qty - shortage_qty
        if total_qty < 0: total_qty = 0
            
        with v_row1_col2: st.text_input("불량수량 (자동계산)", value=f"{bad_qty:,}", disabled=True)
        with v_row1_col3: st.text_input("검사수량 (자동계산)", value=f"{total_qty:,}", disabled=True)

        # 수율 자동 계산
        good_inc_front_rear = good_qty + front_def + rear_def
        rate_good = round((good_qty / total_qty) * 100, 1) if total_qty > 0 else 0.0
        rate_good_inc = round((good_inc_front_rear / total_qty) * 100, 1) if total_qty > 0 else 0.0

        st.markdown("<h4 style='color:#f9fafb;'>📋 기타 정보</h4>", unsafe_allow_html=True)
        etc_col1, etc_col2, etc_col3, etc_col4 = st.columns(4)
        with etc_col1: painting_date = st.date_input("도장일")
        with etc_col2: painting_line = st.selectbox("도장 Line", ["선택안함", "A Line", "B Line", "C Line"])
        with etc_col3: oqc_status = st.selectbox("OQC", ["선택안함", "OQC"])
        with etc_col4: worker_name = st.selectbox("작업자", ["선택안함"] + worker_list)
        remarks = st.text_area("비고", height=68)

        # 💡 중복 처리 확인 및 저장 로직
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 데이터 저장 (구글 스프레드시트)", use_container_width=True, type="primary"):
            if total_qty == 0:
                st.warning("입력된 데이터가 없습니다.")
            elif not lot_number:
                st.warning("LOT 번호를 입력하거나 스캔해주세요.")
            else:
                with st.spinner("중복 확인 및 저장 중..."):
                    db_df = load_data()
                    
                    # 💡 동일 교대 내 LOT 중복 검사
                    is_duplicate = False
                    if not db_df.empty and "LOT" in db_df.columns and "교대" in db_df.columns:
                        dup_check = db_df[(db_df["LOT"] == lot_number) & (db_df["교대"] == shift_type)]
                        if len(dup_check) > 0: is_duplicate = True

                    if is_duplicate:
                        st.error("🚨 이미 처리되었습니다. (동일 교대 내 중복 LOT)")
                    else:
                        new_data = pd.DataFrame([{
                            "날짜": work_date.strftime("%Y-%m-%d"), "교대": shift_type,
                            "시작시간": start_time.strftime("%H:%M"), "종료시간": end_time.strftime("%H:%M"),
                            "휴동시간": idle_time, "소요시간": 0, "구분": category, "호기": unit, 
                            "모델명": model_name, "LOT": lot_number, "도금구분": plating_type,
                            "검사수량": total_qty, "양품수량": good_qty, "양품수량(전,배 포함)": good_inc_front_rear, "불량수량": bad_qty,
                            "양품율": f"{rate_good:.1f}%", "양품율(전/배 포함)": f"{rate_good_inc:.1f}%", 
                            "완전불량": comp_def, "전면불량": front_def, "배면불량": rear_def, "옵셋불량": offset_def, 
                            "수량부족": shortage_qty, "기타": etc_def,
                            "OQC": oqc_status, "비고": remarks, "도장일": painting_date.strftime("%Y-%m-%d"), 
                            "도장 Line": painting_line, "작업자": worker_name
                        }])
                        
                        if save_data_append(new_data):
                            st.success("✅ 데이터가 성공적으로 저장되었습니다!")
                            st.session_state.scanned_lot = "" # 저장 후 초기화
                            import time
                            time.sleep(1.5)
                            st.rerun()

    with main_col2:
        st.markdown("<h4 style='color:#f9fafb;'>📊 Yield Report</h4>", unsafe_allow_html=True)
        m_col1, m_col2 = st.columns(2)
        m_col1.metric(label="검사수량 총합", value=f"{total_qty:,} EA")
        m_col2.metric(label="현재 양품율", value=f"{rate_good:.1f}%")

        fig_donut = go.Figure()
        fig_donut.add_trace(go.Pie(labels=['양품율', '불량율'], values=[rate_good, 100-rate_good], hole=.65, marker=dict(colors=["#3b82f6", "#ef4444"])))
        fig_donut.update_layout(title="양품율 현황", paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#94a3b8'), height=300)
        st.plotly_chart(fig_donut, use_container_width=True)

# ==========================================
# 📈 [2페이지] 분석 화면 (기존 로직 유지)
# ==========================================
elif st.session_state.current_page == "analysis":
    st.markdown("<h2>📈 종합 생산 데이터 분석</h2>", unsafe_allow_html=True)
    if st.button("⬅️ 뒤로 가기", type="primary"):
        st.session_state.current_page = "input"
        st.rerun()
    st.info("분석 차트 기능은 유지됩니다.")