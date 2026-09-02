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

# 💡 QR/바코드 스캔 라이브러리 (파이썬 백엔드 분석용)
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
# 🪄 마법 코드 1: UI 숨김 및 태블릿 앱 최적화
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
    padding-top: 1rem !important; padding-bottom: 1rem !important;
    padding-left: 1.5rem !important; padding-right: 1.5rem !important;
}
button[kind="primary"] {
    background-color: #4b6584 !important; color: white !important; border: none !important;
    font-size: 16px !important; font-weight: bold !important; padding: 10px !important;
}
button[kind="primary"]:hover { background-color: #3b5068 !important; }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

components.html(
    """
    <script>
    if (window.parent && !window.parent.appPluginLoaded) {
        window.parent.appPluginLoaded = true;
        const doc = window.parent.document;
        const head = doc.head;
        const metaTags = [
            { name: "mobile-web-app-capable", content: "yes" },
            { name: "apple-mobile-web-app-capable", content: "yes" },
            { name: "apple-mobile-web-app-status-bar-style", content: "black-translucent" },
            { name: "viewport", content: "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" }
        ];
        metaTags.forEach(tag => {
            let meta = doc.createElement('meta');
            meta.name = tag.name;
            meta.content = tag.content;
            head.appendChild(meta);
        });

        const disableKeyboard = () => {
            if (!doc) return;
            doc.querySelectorAll('div[data-baseweb="select"] input').forEach(el => {
                if (el.getAttribute('inputmode') !== 'none') el.setAttribute('inputmode', 'none');
            });
            doc.querySelectorAll('input').forEach(el => {
                const placeholder = el.getAttribute('placeholder') || '';
                const hasPopup = el.hasAttribute('aria-haspopup');
                if (placeholder.includes('YYYY') || placeholder.includes('HH:MM') || hasPopup) {
                    if (el.getAttribute('inputmode') !== 'none') el.setAttribute('inputmode', 'none');
                }
            });
        };
        const observer = new MutationObserver(() => { disableKeyboard(); });
        if (window.parent.document.body) {
            observer.observe(window.parent.document.body, { childList: true, subtree: true });
        }
        disableKeyboard();
    }
    </script>
    """, height=0, width=0
)

# ----------------------------------------------------
# 🌐 구글 스프레드시트 연동
# ----------------------------------------------------
SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

EXCEL_COLUMNS = [
    "날짜", "교대", "시작시간", "종료시간", "휴동시간", "소요시간", "구분", "호기", 
    "모델명(MI)", "도금구분", "UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", 
    "불량수량", "양품률", "양품율(전/배 포함)", "완전불량률", "전면불량률", "배면불량률", 
    "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타", "육안/OQC", "비고", 
    "도장라인", "도장일", "도장순서", "입고일", "LOT NO.", "CLIP", "BASE", "COVER", 
    "조립기", "월", "작업자"
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

def save_data_overwrite(df):
    sheet = get_sheet()
    if sheet is None: return False
    try:
        try: sheet.clear()
        except Exception: pass
        records_to_insert = [EXCEL_COLUMNS] 
        for _, row in df.iterrows():
            row_data = ["" if str(row.get(col, "")).strip().lower() in ["nan", "nat", "none", "inf", "-inf"] else str(row.get(col, "")).strip() for col in EXCEL_COLUMNS]
            records_to_insert.append(row_data)
        try: sheet.update(values=records_to_insert, range_name='A1', value_input_option='USER_ENTERED')
        except TypeError: sheet.update('A1', records_to_insert, value_input_option='USER_ENTERED')
        load_data.clear()
        return True
    except Exception as e:
        st.error(f"데이터 덮어쓰기 오류: {e}")
        return False

# ----------------------------------------------------
# 🚨 팝업(모달) & Python 이미지 분석 기반 QR 스캐너
# ----------------------------------------------------
@st.dialog("🚨 SBL Warning!")
def show_sbl_warning(defect_type, rate):
    st.markdown(f"### ⚡ [{defect_type}] 불량 제품을 별도 보관 조치 하세요.")
    st.error(f"현재 1차검사 공정의 {defect_type}율이 **{rate:.1f}%** 로 기준치(5.0%) 초과하였습니다.")
    if st.button("✅ 확인 완료 (닫기)", key=f"btn_close_{defect_type}"):
        st.rerun()

@st.dialog("🔒 관리자 인증")
def show_password_dialog():
    st.markdown("분석 데이터를 확인하려면 관리자 비밀번호를 입력하세요.")
    pwd = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
    if st.button("✅ 확인", type="primary", use_container_width=True):
        if pwd == "6233":
            st.session_state.current_page = "analysis"
            st.rerun()
        else:
            st.error("🚨 비밀번호가 일치하지 않습니다.")

# 💡 [해결 완료] 스캔된 값이 텍스트박스 내부 상태(Session State)로 즉시 꽂히도록 수정
@st.dialog("📷 카메라 촬영 및 자동 분석")
def open_camera_qr_scanner():
    if not QR_AVAILABLE:
        st.error("QR 라이브러리(opencv-python-headless, pyzbar)가 설치되지 않았습니다.")
        return
        
    st.info("💡 팁: QR 코드가 화면에 선명하게 보일 때 사진을 찍어주세요.")
    
    tab1, tab2 = st.tabs(["📸 카메라 촬영", "📁 갤러리 앨범"])
    
    target_img = None
    with tab1:
        img_buffer = st.camera_input("카메라 촬영")
        if img_buffer: target_img = img_buffer
        
    with tab2:
        uploaded_img = st.file_uploader("앨범에서 사진 선택", type=['png', 'jpg', 'jpeg'])
        if uploaded_img: target_img = uploaded_img
        
    if target_img:
        with st.spinner("이미지 분석 중..."):
            try:
                image = Image.open(target_img)
                cv2_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                
                gray = cv2.cvtColor(cv2_img, cv2.COLOR_BGR2GRAY)
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                enhanced_gray = clahe.apply(gray)
                
                objs = decode(cv2_img)
                if not objs:
                    objs = decode(enhanced_gray)
                    
                if objs:
                    raw_data = objs[0].data.decode('utf-8')
                    processed_text = raw_data.split('$')[-1] if '$' in raw_data else raw_data
                    
                    st.success(f"✅ 인식 성공: {processed_text}")
                    st.caption(f"원본 바코드: {raw_data}")
                    
                    if st.button("입력창에 적용 및 닫기", type="primary", use_container_width=True):
                        # 💡 [핵심] 변수가 아닌 해당 텍스트박스 객체(Key)에 다이렉트로 값을 덮어씌움
                        st.session_state.lot_input_field = processed_text
                        st.rerun()
                else:
                    st.error("❌ QR 코드를 찾을 수 없습니다. 초점을 맞춰서 다시 촬영해주세요.")
            except Exception as e:
                st.error(f"분석 중 오류 발생: {e}")

# ==========================================
# 🔄 화면 전환 및 상태 관리
# ==========================================
if "current_page" not in st.session_state: st.session_state.current_page = "input"
if "lot_input_field" not in st.session_state: st.session_state.lot_input_field = ""

# ==========================================
# 📈 [2페이지] 종합 분석 데이터 화면 
# ==========================================
def render_analysis_page():
    st.markdown("""
        <div style='background: linear-gradient(135deg, #0f172a 0%, #020617 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1e293b;'>
            <h2 style='color: #f8fafc; margin: 0; font-weight: 600;'>📈 종합 생산 데이터 분석</h2>
        </div>
    """, unsafe_allow_html=True)

    if st.button("⬅️ 뒤로 가기 (데이터 입력 화면으로)", type="primary"):
        st.session_state.current_page = "input"
        st.rerun()
        
    df = load_data().copy()
    if df.empty:
        st.warning("분석할 저장된 데이터가 없습니다.")
        return

    num_cols = ["검사 수량", "양품수량", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    if '시작시간' in df.columns: df['시간대'] = df['시작시간'].str[:13] + "시"
    else: df['시간대'] = df['날짜']

    f_head_col1, f_head_col2 = st.columns([0.95, 0.05])
    with f_head_col1: st.markdown("##### 🔍 상세 분석 필터")
    with f_head_col2:
        with st.popover("🎨"):
            pc1, pc2 = st.columns(2)
            with pc1:
                c_yield1 = st.color_picker("양품율", "#3b82f6")
                c_yield2 = st.color_picker("양품율(포함)", "#10b981")
                c_bad = st.color_picker("불량율", "#ef4444")
                c_comp = st.color_picker("완전불량", "#f43f5e")
            with pc2:
                c_front = st.color_picker("전면불량", "#f97316")
                c_rear = st.color_picker("배면불량", "#eab308")
                c_offset = st.color_picker("옵셋불량", "#a855f7")

    available_dates = sorted(df['날짜'].dropna().unique(), reverse=True)
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1: date_filter_mode = st.radio("📅 분석 기간 설정", ["전체 누적 데이터", "단일 일자 선택", "특정 기간 지정 검색"], horizontal=True)
    with col_opt2: x_axis_mode = st.radio("📊 분석 기준 (X축)", ["일별 (날짜 기준)", "시간별 (시작시간 기준)"], horizontal=True)
    
    if date_filter_mode == "단일 일자 선택":
        selected_date = st.selectbox("분석할 근무일자를 선택하세요", available_dates)
        if selected_date: df = df[df['날짜'] == selected_date]
    elif date_filter_mode == "특정 기간 지정 검색":
        d_col1, d_col2 = st.columns(2)
        try: min_date, max_date = pd.to_datetime(df['날짜']).min().date(), pd.to_datetime(df['날짜']).max().date()
        except: min_date = max_date = datetime.now().date()
        with d_col1: start_date_filter = st.date_input("시작 일자", value=min_date)
        with d_col2: end_date_filter = st.date_input("종료 일자", value=max_date)
        df = df[(df['날짜'] >= start_date_filter.strftime("%Y-%m-%d")) & (df['날짜'] <= end_date_filter.strftime("%Y-%m-%d"))]

    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    with f_col1: selected_model = st.selectbox("🏷️ 모델명", ["전체"] + sorted(df['모델명(MI)'].dropna().astype(str).unique()))
    with f_col2: selected_unit = st.selectbox("⚙️ 호기", ["전체"] + sorted(df['호기'].dropna().astype(str).unique()))
    with f_col3: selected_category = st.selectbox("🔎 검사구분", ["전체"] + sorted(df['구분'].dropna().astype(str).unique()))
    with f_col4: selected_shift = st.selectbox("🌗 교대", ["전체"] + sorted(df['교대'].dropna().astype(str).unique()))

    if selected_model != "전체": df = df[df['모델명(MI)'] == selected_model]
    if selected_unit != "전체": df = df[df['호기'] == selected_unit]
    if selected_category != "전체": df = df[df['구분'] == selected_category]
    if selected_shift != "전체": df = df[df['교대'] == selected_shift]

    if df.empty or df["검사 수량"].sum() == 0:
        st.info("조건에 맞는 데이터가 없습니다.")
        return

    base_col = '시간대' if "시간별" in x_axis_mode else '날짜'
    group_cols = [base_col]
    if selected_model == "전체": group_cols.append('모델명(MI)')

    df_grouped = df.groupby(group_cols)[num_cols].sum().reset_index().sort_values(base_col)
    df_grouped['양품률'] = (df_grouped['양품수량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['양품율(전/배 포함)'] = ((df_grouped['양품수량'] + df_grouped['전면불량'] + df_grouped['배면불량']) / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['불량율'] = (df_grouped['불량수량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['완전불량률'] = (df_grouped['완전불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['전면불량률'] = (df_grouped['전면불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['배면불량률'] = (df_grouped['배면불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['옵셋불량율'] = (df_grouped['옵셋불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)

    is_single_x = len(df_grouped[base_col].unique()) == 1
    def format_labels(s): return s.apply(lambda x: f"{x:.1f}%" if x > 0 else "")

    def create_single_chart(title, metric, base_color):
        fig = go.Figure()
        if selected_model == "전체":
            colors = px.colors.qualitative.Plotly 
            for i, model in enumerate(df_grouped['모델명(MI)'].unique()):
                m_data = df_grouped[df_grouped['모델명(MI)'] == model]
                if is_single_x: fig.add_trace(go.Bar(name=f"{model}-{metric}", x=m_data[base_col], y=m_data[metric], marker_color=colors[i%len(colors)], text=format_labels(m_data[metric]), textposition='auto'))
                else: fig.add_trace(go.Scatter(name=f"{model}-{metric}", x=m_data[base_col], y=m_data[metric], mode='lines+markers+text', marker_color=colors[i%len(colors)], text=format_labels(m_data[metric]), textposition='top center'))
        else:
            if is_single_x: fig.add_trace(go.Bar(name=metric, x=df_grouped[base_col], y=df_grouped[metric], marker_color=base_color, text=format_labels(df_grouped[metric]), textposition='auto'))
            else: fig.add_trace(go.Scatter(name=metric, x=df_grouped[base_col], y=df_grouped[metric], mode='lines+markers+text', marker_color=base_color, text=format_labels(df_grouped[metric]), textposition='top center'))
                
        y_range = [0, df_grouped[metric].max() * 1.2 + 2] if not df_grouped.empty and df_grouped[metric].max() > 0 else [0, 10]
        fig.update_layout(title=title, title_font=dict(color='#f8fafc', size=14), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#f8fafc'), yaxis=dict(gridcolor='#334155', range=y_range), margin=dict(t=50, b=30, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), barmode='group' if is_single_x else None, hovermode="x unified")
        return fig

    st.markdown("---")
    row1_c1, row1_c2, row1_c3 = st.columns(3)
    with row1_c1: st.plotly_chart(create_single_chart("📈 양품율 트렌드", "양품률", c_yield1), use_container_width=True)
    with row1_c2: st.plotly_chart(create_single_chart("📈 양품율(전,배 포함) 트렌드", "양품율(전/배 포함)", c_yield2), use_container_width=True)
    with row1_c3: st.plotly_chart(create_single_chart("📉 불량율 트렌드", "불량율", c_bad), use_container_width=True)
        
    row2_c1, row2_c2, row2_c3, row2_c4 = st.columns(4)
    with row2_c1: st.plotly_chart(create_single_chart("📉 완전불량율 트렌드", "완전불량률", c_comp), use_container_width=True)
    with row2_c2: st.plotly_chart(create_single_chart("📉 전면불량율 트렌드", "전면불량률", c_front), use_container_width=True)
    with row2_c3: st.plotly_chart(create_single_chart("📉 배면불량율 트렌드", "배면불량률", c_rear), use_container_width=True)
    with row2_c4: st.plotly_chart(create_single_chart("📉 옵셋불량율 트렌드", "옵셋불량율", c_offset), use_container_width=True)

# ==========================================
# ⌨️ [1페이지] 데이터 입력 화면 
# ==========================================
if st.session_state.current_page == "input":
    
    if "comp_warned" not in st.session_state: st.session_state.comp_warned = False
    if "front_warned" not in st.session_state: st.session_state.front_warned = False
    if "rear_warned" not in st.session_state: st.session_state.rear_warned = False
    if "offset_warned" not in st.session_state: st.session_state.offset_warned = False

    st.markdown("""
        <div style='background: linear-gradient(135deg, #0f172a 0%, #020617 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1e293b;'>
            <h2 style='color: #f8fafc; margin: 0; font-weight: 600;'>💻 VISION DATA KEY-IN SYSTEM</h2>
        </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.4); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color: #f9fafb; font-weight: 500;'>📋 Information</h4>
            </div>
        """, unsafe_allow_html=True)
        
        default_time = datetime.now().time()
        
        work_date = st.date_input("근무일자", value=datetime.now())
        shift_type = st.selectbox("교대", ["주간", "야간"])
        model_name = st.selectbox("모델명", ["D65S(KRIOS)", "MEM", "Centaur", "Sphinx-E", "Banff", "AV-J", "Seattle", "Juliet-O"])
        
        # 💡 [핵심] 입력창에 key="lot_input_field" 를 부여하여 세션 상태와 완벽 동기화
        st.markdown("<br><b>🔤 LOT NO.</b>", unsafe_allow_html=True)
        lot_number = st.text_input("LOT 입력", placeholder="직접 입력 또는 아래 버튼 스캔", key="lot_input_field")
        
        if st.button("📷 카메라 촬영 및 자동 분석", use_container_width=True):
            open_camera_qr_scanner()

        st.markdown("<br><b>⏱️ 시간 관리</b>", unsafe_allow_html=True)
        start_date = st.date_input("시작일", value=datetime.now())
        start_time = st.time_input("시작시간", value=default_time, key="start_time_key")
        unit = st.selectbox("호기", [f"{i}호기" for i in range(1, 9)])
        
        st.markdown("<br><b>🏁 종료 관리</b>", unsafe_allow_html=True)
        end_date = st.date_input("종료일", value=datetime.now())
        end_time = st.time_input("종료시간", value=default_time, key="end_time_key")
        category = st.selectbox("구분", ["1차 검사", "2차 검사", "3차 검사", "4차 검사", "Sample", "완불재검"])
        
        st.markdown("<br><b>🔌 추가 정보</b>", unsafe_allow_html=True)
        idle_time = st.number_input("휴동시간 (분)", min_value=0, value=0)
        
        start_dt = datetime.combine(start_date, start_time)
        end_dt = datetime.combine(end_date, end_time)
        raw_duration = int((end_dt - start_dt).total_seconds() / 60)
        duration_minutes = max(0, raw_duration - idle_time)
            
        st.text_input("소요시간 (휴동시간 차감됨)", value=f"{duration_minutes:,} 분", disabled=True)
        plating_type = st.selectbox("도금구분", ["A", "B"])

        st.markdown("<br><br><hr>", unsafe_allow_html=True)
        st.write("📊 백업 및 보관용")
        export_df = load_data().copy()
        if not export_df.empty:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False)
                wb = writer.book
                ws = wb.active
                red_font = Font(color="FF0000")
                headers = {cell.value: i for i, cell in enumerate(ws[1])}
                for row in ws.iter_rows(min_row=2):
                    for c_name in ["UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]:
                        if c_name in headers: row[headers[c_name]].number_format = '#,##0'
                    try: yield_val = float(str(row[headers['양품률']].value).replace('%', '').strip())
                    except: yield_val = 100.0
                    if yield_val < 85.0 and '양품률' in headers: row[headers['양품률']].font = red_font
            st.download_button(label="📥 DB 데이터를 엑셀로 다운로드", data=output.getvalue(), file_name=f"VISION_EXPORT_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)
        with st.expander("📤 외부 엑셀 데이터 대량 업로드", expanded=False):
            uploaded_file = st.file_uploader("엑셀 선택", type=['xlsx'])
            if uploaded_file and st.button("스마트 분석 및 구글 DB 저장", type="primary", use_container_width=True):
                with st.spinner("저장 중..."):
                    xls = pd.ExcelFile(uploaded_file)
                    target_sheet = next((s for s in xls.sheet_names if 'Q' in s or '년' in s), xls.sheet_names[0])
                    temp_df = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=100)
                    target_header_idx = next((idx for idx, row in temp_df.iterrows() if sum([1 for kw in ['날짜', '교대', '모델'] if kw in str(row).replace(' ', '')]) >= 2), None)
                    if target_header_idx is not None:
                        import_df = pd.read_excel(xls, sheet_name=target_sheet, header=target_header_idx)
                        def smart_map(c):
                            c=str(c).replace('\n','').replace(' ','')
                            if '모델' in c: return '모델명(MI)'
                            if '날짜' in c or '일자' in c: return '날짜'
                            if '교대' in c: return '교대'
                            if '시작' in c: return '시작시간'
                            if '종료' in c: return '종료시간'
                            if 'LOT' in c.upper(): return 'LOT NO.'
                            if '육안' in c or 'OQC' in c.upper(): return '육안/OQC'
                            if '도장라인' in c or 'LINE' in c.upper(): return '도장라인'
                            if '검사' in c: return '검사 수량'
                            if '양품' in c and ('전/배' in c or '전,배' in c):
                                if '율' in c or '률' in c: return '양품율(전/배 포함)'
                                else: return '양품 수량(전/배 포함)'
                            if '양품' in c:
                                if '율' in c or '률' in c: return '양품률'
                                else: return '양품수량'
                            if '불량수량' in c or ('불량' in c and '수량' in c): return '불량수량'
                            if '완전' in c: return '완전불량'
                            if '전면' in c and '율' not in c: return '전면불량'
                            if '배면' in c and '율' not in c: return '배면불량'
                            if '옵셋' in c: return '옵셋불량'
                            return c
                        import_df.columns = [smart_map(c) for c in import_df.columns]
                        import_df = import_df.dropna(subset=['날짜', '모델명(MI)'])
                        import_df['시작시간'] = import_df['종료시간'] = ""
                        import_df['휴동시간'] = import_df['소요시간'] = import_df['UPH'] = import_df['UPD'] = 0
                        for col in EXCEL_COLUMNS:
                            if col not in import_df.columns: import_df[col] = 0 if col in ["검사 수량"] else ""
                        if save_data_append(import_df[EXCEL_COLUMNS]): st.success("저장 성공!")
                    else: st.error("제목줄을 찾을 수 없습니다.")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📈 종합 분석 데이터 확인", use_container_width=True):
            show_password_dialog()

    main_col1, main_col2 = st.columns([1.1, 0.9])
    save_success_trigger = False

    with main_col1:
        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>📥 VISION Data</h4>
            </div>
        """, unsafe_allow_html=True)
        
        v_row1_col1, v_row1_col2, v_row1_col3 = st.columns(3)
        v_row2_col1, v_row2_col2, v_row2_col3 = st.columns(3)
        v_row3_col1, v_row3_col2, v_row3_col3 = st.columns(3)
        
        with v_row1_col2: good_qty = st.number_input("양품수량", min_value=0, value=0)
        with v_row2_col1: comp_def = st.number_input("완전불량", min_value=0, value=0)
        with v_row2_col2: front_def = st.number_input("전면불량", min_value=0, value=0)
        with v_row2_col3: rear_def = st.number_input("배면불량", min_value=0, value=0)
        with v_row3_col1: offset_def = st.number_input("옵셋불량", min_value=0, value=0)
        with v_row3_col2: shortage_qty = st.number_input("수량부족", min_value=0, value=0)
        with v_row3_col3: etc_def = st.number_input("기타", min_value=0, value=0)
        
        bad_qty = comp_def + front_def + rear_def + offset_def + etc_def
        total_qty = max(0, good_qty + bad_qty - shortage_qty)
            
        with v_row1_col1: st.text_input("검사 수량 (자동)", value=f"{total_qty:,}", disabled=True)
        with v_row1_col3: st.text_input("불량수량 (자동)", value=f"{bad_qty:,}", disabled=True)

        uph_val = int((total_qty / duration_minutes) * 60) if duration_minutes > 0 else 0
        upd_val = uph_val * 22

        good_include_front_rear = good_qty + front_def + rear_def
        if total_qty > 0:
            rate_good = round((good_qty / total_qty) * 100, 1)
            rate_good_inc = round((good_include_front_rear / total_qty) * 100, 1)
            rate_front = round((front_def / total_qty) * 100, 1)
            rate_rear = round((rear_def / total_qty) * 100, 1)
            rate_bad = round((bad_qty / total_qty) * 100, 1)
            comp_rate_num = round(comp_def / total_qty * 100, 1)
            front_rate_num = round(front_def / total_qty * 100, 1)
            rear_rate_num = round(rear_def / total_qty * 100, 1)
            offset_rate_num = round(offset_def / total_qty * 100, 1)
        else:
            rate_good = rate_good_inc = rate_front = rate_rear = rate_bad = 0.0
            comp_rate_num = front_rate_num = rear_rate_num = offset_rate_num = 0.0

        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>📋 기타 정보</h4>
            </div>
        """, unsafe_allow_html=True)
        
        etc_col1, etc_col2, etc_col3, etc_col4 = st.columns(4)
        with etc_col1: in_date = st.date_input("입고일")
        with etc_col2: painting_date = st.date_input("도장일")
        with etc_col3: painting_line = st.selectbox("도장라인", ["선택안함", "A Line", "B Line", "C Line"])
        with etc_col4: painting_order = st.number_input("도장순서", min_value=1, value=1)
        
        parts_col1, parts_col2, parts_col3, parts_col4 = st.columns(4)
        num_options = ["선택안함"] + [str(i) for i in range(1, 11)]
        with parts_col1: clip_val = st.selectbox("CLIP", num_options)
        with parts_col2: base_val = st.selectbox("BASE", num_options)
        with parts_col3: cover_val = st.selectbox("COVER", num_options)
        with parts_col4: assembler_val = st.selectbox("조립기", num_options)

        qc_col1, qc_col2 = st.columns(2)
        with qc_col1: oqc_status = st.selectbox("육안/OQC", ["선택안함", "OQC"])
        with qc_col2: worker_name = st.selectbox("작업자", ["선택안함"] + worker_list)
        
        remarks = st.text_area("비고", height=68, placeholder="특이사항을 입력하세요.")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 데이터 저장 (구글 스프레드시트)", use_container_width=True, type="primary"):
            if total_qty == 0: st.warning("입력된 데이터가 없습니다.")
            elif not lot_number: st.warning("LOT 번호를 입력하거나 스캔해주세요.")
            else:
                with st.spinner("저장 및 중복 검증 중..."):
                    db_df = load_data()
                    is_dup = False
                    if not db_df.empty and "LOT NO." in db_df.columns and "교대" in db_df.columns:
                        if len(db_df[(db_df["LOT NO."] == lot_number) & (db_df["교대"] == shift_type)]) > 0:
                            is_dup = True

                    if is_dup: st.error("🚨 이미 처리되었습니다. (동일 교대 내 중복 LOT)")
                    else:
                        new_data = pd.DataFrame([{
                            "날짜": work_date.strftime("%Y-%m-%d"), "교대": shift_type,
                            "시작시간": start_dt.strftime("%H:%M"), "종료시간": end_dt.strftime("%H:%M"),
                            "휴동시간": idle_time, "소요시간": duration_minutes, "구분": category, "호기": unit, 
                            "모델명(MI)": model_name, "도금구분": plating_type, "UPH": uph_val, "UPD": upd_val,
                            "검사 수량": total_qty, "양품수량": good_qty, "양품 수량(전/배 포함)": good_include_front_rear, "불량수량": bad_qty,
                            "양품률": f"{rate_good:.1f}%", "양품율(전/배 포함)": f"{rate_good_inc:.1f}%",
                            "완전불량률": f"{comp_rate_num:.1f}%", "전면불량률": f"{front_rate_num:.1f}%", "배면불량률": f"{rear_rate_num:.1f}%",
                            "완전불량": comp_def, "전면불량": front_def, "배면불량": rear_def, "옵셋불량": offset_def, "수량부족": shortage_qty, "기타": etc_def,
                            "육안/OQC": oqc_status, "비고": remarks, "도장라인": painting_line, "도장일": painting_date.strftime("%Y-%m-%d"), 
                            "도장순서": painting_order, "입고일": in_date.strftime("%Y-%m-%d"), "LOT NO.": lot_number, 
                            "CLIP": clip_val, "BASE": base_val, "COVER": cover_val, "조립기": assembler_val, 
                            "월": f"{work_date.month}월", "작업자": worker_name
                        }])
                        if save_data_append(new_data):
                            st.success("저장 완료!")
                            # 💡 저장 성공 시 입력창 완전 초기화
                            st.session_state.lot_input_field = ""
                            save_success_trigger = True  

    if category == "1차 검사" and total_qty > 0:
        if comp_rate_num > 5.0:
            if not st.session_state.comp_warned:
                show_sbl_warning("완전불량", comp_rate_num)
                st.session_state.comp_warned = True
        else: st.session_state.comp_warned = False
        if front_rate_num > 5.0:
            if not st.session_state.front_warned:
                show_sbl_warning("전면불량", front_rate_num)
                st.session_state.front_warned = True
        else: st.session_state.front_warned = False
        if rear_rate_num > 5.0:
            if not st.session_state.rear_warned:
                show_sbl_warning("배면불량", rear_rate_num)
                st.session_state.rear_warned = True
        else: st.session_state.rear_warned = False
        if offset_rate_num > 5.0:
            if not st.session_state.offset_warned:
                show_sbl_warning("옵셋불량", offset_rate_num)
                st.session_state.offset_warned = True
        else: st.session_state.offset_warned = False
    else:
        st.session_state.comp_warned = st.session_state.front_warned = st.session_state.rear_warned = st.session_state.offset_warned = False

    with main_col2:
        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>📊 Yield Report</h4>
            </div>
        """, unsafe_allow_html=True)
        
        m_col1, m_col2 = st.columns(2)
        with m_col1: st.metric(label="검사수량 총합", value=f"{total_qty:,} EA")
        with m_col2: st.metric(label="현재 양품율", value=f"{rate_good:.1f}%")

        with st.expander("🎨 Color Option", expanded=False):
            color_col1, color_col2 = st.columns(2)
            with color_col1: c_yield = st.color_picker("✅ 양품 (OK)", "#3b82f6")
            with color_col2: c_bad = st.color_picker("❌ 불량 (NG)", "#ef4444")

        fig_donut = go.Figure(go.Pie(labels=['양품율', '불량율'], values=[rate_good, rate_bad], hole=.65, marker=dict(colors=[c_yield, c_bad], line=dict(color='#0f172a', width=3.5)), hoverinfo="label+percent", textinfo="none"))
        fig_donut.update_layout(title="양품율 / 불량율 점유 분포", title_font={'color': '#94a3b8', 'size': 13}, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=280, margin=dict(t=30, b=10, l=10, r=10), showlegend=True, legend=dict(font=dict(color="#94a3b8"), orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5), annotations=[dict(text=f"{rate_good:.1f}%", x=0.5, y=0.5, font_size=28, font_color="#f8fafc", font_weight="bold", showarrow=False)])
        st.plotly_chart(fig_donut, use_container_width=True)
        
        df_defects = pd.DataFrame({"불량 항목": ['완전불량', '전면불량', '배면불량', '옵셋불량'], "비율 (%)": [comp_rate_num, front_rate_num, rear_rate_num, offset_rate_num]})
        y_max = max(df_defects["비율 (%)"]) * 1.25 if not df_defects.empty and max(df_defects["비율 (%)"]) > 0 else 5

        fig_bar = px.bar(df_defects, x="불량 항목", y="비율 (%)", color="불량 항목", text="비율 (%)", color_discrete_map={'완전불량': c_bad, '전면불량': '#f59e0b', '배면불량': '#10b981', '옵셋불량': '#6366f1'})
        fig_bar.update_traces(texttemplate='%{text:.1f}%', textposition='outside', textfont=dict(color='#f8fafc', size=12))
        fig_bar.update_layout(title="불량 상세 분포", title_font={'color': '#94a3b8', 'size': 13}, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=280, showlegend=False, margin=dict(t=40, b=10, l=10, r=10), xaxis={'tickfont': {'color': '#f8fafc'}, 'title': ''}, yaxis={'tickfont': {'color': '#94a3b8'}, 'gridcolor': '#334155', 'range': [0, y_max]})
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---") 
    with st.expander("📋 최근 저장 데이터 List (구글 시트 연동중 - 수정 가능)", expanded=True):
        df_history = load_data().copy()
        if not df_history.empty:
            recent_10 = df_history.iloc[::-1].head(10).copy()
            valid_cols = [col for col in EXCEL_COLUMNS if col in recent_10.columns]
            
            num_cols = ["UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
            display_df = recent_10[valid_cols].copy()
            for col in num_cols:
                if col in display_df.columns:
                    display_df[col] = display_df[col].apply(lambda x: f"{float(x):,.0f}" if pd.notnull(x) and str(x).replace('.','',1).isdigit() else "0")
                    
            edited_df = st.data_editor(display_df, use_container_width=True, hide_index=True)
            
            if not display_df.equals(edited_df):
                st.info("💡 데이터가 변경되었습니다. 값을 수정한 후 아래 덮어쓰기 버튼을 눌러주세요.")
                if st.button("🔄 변경된 데이터 구글 시트에 덮어쓰기", type="primary"):
                    with st.spinner("구글 스프레드시트에 데이터를 덮어쓰는 중입니다..."):
                        df_full = load_data().copy()
                        changes_applied = False
                        
                        for idx in edited_df.index:
                            orig_row = display_df.loc[idx]
                            edit_row = edited_df.loc[idx]
                            
                            if not orig_row.equals(edit_row):
                                def get_val(col_name):
                                    try: return int(str(edit_row.get(col_name, 0)).replace(',', ''))
                                    except: return 0
                                    
                                good_qty = get_val("양품수량")
                                comp_def = get_val("완전불량")
                                front_def = get_val("전면불량")
                                rear_def = get_val("배면불량")
                                offset_def = get_val("옵셋불량")
                                shortage_qty = get_val("수량부족")
                                etc_def = get_val("기타")
                                
                                bad_qty = comp_def + front_def + rear_def + offset_def + etc_def
                                total_qty = good_qty + bad_qty - shortage_qty
                                if total_qty < 0: total_qty = 0
                                good_include_front_rear = good_qty + front_def + rear_def

                                dur_min = int(edit_row.get("소요시간", 0))
                                uph_update = int((total_qty / dur_min) * 60) if dur_min > 0 else 0
                                upd_update = uph_update * 22
                                
                                if total_qty > 0:
                                    rate_good = round((good_qty / total_qty) * 100, 1)
                                    rate_good_inc = round((good_include_front_rear / total_qty) * 100, 1)
                                    comp_rate_num = round(comp_def / total_qty * 100, 1)
                                    front_rate_num = round(front_def / total_qty * 100, 1)
                                    rear_rate_num = round(rear_def / total_qty * 100, 1)
                                else:
                                    rate_good = rate_good_inc = comp_rate_num = front_rate_num = rear_rate_num = 0.0
                                    
                                for col in valid_cols:
                                    if col == "검사 수량": val = total_qty
                                    elif col == "불량수량": val = bad_qty
                                    elif col == "양품 수량(전/배 포함)": val = good_include_front_rear
                                    elif col == "양품률": val = f"{rate_good:.1f}%"
                                    elif col == "양품율(전/배 포함)": val = f"{rate_good_inc:.1f}%"
                                    elif col == "완전불량률": val = f"{comp_rate_num:.1f}%"
                                    elif col == "전면불량률": val = f"{front_rate_num:.1f}%"
                                    elif col == "배면불량률": val = f"{rear_rate_num:.1f}%"
                                    elif col == "UPH": val = uph_update
                                    elif col == "UPD": val = upd_update
                                    elif col in num_cols: val = get_val(col)
                                    else: val = edit_row[col]
                                    df_full.at[idx, col] = val
                                changes_applied = True
                                
                        if changes_applied:
                            if save_data_overwrite(df_full[valid_cols]):
                                st.success("✅ 변경된 데이터가 구글 스프레드시트에 완벽하게 덮어씌워졌습니다!")
                                import time
                                time.sleep(1)
                                st.rerun()
        else:
            st.caption("현재 구글 시트에 누적된 데이터 이력이 없습니다.")

    if save_success_trigger:
        import time
        time.sleep(1)
        st.rerun()

# ==========================================
# 📈 [새 화면] 종합 데이터 분석 화면 렌더링
# ==========================================
elif st.session_state.current_page == "analysis":
    render_analysis_page()