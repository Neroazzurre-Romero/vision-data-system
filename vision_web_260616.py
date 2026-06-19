import streamlit as st
import pandas as pd
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

# [설정] 작업자 명단
worker_list = ["박경섭", "무고사", "재르소", "김동헌"] 

st.set_page_config(
    page_title="VISION DATA KEY-IN SYSTEM ----- (by. Romero)", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------
# 🪄 마법 코드 1: UI 숨김 및 네이티브 앱 스타일 커스텀
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
        const disableKeyboard = () => {
            const doc = window.parent.document;
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
# 🌐 구글 스프레드시트 연동 (초강력 영구 보존 DB)
# ----------------------------------------------------
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

EXCEL_COLUMNS = [
    "날짜", "교대", "시작시간", "종료시간", "휴동시간", "소요시간", "구분", "호기", "모델명", "도금구분", 
    "검사수량", "양품수량", "양품수량(전,배 포함)", "불량수량", "양품율", "양품율(전/배 포함)", 
    "전면 불량율", "배면 불량율", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", 
    "기타", "OQC", "비고", "도장일", "도장 Line", "작업자"
]
SHEET_NAME = "VISION_DATA_DB"

if "google_credentials" not in st.secrets:
    st.error("🚨 구글 스프레드시트 보안 키(Secrets)가 설정되지 않았습니다!")
    st.info("Streamlit Cloud Settings -> Secrets에 JSON 키를 정확히 입력했는지 확인해주세요.")
    st.stop()

# 💡 [핵심 최적화 1] 구글 API 연결 객체 자체를 영구 캐싱 (중복 접속 완전 차단)
@st.cache_resource
def get_sheet():
    try:
        creds_data = st.secrets["google_credentials"]
        
        if isinstance(creds_data, str):
            clean_data = creds_data.strip().strip("'").strip('"')
            creds_dict = json.loads(clean_data, strict=False)
        else:
            creds_dict = dict(creds_data)
            
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace('\\n', '\n')

        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
        client = gspread.authorize(creds)
        return client.open(SHEET_NAME).sheet1
    except Exception as e:
        st.error(f"🚨 구글 연결 초기화 에러: {e}")
        return None

# 💡 [핵심 최적화 2] 60초 데이터 캐싱 유지 및 중복 헤더 에러 원천 차단
@st.cache_data(ttl=60)
def load_data():
    sheet = get_sheet()
    if sheet is None: return pd.DataFrame(columns=EXCEL_COLUMNS)
    try:
        raw_data = sheet.get_all_values()
        
        if not raw_data or len(raw_data) < 2:
            return pd.DataFrame(columns=EXCEL_COLUMNS)
            
        headers = [str(h).strip() for h in raw_data[0]]
        rows = raw_data[1:]
        
        df = pd.DataFrame(rows, columns=headers)
        
        result_df = pd.DataFrame()
        for col in EXCEL_COLUMNS:
            if col in df.columns:
                if isinstance(df[col], pd.DataFrame):
                    result_df[col] = df[col].iloc[:, 0]
                else:
                    result_df[col] = df[col]
            else:
                result_df[col] = ""
                
        return result_df
        
    except Exception as e:
        if "429" in str(e):
            st.warning("⏳ 구글 서버 요청 한도를 초과했습니다. 딱 1분만 기다렸다가 다시 시도해주세요!")
        else:
            st.error(f"🚨 데이터 로드 접근 에러: [{type(e).__name__}] {str(e)}")
        return pd.DataFrame(columns=EXCEL_COLUMNS)

def save_data_append(df):
    sheet = get_sheet()
    if sheet is None: return False
    try:
        current_df = load_data()
        if current_df.empty:
            sheet.append_row(EXCEL_COLUMNS)
            
        df = df.fillna("")
        sheet.append_rows(df.values.tolist())
        
        load_data.clear() 
        return True
    except Exception as e:
        if "429" in str(e):
            st.error("🚨 너무 잦은 저장으로 구글 서버가 일시 차단했습니다. 1분 후 저장해주세요.")
        else:
            st.error(f"🚨 데이터 저장 오류 상세: [{type(e).__name__}] {str(e)}")
        return False

def save_data_overwrite(df):
    sheet = get_sheet()
    if sheet is None: return False
    try:
        sheet.clear()
        df = df.fillna("")
        data_to_upload = [list(df.columns)] + df.values.tolist()
        sheet.append_rows(data_to_upload)
        
        load_data.clear()
        return True
    except Exception as e:
        if "429" in str(e):
            st.error("🚨 너무 잦은 변경으로 구글 서버가 일시 차단했습니다. 1분 후 덮어쓰기 해주세요.")
        else:
            st.error(f"🚨 데이터 덮어쓰기 오류 상세: [{type(e).__name__}] {str(e)}")
        return False

# ----------------------------------------------------
# 🚨 SBL Warning & 관리자 인증
# ----------------------------------------------------
@st.dialog("🚨 SBL Warning!")
def show_sbl_warning(defect_type, rate):
    st.markdown(f"### ⚡ [{defect_type}] 불량 제품을 별도 보관 조치 하세요.")
    st.error(f"현재 1차검사 공정의 {defect_type}율이 **{rate:.1f}%** 로 기준치(5.0%)를 초과하였습니다.")
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

if "current_page" not in st.session_state:
    st.session_state.current_page = "input"

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

    num_cols = ["검사수량", "양품수량", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    if '시작시간' in df.columns:
        df['시간대'] = df['시작시간'].str[:13] + "시"
    else:
        df['시간대'] = df['날짜']

    f_head_col1, f_head_col2 = st.columns([0.95, 0.05])
    with f_head_col1:
        st.markdown("##### 🔍 상세 분석 필터")
    with f_head_col2:
        with st.popover("🎨"):
            st.markdown("**📊 차트 색상 설정**")
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
    with col_opt1:
        date_filter_mode = st.radio("📅 분석 기간 설정", ["전체 누적 데이터", "단일 일자 선택", "특정 기간 지정 검색"], horizontal=True)
    with col_opt2:
        x_axis_mode = st.radio("📊 분석 기준 (X축)", ["일별 (날짜 기준)", "시간별 (시작시간 기준)"], horizontal=True)
    
    if date_filter_mode == "단일 일자 선택":
        selected_date = st.selectbox("분석할 근무일자를 선택하세요", available_dates)
        if selected_date: df = df[df['날짜'] == selected_date]
    elif date_filter_mode == "특정 기간 지정 검색":
        d_col1, d_col2 = st.columns(2)
        try:
            min_date = pd.to_datetime(df['날짜']).min().date()
            max_date = pd.to_datetime(df['날짜']).max().date()
        except:
            min_date = max_date = datetime.now().date()
        with d_col1: start_date_filter = st.date_input("시작 일자", value=min_date)
        with d_col2: end_date_filter = st.date_input("종료 일자", value=max_date)
        df = df[(df['날짜'] >= start_date_filter.strftime("%Y-%m-%d")) & (df['날짜'] <= end_date_filter.strftime("%Y-%m-%d"))]

    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    with f_col1: selected_model = st.selectbox("🏷️ 모델명", ["전체"] + sorted(df['모델명'].dropna().astype(str).unique()))
    with f_col2: selected_unit = st.selectbox("⚙️ 호기", ["전체"] + sorted(df['호기'].dropna().astype(str).unique()))
    with f_col3: selected_category = st.selectbox("🔎 검사구분", ["전체"] + sorted(df['구분'].dropna().astype(str).unique()))
    with f_col4: selected_shift = st.selectbox("🌗 교대", ["전체"] + sorted(df['교대'].dropna().astype(str).unique()))

    if selected_model != "전체": df = df[df['모델명'] == selected_model]
    if selected_unit != "전체": df = df[df['호기'] == selected_unit]
    if selected_category != "전체": df = df[df['구분'] == selected_category]
    if selected_shift != "전체": df = df[df['교대'] == selected_shift]

    if df.empty or df["검사수량"].sum() == 0:
        st.info("선택한 조건에 유효한 데이터가 없습니다.")
        return

    base_col = '시간대' if "시간별" in x_axis_mode else '날짜'
    group_cols = [base_col]
    if selected_model == "전체": group_cols.append('모델명')

    df_grouped = df.groupby(group_cols)[num_cols].sum().reset_index().sort_values(base_col)
    df_grouped['양품율'] = (df_grouped['양품수량'] / df_grouped['검사수량'] * 100).fillna(0).round(1)
    df_grouped['양품율(전,배 포함)'] = ((df_grouped['양품수량'] + df_grouped['전면불량'] + df_grouped['배면불량']) / df_grouped['검사수량'] * 100).fillna(0).round(1)
    df_grouped['불량율'] = (df_grouped['불량수량'] / df_grouped['검사수량'] * 100).fillna(0).round(1)
    df_grouped['완전불량율'] = (df_grouped['완전불량'] / df_grouped['검사수량'] * 100).fillna(0).round(1)
    df_grouped['전면불량율'] = (df_grouped['전면불량'] / df_grouped['검사수량'] * 100).fillna(0).round(1)
    df_grouped['배면불량율'] = (df_grouped['배면불량'] / df_grouped['검사수량'] * 100).fillna(0).round(1)
    df_grouped['옵셋불량율'] = (df_grouped['옵셋불량'] / df_grouped['검사수량'] * 100).fillna(0).round(1)

    is_single_x = len(df_grouped[base_col].unique()) == 1
    def format_labels(series): return series.apply(lambda x: f"{x:.1f}%" if x > 0 else "")

    def create_single_chart(title, metric, base_color):
        fig = go.Figure()
        if selected_model == "전체":
            colors = px.colors.qualitative.Plotly 
            for i, model in enumerate(df_grouped['모델명'].unique()):
                m_data = df_grouped[df_grouped['모델명'] == model]
                color = colors[i % len(colors)]
                if is_single_x:
                    fig.add_trace(go.Bar(name=f"{model} - {metric}", x=m_data[base_col], y=m_data[metric], marker_color=color, text=format_labels(m_data[metric]), textposition='auto'))
                else:
                    fig.add_trace(go.Scatter(name=f"{model} - {metric}", x=m_data[base_col], y=m_data[metric], mode='lines+markers+text', marker_color=color, text=format_labels(m_data[metric]), textposition='top center', line=dict(width=3), marker=dict(size=8)))
        else:
            if is_single_x:
                fig.add_trace(go.Bar(name=metric, x=df_grouped[base_col], y=df_grouped[metric], marker_color=base_color, text=format_labels(df_grouped[metric]), textposition='auto'))
            else:
                fig.add_trace(go.Scatter(name=metric, x=df_grouped[base_col], y=df_grouped[metric], mode='lines+markers+text', marker_color=base_color, text=format_labels(df_grouped[metric]), textposition='top center', line=dict(width=3), marker=dict(size=8)))
                
        max_val = df_grouped[metric].max() if not df_grouped.empty else 0
        y_range = [0, max_val * 1.2 + 2] if max_val > 0 else [0, 10]
        fig.update_layout(title=title, title_font=dict(color='#f8fafc', size=14), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#f8fafc'), xaxis=dict(gridcolor='#334155', title=""), yaxis=dict(gridcolor='#334155', title="비율 (%)", range=y_range), margin=dict(t=50, b=30, l=10, r=10), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(color='#cbd5e1')), barmode='group' if is_single_x else None, hovermode="x unified")
        return fig

    st.markdown("---")
    row1_c1, row1_c2, row1_c3 = st.columns(3)
    with row1_c1: st.plotly_chart(create_single_chart("📈 양품율 트렌드", "양품율", c_yield1), use_container_width=True)
    with row1_c2: st.plotly_chart(create_single_chart("📈 양품율(전,배 포함) 트렌드", "양품율(전,배 포함)", c_yield2), use_container_width=True)
    with row1_c3: st.plotly_chart(create_single_chart("📉 불량율 트렌드", "불량율", c_bad), use_container_width=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    row2_c1, row2_c2, row2_c3, row2_c4 = st.columns(4)
    with row2_c1: st.plotly_chart(create_single_chart("📉 완전불량율 트렌드", "완전불량율", c_comp), use_container_width=True)
    with row2_c2: st.plotly_chart(create_single_chart("📉 전면불량율 트렌드", "전면불량율", c_front), use_container_width=True)
    with row2_c3: st.plotly_chart(create_single_chart("📉 배면불량율 트렌드", "배면불량율", c_rear), use_container_width=True)
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
        model_name = st.selectbox("모델명", ["MEM", "Centaur", "Sphinx-E", "Banff", "Krios", "AV-J", "Seattle", "Juliet-O"])
        
        st.markdown("<br><b>⏱️ 시간 관리</b>", unsafe_allow_html=True)
        start_date = st.date_input("시작일", value=datetime.now())
        start_time = st.time_input("시작시간", value=default_time, key="start_time_key")
        unit = st.selectbox("호기", [f"{i}호기" for i in range(1, 9)])
        
        st.markdown("<br><b>🏁 종료 관리</b>", unsafe_allow_html=True)
        end_date = st.date_input("종료일", value=datetime.now())
        end_time = st.time_input("종료시간", value=default_time, key="end_time_key")
        category = st.selectbox("구분", ["1차검사", "2차검사", "3차검사", "4차검사", "Sample", "완불재검"])
        
        st.markdown("<br><b>🔌 추가 정보</b>", unsafe_allow_html=True)
        idle_time = st.number_input("휴동시간 (분)", min_value=0, value=0)
        
        start_dt = datetime.combine(start_date, start_time)
        end_dt = datetime.combine(end_date, end_time)
        raw_duration = int((end_dt - start_dt).total_seconds() / 60)
        duration_minutes = max(0, raw_duration - idle_time)
            
        st.text_input("소요시간 (휴동시간 차감됨)", value=f"{duration_minutes:,} 분", disabled=True)
        plating_type = st.selectbox("도금구분", ["A", "B"])

        # 💡 엑셀 다운로드 버튼 (클라우드 최적화)
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
                num_format_cols = ["검사수량", "양품수량", "양품수량(전,배 포함)", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
                
                for row in ws.iter_rows(min_row=2):
                    for c_name in num_format_cols:
                        if c_name in headers: row[headers[c_name]].number_format = '#,##0'
                    model_val = row[headers['모델명']].value if '모델명' in headers else ""
                    yield_str = str(row[headers['양품율']].value) if '양품율' in headers else "0"
                    try: yield_val = float(yield_str.replace('%', '').strip())
                    except ValueError: yield_val = 100.0
                    
                    limit = 91.4 if model_val == 'Centaur' else 93.2 if model_val == 'MEM' else 85.0
                    if yield_val < limit and '양품율' in headers: row[headers['양품율']].font = red_font
                    
                    try: total_q = float(row[headers['검사수량']].value)
                    except: total_q = 0
                        
                    if total_q > 0:
                        if '완전불량' in headers and float(row[headers['완전불량']].value or 0) / total_q * 100 > 5.0: row[headers['완전불량']].font = red_font
                        if '전면불량' in headers and float(row[headers['전면불량']].value or 0) / total_q * 100 > 5.0: 
                            row[headers['전면불량']].font = red_font
                            if '전면 불량율' in headers: row[headers['전면 불량율']].font = red_font
                        if '배면불량' in headers and float(row[headers['배면불량']].value or 0) / total_q * 100 > 5.0: 
                            row[headers['배면불량']].font = red_font
                            if '배면 불량율' in headers: row[headers['배면 불량율']].font = red_font
                        if '옵셋불량' in headers and float(row[headers['옵셋불량']].value or 0) / total_q * 100 > 5.0: row[headers['옵셋불량']].font = red_font

            excel_data = output.getvalue()
            st.download_button(
                label="📥 현재 DB 데이터를 엑셀로 다운로드",
                data=excel_data,
                file_name=f"VISION_DATA_EXPORT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.button("📥 현재 DB 데이터를 엑셀로 다운로드", disabled=True, use_container_width=True)
            st.caption("저장된 데이터가 없습니다.")

        # ==========================================
        # 📥 외부 엑셀 데이터 대량 업로드
        # ==========================================
        st.markdown("<br><hr>", unsafe_allow_html=True)
        with st.expander("📤 외부 엑셀 데이터 업로드 (스마트 인식 기능 탑재)", expanded=False):
            st.info("💡 엑셀 상단에 복잡한 서식이나 요약표가 있어도 시스템이 진짜 데이터가 시작되는 표를 자동으로 찾아냅니다.")
            uploaded_file = st.file_uploader("엑셀 파일(.xlsx) 선택", type=['xlsx'])
            if uploaded_file is not None:
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names
                
                # 'Q' 나 '년'이 포함된 데이터 시트를 1순위로 자동 추천 (예: 2026년 1Q)
                recommended_sheet = next((s for s in sheet_names if 'Q' in s or '년' in s), sheet_names[0])
                target_sheet = st.selectbox("📂 데이터가 있는 시트 선택", sheet_names, index=sheet_names.index(recommended_sheet))
                
                if st.button("구글 스프레드시트에 대량 저장", type="primary", use_container_width=True):
                    with st.spinner("데이터 서식을 해독하고 구글 서버에 전송 중입니다..."):
                        try:
                            # 💡 1. 엑셀 상단 서식 무시하고 진짜 헤더(제목줄) 찾기 (최대 100줄 깊이까지 탐색)
                            temp_df = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=100)
                            target_header_idx = None
                            
                            for idx, row in temp_df.iterrows():
                                row_str = row.astype(str).str.replace(r'[\n\r\s]+', '', regex=True)
                                # '날짜', '교대', '모델' 키워드가 2개 이상 있는 줄을 진짜 데이터 표의 시작점으로 간주!
                                match_count = sum([1 for kw in ['날짜', '교대', '모델', '시작'] if sum(row_str.str.contains(kw)) > 0])
                                if match_count >= 2:
                                    target_header_idx = idx
                                    break
                                    
                            if target_header_idx is None:
                                st.error(f"🚨 '{target_sheet}' 시트에서 표의 제목(헤더) 줄을 찾을 수 없습니다.")
                            else:
                                import_df = pd.read_excel(xls, sheet_name=target_sheet, header=target_header_idx)
                                
                                # 💡 2. 컬럼명 표준화 (부분 일치 스마트 매핑 엔진)
                                def smart_map_column(c):
                                    c = str(c).replace('\n', '').replace('\r', '').replace(' ', '')
                                    if '모델' in c: return '모델명'
                                    if '날짜' in c or '일자' in c: return '날짜'
                                    if '교대' in c: return '교대'
                                    if '시작' in c: return '시작시간'
                                    if '종료' in c: return '종료시간'
                                    if '휴동' in c: return '휴동시간'
                                    if '소요' in c: return '소요시간'
                                    if '도금' in c: return '도금구분'
                                    if '호기' in c: return '호기'
                                    if '검사' in c: return '검사수량'
                                    
                                    # 전,배 포함 우선 처리
                                    if '전/배' in c or '전,배' in c:
                                        if '율' in c or '수율' in c: return '양품율(전/배 포함)'
                                        else: return '양품수량(전,배 포함)'
                                        
                                    if '양품' in c:
                                        if '율' in c or '수율' in c: return '양품율'
                                        else: return '양품수량'
                                        
                                    if '불량수량' in c or ('불량' in c and '수량' in c): return '불량수량'
                                    
                                    if '완전' in c: return '완전불량'
                                    if '전면' in c and '율' not in c: return '전면불량'
                                    if '배면' in c and '율' not in c: return '배면불량'
                                    if '옵셋' in c or '오프셋' in c: return '옵셋불량'
                                    
                                    if '부족' in c: return '수량부족'
                                    if '기타' in c: return '기타'
                                    
                                    if 'OQC' in c.upper(): return 'OQC'
                                    if '비고' in c: return '비고'
                                    if '도장일' in c: return '도장일'
                                    if 'LINE' in c.upper() or '라인' in c: return '도장 Line'
                                    if '작업자' in c or '성명' in c or '이름' in c: return '작업자'
                                    
                                    if '구분' in c: return '구분'
                                    return c
                                
                                import_df.columns = [smart_map_column(c) for c in import_df.columns]
                                
                                # 데이터가 없는 빈 줄이나 요약 줄(합계, 평균 등) 제거
                                import_df = import_df.dropna(subset=['날짜', '모델명'])
                                import_df = import_df[~import_df['날짜'].astype(str).str.contains('합계|총계|평균', na=False)]
                                
                                # 💡 3. 스마트 번역기 1: 날짜 포맷 통일 (26.06.19 등 다양한 형식을 2026-06-19 로 변환)
                                def parse_smart_date(d_val):
                                    d_str = str(d_val).strip()
                                    if d_str in ('nan', 'NaT', 'None', ''): return None
                                    if len(d_str.split('.')) == 3 and len(d_str.split('.')[0]) == 2:
                                        d_str = "20" + d_str  # 26.x.x 를 2026.x.x 로
                                    try: return pd.to_datetime(d_str).strftime("%Y-%m-%d")
                                    except: return None
                                
                                import_df['날짜'] = import_df['날짜'].apply(parse_smart_date)
                                import_df = import_df.dropna(subset=['날짜']) 
                                
                                # 💡 4. 스마트 번역기 2: 시간 포맷 통일 (오전/오후/시/분 한글 변환 완벽 대응)
                                def parse_smart_time(t_val, date_val):
                                    t_str = str(t_val).strip()
                                    if t_str in ('nan', 'NaT', 'None', ''): return f"{date_val} 00:00"
                                    if len(t_str) > 10 and '-' in t_str:
                                        try: return pd.to_datetime(t_str).strftime("%Y-%m-%d %H:%M")
                                        except: return t_str
                                    
                                    t_str_clean = t_str.replace('오전', 'AM').replace('오후', 'PM').replace('시', ':00').replace('분', '').replace(' ', '')
                                    try: 
                                        parsed_time = pd.to_datetime(t_str_clean).strftime("%H:%M")
                                        return f"{date_val} {parsed_time}"
                                    except: 
                                        if len(t_str) <= 5: return f"{date_val} {t_str}"
                                        return f"{date_val} 00:00"
                                            
                                import_df['시작시간'] = import_df.apply(lambda x: parse_smart_time(x.get('시작시간'), x['날짜']), axis=1)
                                import_df['종료시간'] = import_df.apply(lambda x: parse_smart_time(x.get('종료시간'), x['날짜']), axis=1)
                                
                                # 💡 5. 스마트 번역기 3: 엑셀 속 제각각인 모델명을 시스템 표준(드롭다운)으로 강제 매핑 변환!
                                def parse_smart_model(m_val):
                                    m_str = str(m_val).upper().replace(' ', '')
                                    if 'CENTAUR' in m_str or 'D51' in m_str: return 'Centaur'
                                    if 'MEM' in m_str or 'D53' in m_str: return 'MEM'
                                    if 'SPHINX' in m_str or 'D31P' in m_str: return 'Sphinx-E'
                                    if 'BANFF' in m_str or 'D14' in m_str: return 'Banff'
                                    if 'KRIOS' in m_str: return 'Krios'
                                    if 'AV-J' in m_str or 'D33C' in m_str: return 'AV-J'
                                    if 'SEATTLE' in m_str: return 'Seattle'
                                    if 'JULIET' in m_str: return 'Juliet-O'
                                    return str(m_val) # 알 수 없는 모델은 원본 텍스트 유지
                                
                                import_df['모델명'] = import_df['모델명'].apply(parse_smart_model)
                                
                                # 누락된 필수 컬럼들 빈 값(또는 0)으로 초기화
                                num_cols = ["휴동시간", "소요시간", "검사수량", "양품수량", "양품수량(전,배 포함)", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
                                rate_cols = ["양품율", "양품율(전/배 포함)", "전면 불량율", "배면 불량율"]
                                
                                for col in EXCEL_COLUMNS:
                                    if col not in import_df.columns:
                                        if col in num_cols: import_df[col] = 0
                                        elif col in rate_cols: import_df[col] = "0.0%"
                                        else: import_df[col] = ""
                                
                                import_df = import_df[EXCEL_COLUMNS]
                                
                                if save_data_append(import_df):
                                    st.success(f"✅ 스마트 분석 성공! [{target_sheet}] 시트에서 데이터 {len(import_df)}건을 추출해 DB에 영구 저장했습니다!")
                                    import time
                                    time.sleep(2)
                                    st.rerun()
                        except Exception as e:
                            st.error(f"데이터 업로드 및 스마트 분석 중 오류 발생: {e}")

        # ==========================================
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
        
        data_input_container = st.container()
        with data_input_container:
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
            total_qty = good_qty + bad_qty - shortage_qty
            if total_qty < 0: total_qty = 0
                
            with v_row1_col1: st.text_input("검사수량", value=f"{total_qty:,}", disabled=True)
            with v_row1_col3: st.text_input("불량수량", value=f"{bad_qty:,}", disabled=True)

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
        
        metadata_container = st.container()
        with metadata_container:
            etc_col1, etc_col2, etc_col3, etc_col4 = st.columns(4)
            with etc_col1: painting_date = st.date_input("도장일")
            with etc_col2: painting_line = st.selectbox("도장 Line", ["선택안함", "A Line", "B Line", "C Line"])
            with etc_col3: oqc_status = st.selectbox("OQC", ["선택안함", "OQC"])
            with etc_col4: worker_name = st.selectbox("작업자", ["선택안함"] + worker_list)
                
            remarks = st.text_area("비고", height=68, placeholder="특이사항을 입력하세요.")

        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("💾 데이터 저장 (구글 스프레드시트)", use_container_width=True, type="primary"):
            if total_qty == 0 and good_qty == 0 and bad_qty == 0:
                st.warning("입력된 데이터가 없습니다.")
            else:
                with st.spinner("구글 스프레드시트에 영구 저장 중..."):
                    new_data = pd.DataFrame([{
                        "날짜": work_date.strftime("%Y-%m-%d"), "교대": shift_type,
                        "시작시간": start_dt.strftime("%Y-%m-%d %H:%M"), "종료시간": end_dt.strftime("%Y-%m-%d %H:%M"),
                        "휴동시간": idle_time, "소요시간": duration_minutes, "구분": category, "호기": unit, "모델명": model_name, "도금구분": plating_type,
                        "검사수량": total_qty, "양품수량": good_qty, "양품수량(전,배 포함)": good_include_front_rear, "불량수량": bad_qty,
                        "양품율": f"{rate_good:.1f}%", "양품율(전/배 포함)": f"{rate_good_inc:.1f}%", "전면 불량율": f"{rate_front:.1f}%", "배면 불량율": f"{rate_rear:.1f}%",
                        "완전불량": comp_def, "전면불량": front_def, "배면불량": rear_def, "옵셋불량": offset_def, "수량부족": shortage_qty, "기타": etc_def,
                        "OQC": oqc_status, "비고": remarks, "도장일": painting_date.strftime("%Y-%m-%d"), "도장 Line": painting_line, "작업자": worker_name
                    }])
                    
                    if save_data_append(new_data):
                        st.success("데이터가 구글 스프레드시트에 영구적으로 안전하게 저장되었습니다!")
                        save_success_trigger = True  

    if category == "1차검사" and total_qty > 0:
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
        st.session_state.comp_warned = False; st.session_state.front_warned = False
        st.session_state.rear_warned = False; st.session_state.offset_warned = False

    with main_col2:
        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>📊 Yield Report</h4>
            </div>
        """, unsafe_allow_html=True)
        
        analytics_container = st.container()
        with analytics_container:
            m_col1, m_col2 = st.columns(2)
            with m_col1: st.metric(label="검사수량 총합", value=f"{total_qty:,} EA")
            with m_col2: st.metric(label="현재 양품율", value=f"{rate_good:.1f}%")

            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("🎨 Color Option", expanded=False):
                color_col1, color_col2 = st.columns(2)
                with color_col1: c_yield = st.color_picker("✅ 양품 (OK)", "#3b82f6")
                with color_col2: c_bad = st.color_picker("❌ 불량 (NG)", "#ef4444")

            fig_donut = go.Figure()
            fig_donut.add_trace(go.Pie(
                labels=['양품율', '불량율'], values=[rate_good, rate_bad],
                hole=.65, marker=dict(colors=[c_yield, c_bad], line=dict(color='#0f172a', width=3.5)),
                hoverinfo="label+percent", textinfo="none", sort=False, pull=[0.02, 0.02]
            ))
            fig_donut.update_layout(title="양품율 / 불량율 점유 분포", title_font={'color': '#94a3b8', 'size': 13}, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=280, margin=dict(t=30, b=10, l=10, r=10), showlegend=True, legend=dict(font=dict(color="#94a3b8"), orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5), annotations=[dict(text=f"{rate_good:.1f}%", x=0.5, y=0.5, font_size=28, font_color="#f8fafc", font_weight="bold", showarrow=False)])
            st.plotly_chart(fig_donut, use_container_width=True)
            
            df_defects = pd.DataFrame({"불량 항목": ['완전불량', '전면불량', '배면불량', '옵셋불량'], "비율 (%)": [comp_rate_num, front_rate_num, rear_rate_num, offset_rate_num]})
            max_rate = max(df_defects["비율 (%)"]) if not df_defects.empty else 0
            y_max = max_rate * 1.25 if max_rate > 0 else 5

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
            
            num_cols = ["검사수량", "양품수량", "양품수량(전,배 포함)", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
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
                                
                                if total_qty > 0:
                                    rate_good = round((good_qty / total_qty) * 100, 1)
                                    rate_good_inc = round((good_include_front_rear / total_qty) * 100, 1)
                                    rate_front = round((front_def / total_qty) * 100, 1)
                                    rate_rear = round((rear_def / total_qty) * 100, 1)
                                else:
                                    rate_good = rate_good_inc = rate_front = rate_rear = 0.0
                                    
                                for col in valid_cols:
                                    if col == "검사수량": val = total_qty
                                    elif col == "불량수량": val = bad_qty
                                    elif col == "양품수량(전,배 포함)": val = good_include_front_rear
                                    elif col == "양품율": val = f"{rate_good:.1f}%"
                                    elif col == "양품율(전/배 포함)": val = f"{rate_good_inc:.1f}%"
                                    elif col == "전면 불량율": val = f"{rate_front:.1f}%"
                                    elif col == "배면 불량율": val = f"{rate_rear:.1f}%"
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
        st.rerun()

# ==========================================
# 📈 [새 화면] 종합 데이터 분석 화면 렌더링
# ==========================================
elif st.session_state.current_page == "analysis":
    render_analysis_page()
