import os
import re
from datetime import datetime
import pandas as pd
import streamlit as st

# 페이지 설정
st.set_page_config(page_title="Scenario Analysis Dashboard", layout="wide")

# -----------------------------------------------------------------------------
# shadcn/ui 스타일 고유 CSS 주입
# -----------------------------------------------------------------------------
st.markdown("""
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css');
    
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
        background-color: #fafafa;
        color: #09090b;
    }
    
    .main-title {
        font-size: 1.8rem;
        font-weight: 700;
        letter-spacing: -0.02em;
        color: #09090b;
        margin-bottom: 1.5rem;
    }
    
    .section-title {
        font-size: 1.3rem;
        font-weight: 600;
        letter-spacing: -0.01em;
        color: #09090b;
        margin-top: 2rem;
        margin-bottom: 0.75rem;
    }
    
    .subsection-title {
        font-size: 1rem;
        font-weight: 600;
        color: #71717a;
        margin-top: 0.25rem;
        margin-bottom: 0.75rem;
    }
    
    hr {
        border: 0;
        border-top: 1px solid #e4e4e7;
        margin: 2.5rem 0;
    }
    
    div[data-testid="stVerticalBlockBorderContainer"] {
        background-color: #ffffff !important;
        border: 1px solid #e4e4e7 !important;
        border-radius: 0.5rem !important;
        padding: 1.5rem !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05) !important;
        margin-bottom: 1.5rem !important;
    }
    
    section[data-testid="stFileUploadDropzone"] {
        border: 1px dashed #e4e4e7 !important;
        border-radius: 0.5rem !important;
        background-color: #fcfcfc !important;
    }

    /* 일자별 요약 목록 UI 스타일 */
    .summary-date-card {
        background-color: #f8fafc;
        border-left: 4px solid #999999;
        padding: 14px 18px;
        margin-bottom: 14px;
        border-radius: 6px;
    }
    .summary-date-header {
        font-weight: 700;
        font-size: 1.05rem;
        color: #1e293b;
        margin-bottom: 8px;
    }
    .summary-action-title {
        font-weight: 600;
        font-size: 0.95rem;
        color: #0f172a;
        margin-top: 8px;
        margin-bottom: 4px;
    }
    .summary-stock-list {
        margin-left: 20px;
        margin-top: 2px;
        margin-bottom: 8px;
        padding-left: 0;
    }
    .summary-stock-item {
        font-size: 0.9rem;
        color: #334155;
        line-height: 1.5;
    }
    .stock-link {
        color: #2563eb !important;
        text-decoration: underline !important;
        font-weight: 600;
    }
    .stock-link:hover {
        color: #1d4ed8 !important;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 1. 환경 설정 및 상수 정의
# -----------------------------------------------------------------------------
ANALYSIS_DIR = "analysis"
# 요청하신 정규표현식 패턴 적용
FILE_PATTERN = re.compile(
    r"^result_(\d{8})_(\d{4}-\d{2}-\d{2})_(\d{6})(?:_([a-zA-Z0-9]+))?\.xlsx$"
)

BASE_COLS = ["종목코드", "종목명"]
TARGET_COLS = [
    "관리종목지정우려시장안내",
    "관리종목지정조치",
    "관리종목지정해제조치",
    "상장폐지우려시장안내(D-10)",
    "상장폐지우려시장안내(D-5)",
    "상장폐지사유발생",
]

if not os.path.exists(ANALYSIS_DIR):
    os.makedirs(ANALYSIS_DIR)


# -----------------------------------------------------------------------------
# 2. 헬퍼 함수 정의
# -----------------------------------------------------------------------------
def get_valid_files():
    if not os.path.exists(ANALYSIS_DIR):
        return []

    files = os.listdir(ANALYSIS_DIR)
    valid_files = []

    for file in files:
        match = FILE_PATTERN.match(file)
        if match:
            base_date = match.group(1)      # Group 1: 8자리 기준일 (예: 20260101)
            mod_date = match.group(2)       # Group 2: 수정 날짜 (예: 2026-01-01)
            mod_time = match.group(3)       # Group 3: 수정 시각 (예: 120000)
            market_type = match.group(4)    # Group 4: kospi, kosdaq 등 (없으면 None)
            
            full_mod_datetime = f"{mod_date}_{mod_time}"
            market_label = market_type.upper() if market_type else ""
            
            valid_files.append({
                "filename": file, 
                "base_date": base_date, 
                "market_type": market_label,
                "mod_datetime": full_mod_datetime,
                "display_mod": f"{mod_date} {mod_time[:2]}:{mod_time[2:4]}:{mod_time[4:]}"
            })

    if not valid_files:
        return []

    df_files = pd.DataFrame(valid_files)
    df_files = df_files.sort_values(by=["base_date", "market_type", "mod_datetime"], ascending=[False, True, False])
    # 같은 기준일(base_date)과 시장구분(market_type) 중 최신 수정본만 유지
    df_unique = df_files.drop_duplicates(subset=["base_date", "market_type"], keep="first")

    file_list = [
        (
            row["filename"],
            f"{row['market_type']} 기준일: {row['base_date'][:4]}-{row['base_date'][4:6]}-{row['base_date'][6:]} (수정: {row['display_mod']})",
            row["base_date"],
            row["market_type"]
        )
        for _, row in df_unique.iterrows()
    ]
    return file_list


def is_matching_action(raw_action, target_col):
    """단일 시장조치 문구와 열(Target Column) 간의 정밀 매칭 검사"""
    raw = str(raw_action).replace(" ", "").strip()
    target = str(target_col).replace(" ", "").strip()
    
    if target in raw or raw in target:
        return True
        
    if "상장폐지" in target and "상장폐지" in raw:
        if "D-10" in target and "D-10" in raw: return True
        if "D-5" in target and "D-5" in raw: return True
        if "사유" in target and "사유" in raw: return True
        
    if "관리종목" in target and "관리종목" in raw:
        if "우려" in target and "우려" in raw: return True
        if "지정조치" in target and "지정조치" in raw: return True
        if "해제" in target and "해제" in raw: return True

    return False


def process_scenario_data(df_raw, prefix_filter):
    """' / '로 구별된 시장조치를 각각 개별 항목으로 분리하여 표 형식 데이터 생성"""
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=BASE_COLS + TARGET_COLS)

    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    required_cols = ["날짜", "종목코드", "종목명", "시장조치"]
    if not all(col in df_raw.columns for col in required_cols):
        return pd.DataFrame(columns=BASE_COLS + TARGET_COLS)

    df_working = df_raw.copy()
    df_working["날짜"] = pd.to_datetime(df_working["날짜"]).dt.strftime("%Y-%m-%d")
    
    df_working["종목코드"] = df_working["종목코드"].fillna("").astype(str).str.strip()
    df_working["종목코드"] = df_working["종목코드"].apply(lambda x: x.split('.')[0] if '.' in x else x)
    df_working["종목코드"] = df_working["종목코드"].apply(lambda x: x.zfill(6) if x != "" else "")
    
    df_working["종목명"] = df_working["종목명"].fillna("").astype(str).str.strip()
    df_working["시장조치"] = df_working["시장조치"].fillna("").astype(str).str.strip()

    # 복합 사유(' / ')를 행별로 분리(explode)
    exploded_rows = []
    for _, row in df_working.iterrows():
        actions = [a.strip() for a in str(row["시장조치"]).split("/") if a.strip()]
        for act in actions:
            exploded_rows.append({
                "날짜": row["날짜"],
                "종목코드": row["종목코드"],
                "종목명": row["종목명"],
                "시장조치": act
            })
    
    df_exploded = pd.DataFrame(exploded_rows)
    if df_exploded.empty:
        return pd.DataFrame(columns=BASE_COLS + TARGET_COLS)

    df_master = df_exploded[["종목코드", "종목명"]].drop_duplicates().set_index("종목코드")
    if "" in df_master.index:
        df_master = df_master.drop(index="")
        
    for col in TARGET_COLS:
        df_master[col] = ""

    has_data = False

    for idx, row in df_exploded.iterrows():
        raw_action = row["시장조치"]
        if prefix_filter in raw_action:
            for col in TARGET_COLS:
                if is_matching_action(raw_action, col):
                    has_data = True
                    code = row["종목코드"]
                    if code in df_master.index and df_master.loc[code, col] == "":
                        df_master.loc[code, col] = row["날짜"]

    df_result = df_master.reset_index()
    if not has_data:
        return pd.DataFrame(columns=BASE_COLS + TARGET_COLS)

    df_result = df_result.fillna("")
    df_result["종목코드"] = df_result["종목코드"].astype(str).str.zfill(6)
    return df_result[BASE_COLS + TARGET_COLS]


def display_daily_summary_card(df_raw, scenario_name):
    """[일자별 시장조치 요약] 목록형 카드 (종목명 오름차순 정렬 및 번호 리스트 표출)"""
    with st.container(border=True):
        st.markdown(f'<div class="subsection-title">일자별 시장조치 요약 (향후 10거래일)</div>', unsafe_allow_html=True)
        
        if df_raw is None or df_raw.empty:
            st.caption("요약할 데이터가 없습니다.")
            return

        df = df_raw.copy()
        df.columns = [str(c).strip() for c in df.columns]
        required_cols = ["날짜", "종목코드", "종목명", "시장조치"]
        
        if not all(col in df.columns for col in required_cols):
            st.caption("필수 컬럼이 부족하여 일자별 요약을 표시할 수 없습니다.")
            return

        df["날짜"] = pd.to_datetime(df["날짜"]).dt.strftime("%Y-%m-%d")
        df["종목코드"] = df["종목코드"].fillna("").astype(str).str.strip().apply(lambda x: x.split('.')[0] if '.' in x else x).str.zfill(6)
        df["종목명"] = df["종목명"].fillna("").astype(str).str.strip()
        df["시장조치"] = df["시장조치"].fillna("").astype(str).str.strip()

        # 복합 시장조치(' / ') 분리
        exploded_rows = []
        for _, row in df.iterrows():
            actions = [a.strip() for a in str(row["시장조치"]).split("/") if a.strip()]
            for act in actions:
                exploded_rows.append({
                    "날짜": row["날짜"],
                    "종목코드": row["종목코드"],
                    "종목명": row["종목명"],
                    "시장조치": act
                })
        
        df_exploded = pd.DataFrame(exploded_rows)
        if df_exploded.empty:
            st.caption("표시할 시장조치 요약 내역이 없습니다.")
            return

        # 유효 날짜 오름차순 추출 후 상위 10거래일 선택
        unique_dates = sorted([d for d in df_exploded["날짜"].unique() if d])[:10]
        if not unique_dates:
            st.caption("표시할 시장조치 요약 내역이 없습니다.")
            return

        df_filtered = df_exploded[df_exploded["날짜"].isin(unique_dates)]

        html_content = ""
        for date in unique_dates:
            date_df = df_filtered[df_filtered["날짜"] == date]
            if date_df.empty:
                continue

            html_content += f'<div class="summary-date-card">'
            html_content += f'<div class="summary-date-header">{date}</div>'
            
            action_groups = date_df.groupby("시장조치")
            for action, group in action_groups:
                if not action:
                    continue
                
                # 종목명 기준 오름차순 정렬
                sorted_group = group.sort_values(by="종목명", ascending=True)
                
                html_content += f'<div class="summary-action-title">{action}({len(sorted_group)}건)</div>'
                html_content += '<ol class="summary-stock-list">'
                
                # 번호 형태(<ol>/<li>)로 종목 표출
                for _, row in sorted_group.iterrows():
                    stock_link = f'<a href="https://finance.naver.com/item/main.naver?code={row["종목코드"]}" target="_blank" class="stock-link">{row["종목코드"]}</a>'
                    html_content += f'<li class="summary-stock-item">{stock_link} - {row["종목명"]}</li>'
                    
                html_content += '</ol>'
                
            html_content += '</div>'

        st.markdown(html_content, unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 3. 메인 상단 화면 - 데이터 관리 및 파일 업로드 기능
# -----------------------------------------------------------------------------
st.markdown('<div class="main-title">시나리오별 시장조치 데이터 분석</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">데이터 관리</div>', unsafe_allow_html=True)

with st.container(border=True):
    uploaded_file = st.file_uploader("새로운 분석 파일 등록 (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        filename = uploaded_file.name
        if FILE_PATTERN.match(filename):
            save_path = os.path.join(ANALYSIS_DIR, filename)
            with open(save_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success(f"파일 업로드 완료: {filename}")
            st.rerun()
        else:
            st.error("파일명 규칙이 맞지 않습니다. (예: result_20260101_2026-01-01_120000.xlsx 또는 result_20260101_2026-01-01_120000_kospi.xlsx)")

    available_files = get_valid_files()
    if not available_files:
        st.info("analysis 폴더에 조건에 맞는 파일이 없거나 업로드된 파일이 없습니다.")
        st.stop()

    file_options = {display: (name, base_date, market_type) for name, display, base_date, market_type in available_files}
    selected_display = st.selectbox("분석 데이터 선택", options=list(file_options.keys()))
    selected_filename, selected_base_date, selected_market_type = file_options[selected_display]


# -----------------------------------------------------------------------------
# 4. 데이터 로드
# -----------------------------------------------------------------------------
file_path = os.path.join(ANALYSIS_DIR, selected_filename)
try:
    df_worst_raw = pd.read_excel(file_path, sheet_name="worst_scenario")
    df_best_raw = pd.read_excel(file_path, sheet_name="best_scenario")
except Exception as e:
    st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
    st.stop()


# -----------------------------------------------------------------------------
# 5. 메인 화면 - 데이터 시각화 및 네이버 금융 링크 연동
# -----------------------------------------------------------------------------
st.markdown(f'<div class="section-title" style="color: #27272a;">기준일 : {selected_base_date} ({selected_market_type})</div>', unsafe_allow_html=True)


def display_scenario_data(df_raw, scenario_name):
    df_market_cap = process_scenario_data(df_raw, "시가총액")
    df_stock_price = process_scenario_data(df_raw, "주가")
    
    if df_market_cap.empty and df_stock_price.empty:
        df_total = pd.DataFrame(columns=BASE_COLS + TARGET_COLS)
    elif df_market_cap.empty:
        df_total = df_stock_price.copy()
    elif df_stock_price.empty:
        df_total = df_market_cap.copy()
    else:
        df_total = df_market_cap.copy()
        df_stock_idx = df_stock_price.set_index("종목코드")
        
        for code in df_stock_idx.index:
            if code not in df_total["종목코드"].values:
                new_row = {"종목코드": code, "종목명": df_stock_idx.loc[code, "종목명"]}
                for c in TARGET_COLS:
                    new_row[c] = df_stock_idx.loc[code, c]
                df_total = pd.concat([df_total, pd.DataFrame([new_row])], ignore_index=True)
            else:
                idx = df_total[df_total["종목코드"] == code].index[0]
                for col in TARGET_COLS:
                    if df_total.loc[idx, col] == "" and df_stock_idx.loc[code, col] != "":
                        df_total.loc[idx, col] = df_stock_idx.loc[code, col]
                        
    # 종목코드 값에 네이버 금융 URL 결합
    def convert_to_link(df):
        if not df.empty:
            df["종목코드"] = df["종목코드"].astype(str).str.zfill(6)
            df["종목코드"] = "https://finance.naver.com/item/main.naver?code=" + df["종목코드"]
        return df

    df_total = convert_to_link(df_total)
    df_market_cap = convert_to_link(df_market_cap)
    df_stock_price = convert_to_link(df_stock_price)

    # LinkColumn 구성 (code= 뒤 6자리만 표출)
    link_config = {
        "종목코드": st.column_config.LinkColumn(
            "종목코드",
            display_text=r"code=(.*)"
        )
    }
                        
    # 전체 결과 카드
    with st.container(border=True):
        st.markdown(f'<div class="subsection-title">{scenario_name} 종합 현황</div>', unsafe_allow_html=True)
        st.dataframe(df_total, use_container_width=True, hide_index=True, column_config=link_config)
    
    # 2단 분할 서브 카드
    col1, col2 = st.columns(2)
    
    with col1:
        with st.container(border=True):
            st.markdown('<div class="subsection-title">시가총액 기준 시장조치</div>', unsafe_allow_html=True)
            if not df_market_cap.empty:
                st.dataframe(df_market_cap, use_container_width=True, hide_index=True, column_config=link_config)
            else:
                st.caption("해당하는 시가총액 관련 조치 데이터가 없습니다.")
        
    with col2:
        with st.container(border=True):
            st.markdown('<div class="subsection-title">주가 기준 시장조치</div>', unsafe_allow_html=True)
            if not df_stock_price.empty:
                st.dataframe(df_stock_price, use_container_width=True, hide_index=True, column_config=link_config)
            else:
                st.caption("해당되는 주가 관련 조치 데이터가 없습니다.")


# 1. Worst Scenario 영역
st.markdown('<div class="section-title" style="color: #dc2626;">Worst Scenario 분석</div>', unsafe_allow_html=True)
display_daily_summary_card(df_worst_raw, "Worst Scenario")
display_scenario_data(df_worst_raw, "Worst Scenario")

st.markdown("<hr>", unsafe_allow_html=True)

# 2. Best Scenario 영역
st.markdown('<div class="section-title" style="color: #16a34a;">Best Scenario 분석</div>', unsafe_allow_html=True)
display_daily_summary_card(df_best_raw, "Best Scenario")
display_scenario_data(df_best_raw, "Best Scenario")
