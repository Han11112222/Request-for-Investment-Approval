import io
import time
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px

import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

# -----------------------------
# 페이지 설정
# -----------------------------
st.set_page_config(page_title="배관투자 자동업데이트", layout="wide")
st.title("엑셀 → 구글 스프레드시트 자동업데이트")
st.caption("파일 업로드 후 [스프레드시트 업데이트] 버튼을 누르면 지정한 시트에 덮어쓰기 또는 신규 시트로 기록됩니다.")

# -----------------------------
# 구글 인증 및 스프레드시트 접근
# -----------------------------
GSHEET_ID = st.secrets["gsheet_id"]
SERVICE_ACCOUNT_INFO = st.secrets["gcp_service_account"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
gc = gspread.authorize(creds)
sh = gc.open_by_key(GSHEET_ID)

# -----------------------------
# 업로드 UI
# -----------------------------
uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

colA, colB, colC = st.columns([1, 1, 1])
with colA:
    write_mode = st.radio("기록 방식", ["덮어쓰기(기존 시트)", "신규 시트 생성"], horizontal=True)
with colB:
    target_ws_name = st.text_input("기존/신규 시트명", value="자동업데이트")
with colC:
    make_charts = st.toggle("차트 생성", value=True)

# -----------------------------
# 엑셀 → DataFrame 로드
# -----------------------------
def load_excel_to_df(file: io.BytesIO) -> dict:
    dfs = {}
    xls = pd.ExcelFile(file, engine="openpyxl")
    for sheet in xls.sheet_names:
        df = xls.parse(sheet_name=sheet).dropna(how="all").dropna(axis=1, how="all")
        dfs[sheet] = df
    return dfs

def get_or_create_worksheet(spreadsheet, name):
    try:
        return spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=name, rows=1000, cols=26)

def write_df(spreadsheet, ws_name, df: pd.DataFrame):
    ws = get_or_create_worksheet(spreadsheet, ws_name)
    ws.clear()
    set_with_dataframe(ws, df)  # DataFrame을 통째로 시트에 쓰기
    return ws

def find_table_for_charts(dfs: dict) -> pd.DataFrame | None:
    keys = list(dfs.keys())
    for name in keys:
        df = dfs[name]
        cols = [str(c) for c in df.columns]
        if any(k in "".join(cols) for k in ["계획", "실적", "승인", "금액"]):
            return df
    return dfs[keys[0]] if keys else None

# -----------------------------
# 메인 로직
# -----------------------------
if uploaded:
    dfs = load_excel_to_df(uploaded)
    st.success(f"시트 {len(dfs)}개를 읽었습니다: {', '.join(list(dfs.keys())[:6])} ...")

    first = list(dfs.keys())[0]
    st.subheader("미리보기")
    st.dataframe(dfs[first].head(30), use_container_width=True)

    # 구글 시트 업데이트 버튼
    if st.button("📝 스프레드시트 업데이트", type="primary", use_container_width=True):
        with st.spinner("구글 스프레드시트에 쓰는 중..."):
            written = []
            base_ws = (target_ws_name or "자동업데이트").strip()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            for sheet_name, df in dfs.items():
                ws_name = base_ws if write_mode.startswith("덮어쓰기") else f"{base_ws}_{sheet_name}_{timestamp}"
                write_df(sh, ws_name, df)
                written.append(ws_name)
                time.sleep(0.2)

        st.success(f"업데이트 완료: {', '.join(written[:6])} ...")
        st.link_button("스프레드시트 열기", f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}", use_container_width=True)

    # 차트 생성
    if make_charts:
        st.subheader("투자계획/실적 & 승인 비율 대시보드")
        chart_df = find_table_for_charts(dfs).copy()

        cols = list(chart_df.columns)
        cat_col = next((c for c in cols if any(k in str(c) for k in ["구분", "항목", "분류", "계정"])), cols[0])

        # 숫자열 변환
        for c in cols:
            chart_df[c] = pd.to_numeric(chart_df[c], errors="ignore")
        num_cols = [c for c in cols if pd.api.types.is_numeric_dtype(chart_df[c])]

        plan_col = next((c for c in num_cols if "계획" in str(c)), None)
        actual_col = next((c for c in num_cols if "실적" in str(c)), None)
        if plan_col is None and len(num_cols) >= 2:
            plan_col, actual_col = num_cols[:2]

        # 막대그래프
        bar_df = chart_df[[cat_col, plan_col, actual_col]].dropna()
        bar_df = bar_df.rename(columns={cat_col: "항목", plan_col: "계획", actual_col: "실적"})
        fig_bar = px.bar(
            bar_df.melt(id_vars="항목", value_vars=["계획", "실적"], var_name="구분", value_name="값"),
            x="값", y="항목", color="구분", barmode="group", orientation="h",
            title="투자계획(사업계획) vs 실적"
        )
        fig_bar.update_layout(height=500, legend_title_text="")
        st.plotly_chart(fig_bar, use_container_width=True)

        # 도넛그래프
        ratio_source = next((c for c in cols if "승인" in str(c) and pd.api.types.is_numeric_dtype(chart_df[c])), None)
        pie_df = chart_df[[cat_col, ratio_source or plan_col]].rename(columns={cat_col: "항목", (ratio_source or plan_col): "값"}).dropna()
        fig_pie = px.pie(pie_df, names="항목", values="값", hole=0.6, title="배관투자 승인 비율(가중치 기준)")
        fig_pie.update_layout(height=520)
        st.plotly_chart(fig_pie, use_container_width=True)
else:
    st.info("엑셀(.xlsx)을 업로드하면 미리보기와 차트가 표시되고, 버튼으로 구글 시트에 반영됩니다.")
