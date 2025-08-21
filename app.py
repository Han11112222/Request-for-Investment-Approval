import io
import time
import pandas as pd
import streamlit as st
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
import plotly.express as px

st.set_page_config(page_title="엑셀 → 구글 스프레드시트 자동업데이트", layout="wide")

# ---- 인증 & 클라이언트 ----
@st.cache_resource
def get_gspread_client():
    sa_info = st.secrets["gcp_service_account"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

def get_worksheet(client, spreadsheet_id: str, sheet_name: str):
    sh = client.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=50)
    return ws

# ---- UI ----
st.title("엑셀 → 구글 스프레드시트 자동업데이트")
st.caption("파일 업로드 후 **[스프레드시트 업데이트]** 버튼을 누르면 지정한 시트에 덮어씁니다.")

gsheet_id = st.secrets["gsheet_id"]
client = get_gspread_client()

col1, col2, col3 = st.columns([1.2, 1, 1])
with col1:
    file = st.file_uploader("엑셀(.xlsx) 파일 업로드", type=["xlsx"])
with col2:
    target_sheet = st.text_input("기록할 시트 이름", value="자동업데이트")
with col3:
    do_backup = st.toggle("업데이트 전 백업 시트 생성", value=True)

# ---- 파일 처리 ----
df = None
if file is not None:
    # 여러 시트가 있는 경우 첫 번째 시트 사용 (필요하면 sheet_name= 옵션 바꿔도 OK)
    df = pd.read_excel(io.BytesIO(file.read()))
    st.success(f"업로드 완료: {file.name} · {df.shape[0]}행 × {df.shape[1]}열")
    with st.expander("미리보기", expanded=True):
        st.dataframe(df, use_container_width=True)

    # 간단 그래프 예시(열 이름에 맞게 자동 탐색)
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    if len(numeric_cols) >= 1:
        st.subheader("막대 그래프 (첫 번째 숫자열 기준 상위 10)")
        top = df.nlargest(10, numeric_cols[0])
        fig1 = px.bar(top, x=top.columns[0], y=numeric_cols[0], title=f"{numeric_cols[0]} 상위 10")
        st.plotly_chart(fig1, use_container_width=True)

    if len(numeric_cols) >= 1:
        st.subheader("도넛 차트 (첫 번째 숫자열 합계 기준)")
        # 범주형 첫 컬럼과 첫 숫자열로 집계
        cat_col = df.columns[0]
        pie_data = df.groupby(cat_col, as_index=False)[numeric_cols[0]].sum().nlargest(6, numeric_cols[0])
        fig2 = px.pie(pie_data, names=cat_col, values=numeric_cols[0], hole=0.55, title=f"{cat_col}별 {numeric_cols[0]} 비율")
        st.plotly_chart(fig2, use_container_width=True)

# ---- 구글 시트 업데이트 ----
if st.button("스프레드시트 업데이트", type="primary", disabled=(df is None)):
    try:
        ws = get_worksheet(client, gsheet_id, target_sheet)

        # 백업(선택): 기존 시트를 복제하여 타임스탬프 백업
        if do_backup:
            sh = client.open_by_key(gsheet_id)
            ts = time.strftime("%Y%m%d_%H%M%S")
            sh.duplicate_sheet(
                source_sheet_id=ws._properties["sheetId"],
                new_sheet_name=f"{target_sheet}_backup_{ts}"
            )

        # 시트 비우고 덮어쓰기
        ws.clear()
        set_with_dataframe(ws, df)
        st.success(f"✅ '{target_sheet}' 시트에 {df.shape[0]}행 {df.shape[1]}열 업데이트 완료!")
    except gspread.exceptions.APIError as e:
        st.error(f"Google API 오류: {e}")
    except Exception as e:
        st.error(f"업데이트 실패: {e}")

st.caption("※ 시트 공유: 서비스계정 이메일을 구글 시트 **편집자**로 공유해야 합니다.")
