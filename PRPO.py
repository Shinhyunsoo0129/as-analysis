import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# 타이틀
st.title("📦 발주 및 입고 지연 분석기")

# ✅ 안내 문구 추가 (빨간색 문구 + 파란색 강조)
st.markdown("""
<p style='color:red; font-size:17px;'>
※ 업로드할 파일은 ERP의 구매관리 메뉴 중 
<b><span style='color:blue;'>구매요청현황</span></b>에서 다운 받은 파일을 업로드하세요!
</p>
""", unsafe_allow_html=True)

# 파일 업로드
uploaded_file = st.file_uploader("✔ 분석할 Excel 파일을 업로드하세요 (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=str)
    today = pd.to_datetime(datetime.today().date())

    # 날짜 컬럼 변환
    date_cols = ["요청일자", "발주일자", "납기일자", "최근입고일자"]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    # 결재완료(확정) 필터링
    df = df[df["구매요청상태"] == "결재완료(확정)"].copy()

    # 발주지연 판별
    def classify_order(row):
        if pd.notnull(row["발주일자"]):
            return "발주지연" if row["발주일자"] > row["요청일자"] else "정상"
        else:
            return "정상" if row["요청일자"] > today else "발주지연"

    df["발주지연"] = df.apply(classify_order, axis=1)

    # 입고지연 판별
    def classify_delivery(row):
        if pd.notnull(row["최근입고일자"]):
            return "입고지연" if row["최근입고일자"] > row["납기일자"] else "정상"
        else:
            return "정상" if row["납기일자"] > today else "입고지연"

    df["입고지연"] = df.apply(classify_delivery, axis=1)

    # 발주지연 요약
    order_delay_summary = df.groupby(["구매그룹", "발주지연"]).size().unstack(fill_value=0)
    order_delay_summary.loc["총합계"] = order_delay_summary.sum()
    order_delay_summary = order_delay_summary.reset_index()

    st.subheader("📌 구매그룹별 발주지연 건수")
    st.dataframe(order_delay_summary)

    # 입고지연 요약
    delivery_delay_summary = df.groupby(["구매그룹", "입고지연"]).size().unstack(fill_value=0)
    delivery_delay_summary.loc["총합계"] = delivery_delay_summary.sum()
    delivery_delay_summary = delivery_delay_summary.reset_index()

    st.subheader("📌 구매그룹별 입고지연 건수")
    st.dataframe(delivery_delay_summary)

    # 피벗 교차 집계
    pivot_summary = df.pivot_table(
        index="발주지연",
        columns="입고지연",
        values="프로젝트",
        aggfunc="count",
        margins=True,
        margins_name="총합계",
        fill_value=0
    )
    pivot_summary.index.name = "발주지연"
    pivot_summary.columns.name = "입고지연"

    # 엑셀 변환 함수
    def convert_to_excel(data_df, order_summary, delivery_summary, pivot_df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            data_df.to_excel(writer, index=False, sheet_name="지연리스트")
            order_summary.to_excel(writer, index=False, sheet_name="발주지연 요약")
            delivery_summary.to_excel(writer, index=False, sheet_name="입고지연 요약")
            pivot_df.to_excel(writer, sheet_name="지연교차표")
        return output.getvalue()

    # 다운로드 버튼
    st.download_button(
        label="📥 발주지연 및 입고지연 리스트 다운로드",
        data=convert_to_excel(df, order_delay_summary, delivery_delay_summary, pivot_summary),
        file_name="발주_입고지연_리포트.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 전체 데이터 표시
    st.subheader("📋 전체 데이터")
    st.dataframe(df)
