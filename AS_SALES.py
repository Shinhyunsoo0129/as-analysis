import streamlit as st
import pandas as pd
import io
import plotly.express as px

# 넓은 레이아웃 사용
st.set_page_config(page_title="유상 AS 매출 집계", layout="wide")

# 제목 + 안내 문구
st.markdown(
    """
    <h1 style='display: inline;'>📈 유상 AS 매출 집계</h1><br>
    <span style='color: red; font-size: 24px; white-space: nowrap; display: inline-block;'>
        ※ 업로드할 파일은 ERP의 <span style="color: blue;"><u>'AS관리'</u></span> 메뉴의 
        <span style="color: blue;"><u>'AS프로젝트매출관리'</u></span>에서 다운 받은 파일을 업로드하세요!
    </span>
    """,
    unsafe_allow_html=True
)

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("📤 엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(
            f"""
            ❌ 엑셀 파일을 열 수 없습니다.  
            🔒 DRM(디지털 권한 관리)으로 보호된 파일일 수 있습니다.  
            👉 오류 내용: {e}
            """
        )
        st.stop()

    # 0) AS구분 필터 (유상, 단품판매만 포함)
    df = df[df["AS구분"].isin(["유상", "단품판매"])]

    # 1) 제품군 분류
    def classify_product_group(x):
        if x in ["설비제어", "중단사업", "가스솔루션", "공용", "공무"]:
            return "설비제어"
        elif x == "평형수처리":
            return "BWMS"
        elif x == "배전반":
            return "배전반"
        else:
            return None

    df["제품군"] = df["제품군(1)"].apply(classify_product_group)
    df = df[df["제품군"].notnull()]

    # 필터 UI
    담당자_list = ["전체"] + sorted(df["담당자"].dropna().unique().tolist())
    제품군_list = ["전체"] + sorted(df["제품군"].dropna().unique().tolist())

    선택_담당자 = st.selectbox("담당자를 선택하세요", 담당자_list)
    선택_제품군 = st.selectbox("제품군을 선택하세요", 제품군_list)

    필터된_df = df.copy()
    if 선택_담당자 != "전체":
        필터된_df = 필터된_df[필터된_df["담당자"] == 선택_담당자]
    if 선택_제품군 != "전체":
        필터된_df = 필터된_df[필터된_df["제품군"] == 선택_제품군]

    집계결과 = (
        필터된_df.groupby("제품군")[["당월매출액", "당월매출원가", "당월손익"]]
        .sum()
        .reset_index()
    )

    집계결과["이익율(%)"] = 집계결과.apply(
        lambda row: (row["당월손익"] / row["당월매출액"] * 100) if row["당월매출액"] != 0 else 0,
        axis=1
    )

    total_row = pd.DataFrame({
        "제품군": ["합계"],
        "당월매출액": [집계결과["당월매출액"].sum()],
        "당월매출원가": [집계결과["당월매출원가"].sum()],
        "당월손익": [집계결과["당월손익"].sum()],
    })
    total_row["이익율(%)"] = (
        total_row["당월손익"] / total_row["당월매출액"] * 100
    ).fillna(0)

    집계결과 = pd.concat([집계결과, total_row], ignore_index=True)

    포맷된_집계결과 = 집계결과.copy()
    for col in ["당월매출액", "당월매출원가", "당월손익"]:
        포맷된_집계결과[col] = 포맷된_집계결과[col].apply(lambda x: f"₩{x:,.0f}")
    포맷된_집계결과["이익율(%)"] = 집계결과["이익율(%)"].apply(lambda x: f"{x:.2f}%")

    st.subheader("📊 집계 결과 (단위: 원 ₩)")
    st.dataframe(포맷된_집계결과, use_container_width=True)

    chart_df = 집계결과[집계결과["제품군"] != "합계"]
    melt_df = chart_df.melt(
        id_vars="제품군", 
        value_vars=["당월매출액", "당월매출원가", "당월손익"],
        var_name="항목", value_name="금액"
    )

    st.subheader("📊 제품군별 매출/원가/손익 차트")
    fig = px.bar(
        melt_df,
        x="제품군",
        y="금액",
        color="항목",
        barmode="group",
        text="금액",
        title="제품군별 매출/원가/손익 비교",
        height=500,
    )
    fig.update_traces(texttemplate="%{text:,}", textposition="outside")
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
    st.plotly_chart(fig, use_container_width=True)

    사용된_제품군 = chart_df["제품군"].unique().tolist()
    원본데이터 = 필터된_df[필터된_df["제품군"].isin(사용된_제품군)]

    def convert_df_to_excel(summary_df, original_df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            summary_df.to_excel(writer, index=False, sheet_name="집계결과")
            original_df.to_excel(writer, index=False, sheet_name="원본데이터")

            workbook = writer.book
            ws1 = writer.sheets["집계결과"]
            ws2 = writer.sheets["원본데이터"]

            for i, col in enumerate(summary_df.columns):
                ws1.set_column(i, i, 15)
            for i, col in enumerate(original_df.columns):
                ws2.set_column(i, i, 18)

            chart_data = summary_df[summary_df["제품군"] != "합계"]
            chart = workbook.add_chart({'type': 'column'})
            for idx, column in enumerate(["당월매출액", "당월매출원가", "당월손익"]):
                chart.add_series({
                    'name': column,
                    'categories': ['집계결과', 1, 0, len(chart_data), 0],
                    'values': ['집계결과', 1, idx + 1, len(chart_data), idx + 1],
                    'data_labels': {'value': True},
                })
            chart.set_title({'name': '제품군별 매출/원가/손익'})
            chart.set_x_axis({'name': '제품군'})
            chart.set_y_axis({'name': '금액'})
            chart.set_style(11)
            ws1.insert_chart("G2", chart)

            색상리스트 = ['#FFEBEE', '#E3F2FD', '#E8F5E9', '#FFF8E1', '#F3E5F5', '#E0F2F1']
            제품군_고유값 = original_df["제품군"].dropna().unique().tolist()
            색상매핑 = {v: 색상리스트[i % len(색상리스트)] for i, v in enumerate(제품군_고유값)}

            for i, 제품군 in enumerate(제품군_고유값):
                조건포맷 = workbook.add_format({'bg_color': 색상매핑[제품군]})
                ws2.conditional_format(
                    f"A2:Z{len(original_df)+1}", {
                        'type': 'formula',
                        'criteria': f'=$Z2="{제품군}"',
                        'format': 조건포맷
                    }
                )

        return output.getvalue()

    excel_data = convert_df_to_excel(집계결과, 원본데이터)
    st.download_button(
        label="📥 집계 결과 엑셀 다운로드",
        data=excel_data,
        file_name="AS_매출_집계.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
