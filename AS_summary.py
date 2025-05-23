import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📊 AS 접수/조치/조치일 집계 시스템")

# 안내 문구 (uc81c목 바로 아래)
st.markdown("""
<p style='font-size:24px; color:red;'>
※ 업로드할 파일은 ERP의 
<span style='color:blue; font-weight:bold;'>"AS현황 및 최종완료"</span>
에서 다운 받은 파일을 업로드하세요!
</p>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요.", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    df = df[df['전자결재번호상태'] == '종결'].copy()
    df['AS접수일자'] = pd.to_datetime(df['AS접수일자'], errors='coerce')
    df['기술적종료일자'] = pd.to_datetime(df['기술적종료일자'], errors='coerce')
    today = pd.to_datetime(datetime.today().date())

    df['AS접수년월'] = df['AS접수일자'].dt.strftime('%Y%m')
    df['기술적종료년월'] = df['기술적종료일자'].dt.strftime('%Y%m')

    def classify_status(s):
        if s in ['접수', '조치중']:
            return '조치중'
        elif s in ['기술적종료', '공사완료', '최종완료']:
            return '조치완료'
        return '기타'

    df['진행상태'] = df['AS진행상태'].apply(classify_status)
    df['당월조치대상'] = np.where(df['AS접수년월'] == df['기술적종료년월'], 'O', 'X')
    df['조치일'] = (df['기술적종료일자'] - df['AS접수일자']).dt.days
    df['미조치일'] = np.where(df['기술적종료일자'].isna(), (today - df['AS접수일자']).dt.days, np.nan)

    reports = {'필터링_원본결과': df.copy()}

    def generate_report(title, sheet_name, filter_df, values, aggfuncs, suffix_cols=None, y_label="건수", rename_map=None, chart_use_avg_col=None):
        st.markdown(f"## {title}")
        pivot = pd.pivot_table(
            filter_df,
            index=['접수담당자', 'AS접수년월'],
            values=values,
            aggfunc=aggfuncs,
            fill_value=0
        ).reset_index()

        if suffix_cols:
            for new_col, numerator, denominator in suffix_cols:
                pivot[new_col] = pivot[numerator] / pivot[denominator]
                pivot[new_col] = pivot[new_col].round(1)

        pivot['담당자_년월'] = pivot['접수담당자'] + "_" + pivot['AS접수년월']

        if rename_map:
            pivot = pivot.rename(columns=rename_map)
            values = [rename_map.get(v, v) for v in values]
            if suffix_cols:
                for col_tuple in suffix_cols:
                    values.append(col_tuple[0])

        reports[sheet_name] = pivot.copy()
        st.dataframe(pivot)

        chart_cols = values.copy()
        if chart_use_avg_col:
            chart_cols = [v if v != chart_use_avg_col[0] else chart_use_avg_col[1] for v in values]

        if len(chart_cols) == 1:
            fig = px.bar(
                pivot,
                x='담당자_년월',
                y=chart_cols[0],
                labels={'담당자_년월': '담당자 및 년월', chart_cols[0]: y_label},
                title=title,
                text=chart_cols[0]
            )
        else:
            melted = pivot.melt(id_vars='담당자_년월', value_vars=chart_cols, var_name='항목', value_name='값')
            fig = px.bar(
                melted,
                x='담당자_년월',
                y='값',
                color='항목',
                barmode='group',
                labels={'담당자_년월': '담당자 및 년월', '값': y_label},
                title=title,
                text='값'
            )

        fig.update_traces(textposition='outside', textfont_size=12)
        fig.update_layout(uniformtext_minsize=10, uniformtext_mode='hide')
        st.plotly_chart(fig, use_container_width=True)

    # 보고서 1
    generate_report(
        "담당자 및 월별 접수 건수",
        "접수건수",
        df,
        values=['AS접수번호'],
        aggfuncs={'AS접수번호': 'count'},
        y_label="접수 건수"
    )

    # 보고서 2
    df2 = df.copy()
    df2['조치완료건수'] = np.where(df2['진행상태'] == '조치완료', 1, 0)
    df2['조치중건수'] = np.where(df2['진행상태'] == '조치중', 1, 0)

    generate_report(
        "담당자 및 월별 접수 및 조치완료 건수",
        "접수및조치건수",
        df2,
        values=['AS접수번호', '조치완료건수', '조치중건수'],
        aggfuncs={
            'AS접수번호': 'count',
            '조치완료건수': 'sum',
            '조치중건수': 'sum'
        },
        rename_map={'AS접수번호': 'AS접수건수'},
        y_label="건수"
    )

    # 보고서 3
    generate_report(
        "담당자 및 월별 조치기간",
        "조치기간",
        df[df['진행상태'] == '조치완료'],
        values=['AS접수번호', '조치일'],
        aggfuncs={
            'AS접수번호': 'count',
            '조치일': 'sum'
        },
        suffix_cols=[('평균 조치일', '조치일', 'AS접수번호')],
        rename_map={'AS접수번호': 'AS접수건수'},
        y_label="조치일 수",
        chart_use_avg_col=('조치일', '평균 조치일')
    )

    # 보고서 4
    generate_report(
        "담당자 및 월별 미조치기간",
        "미조치기간",
        df[df['진행상태'] == '조치중'],
        values=['AS접수번호', '미조치일'],
        aggfuncs={
            'AS접수번호': 'count',
            '미조치일': 'sum'
        },
        suffix_cols=[('평균 미조치일', '미조치일', 'AS접수번호')],
        rename_map={'AS접수번호': 'AS접수건수'},
        y_label="미조치일 수",
        chart_use_avg_col=('미조치일', '평균 미조치일')
    )

    # 보고서 5
    generate_report(
        "담당자 및 월별 당월조치대상",
        "당월조치대상",
        df[df['당월조치대상'] == 'O'],
        values=['AS접수번호', '조치일'],
        aggfuncs={
            'AS접수번호': 'count',
            '조치일': 'sum'
        },
        suffix_cols=[('평균 조치일', '조치일', 'AS접수번호')],
        rename_map={'AS접수번호': 'AS접수건수'},
        y_label="조치일 수",
        chart_use_avg_col=('조치일', '평균 조치일')
    )

    # 전체 엘셀 다운로드
    if reports:
        st.markdown("### 📦 전체 보고서 통합 다운로드")
        output_all = io.BytesIO()
        with pd.ExcelWriter(output_all, engine='xlsxwriter') as writer:
            for sheet, data in reports.items():
                sheet_name = sheet[:31]
                data.to_excel(writer, index=False, sheet_name=sheet_name)
        st.download_button("📅 전체 집계 결과 엘셀 다운로드", output_all.getvalue(), file_name="AS_분석_보고서.xlsx")
