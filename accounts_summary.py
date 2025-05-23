import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import plotly.graph_objects as go

st.set_page_config(page_title='AS채권현황 분석 및 점검 시스템')
st.title('AS채권현황 분석 및 점검 시스템')

st.markdown(
    '<div style="font-size:18px; color:red;">※ 업로드할 파일은 ERP의 <span style="color:blue;">AS현황 및 최종완료</span>에서 다운 받은 파일을 업로드하세요!</div>',
    unsafe_allow_html=True
)
st.markdown(
    '<div style="font-size:18px; color:red;">※ 업로드할 파일은 ERP의 <span style="color:blue;">AS비용현황</span>에서 다운 받은 파일을 업로드하세요!</div>',
    unsafe_allow_html=True
)

as_status_file = st.file_uploader("AS현황 및 최종완료 파일 업로드", type=["xlsx"])
as_cost_file = st.file_uploader("AS비용현황 파일 업로드", type=["xlsx"])

if as_status_file and as_cost_file:
    today = pd.to_datetime(datetime.today().date())

    df_status = pd.read_excel(as_status_file)
    df_status = df_status.dropna(subset=['AS접수번호'])
    df_status = df_status[df_status['전자결재번호상태'] == '종결']
    status_map = df_status.set_index('AS접수번호')[['전자결재번호상태', '발주처명']]

    df_cost = pd.read_excel(as_cost_file, skiprows=[1])
    df_cost = df_cost[(df_cost['AS구분'] != '무상') & df_cost['AS구분'].notna()]
    df_cost = df_cost[~df_cost['진행상태'].isin(['접수취소', '최종완료'])]
    df_cost = df_cost[df_cost['입금상태'] != '입금완료']

    df_cost = df_cost.merge(status_map, how='inner', left_on='AS접수번호', right_index=True)

    def assign_category(row):
        if row['청구상태'] == '청구완료':
            return '청구'
        elif row['진행상태'] in ['기술적종료', '공사완료'] and row['청구상태'] in ['미청구', '부분청구']:
            return '미청구'
        elif row['진행상태'] in ['접수', '조치중']:
            return 'AS진행'
        else:
            return '미구분'

    df_cost['구분'] = df_cost.apply(assign_category, axis=1)

    df_cost['인보이스발행일자'] = df_cost.apply(
        lambda row: row['청구일자'] if pd.isna(row['인보이스발행일자']) and row['구분'] == '청구' else row['인보이스발행일자'],
        axis=1
    )

    def assign_type(row):
        if row['구분'] == '청구' and pd.notna(row['인보이스발행일자']):
            days = (today - pd.to_datetime(row['인보이스발행일자'])).days
            if days <= 30:
                return '정상(미입금)'
            elif days <= 60:
                return '입금지연_30일 경과'
            elif days <= 90:
                return '입금지연_60일 경과'
            elif days <= 120:
                return '입금지연_90일 경과'
            else:
                return '입금지연_120일 경과'
        elif row['구분'] == '미청구' and pd.notna(row['기술적종료일자']):
            days = (today - pd.to_datetime(row['기술적종료일자'])).days
            return '정상(미청구)' if days <= 60 else '청구지연'
        elif row['구분'] == 'AS진행':
            days = (today - pd.to_datetime(row['접수일자'])).days
            if row['진행상태'] == '조치중' or days <= 180:
                return 'AS조치중'
            else:
                return '조치지연'
        return None

    df_cost['유형'] = df_cost.apply(assign_type, axis=1)

    def calc_balance(row):
        if row['유형'] in ['입금지연_30일 경과', '입금지연_60일 경과', '입금지연_90일 경과', '입금지연_120일 경과']:
            return row['청구금액(원화)'] - row['입금액(원화)']
        return 0

    df_cost['미입금잔액'] = df_cost.apply(calc_balance, axis=1)

    order = ['AS진행', '미청구', '청구', '합계']
    type_order = [
        'AS조치중', '조치지연', '정상(미청구)', '청구지연',
        '정상(미입금)', '입금지연_30일 경과', '입금지연_60일 경과', '입금지연_90일 경과', '입금지연_120일 경과', '-'
    ]

    pivot = df_cost.pivot_table(
        index=['구분', '유형'],
        values=['AS접수번호', '미입금잔액', '도급금(원화)', '청구금액(원화)'],
        aggfunc={'AS접수번호': 'count', '미입금잔액': 'sum', '도급금(원화)': 'sum', '청구금액(원화)': 'sum'},
        fill_value=0
    ).reset_index()

    total_row = pd.DataFrame({
        '구분': ['합계'],
        '유형': ['-'],
        'AS접수번호': [pivot['AS접수번호'].sum()],
        '미입금잔액': [pivot['미입금잔액'].sum()],
        '도급금(원화)': [pivot['도급금(원화)'].sum()],
        '청구금액(원화)': [pivot['청구금액(원화)'].sum()]
    })

    pivot = pd.concat([pivot, total_row], ignore_index=True)
    pivot['구분'] = pd.Categorical(pivot['구분'], categories=order, ordered=True)
    pivot['유형'] = pd.Categorical(pivot['유형'], categories=type_order, ordered=True)
    pivot = pivot.sort_values(['구분', '유형']).reset_index(drop=True)

    formatted = pivot.copy()
    for col in ['AS접수번호', '도급금(원화)', '미입금잔액', '청구금액(원화)']:
        formatted[col] = formatted[col].apply(lambda x: f"{int(x):,}")

    st.subheader("AS채권현황 요약 집계표")
    st.dataframe(formatted)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_cost.to_excel(writer, index=False, sheet_name='AS채권현황 점검 결과')
        pivot.to_excel(writer, index=False, sheet_name='AS채권현황 요약 집계표')
    st.download_button(
        label="AS채권현황 점검 결과 다운로드",
        data=output.getvalue(),
        file_name="AS채권현황_점검결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 차트 생성
    chart_definitions = {
        'AS조치중': '도급금(원화)',
        '조치지연': '도급금(원화)',
        '정상(미청구)': '도급금(원화)',
        '청구지연': '도급금(원화)',
        '정상(미입금)': '청구금액(원화)'
    }

    chart_data = []
    for k, v in chart_definitions.items():
        row = pivot[pivot['유형'] == k]
        if not row.empty:
            chart_data.append({
                '유형': k,
                '건수': int(row['AS접수번호'].values[0]),
                '금액': int(row[v].values[0])
            })

    delay_types = ['입금지연_30일 경과', '입금지연_60일 경과', '입금지연_90일 경과', '입금지연_120일 경과']
    delay_rows = pivot[pivot['유형'].isin(delay_types)]
    if not delay_rows.empty:
        chart_data.append({
            '유형': '입금지연합계',
            '건수': int(delay_rows['AS접수번호'].sum()),
            '금액': int(delay_rows['미입금잔액'].sum())
        })

    chart_df = pd.DataFrame(chart_data)
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=chart_df['유형'],
        y=chart_df['건수'],
        name='건수',
        text=[f"{x:,}" for x in chart_df['건수']],
        textposition='outside',
        marker_color='deepskyblue'
    ))
    fig.add_trace(go.Bar(
        x=chart_df['유형'],
        y=chart_df['금액'],
        name='금액',
        text=[f"{x:,}" for x in chart_df['금액']],
        textposition='outside',
        marker_color='lightblue'
    ))

    fig.update_layout(
        title='AS채권현황 요약 차트 (요약 기준)',
        barmode='group',
        xaxis_title='유형',
        yaxis_title='합계',
        uniformtext_minsize=8,
        uniformtext_mode='hide',
        margin=dict(l=40, r=40, t=60, b=120)
    )
    st.plotly_chart(fig, use_container_width=True)
