import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import matplotlib.pyplot as plt
import matplotlib
import plotly.express as px
from datetime import datetime

# 한글 폰트 설정
matplotlib.rcParams['font.family'] = 'Malgun Gothic'
matplotlib.rcParams['axes.unicode_minus'] = False

st.set_page_config(page_title="AS 처리율 계산기", layout="wide")

# 제목 및 안내문
st.markdown(
    """
    <h1 style='display: inline;'>📊 AS 처리율 계산기</h1>
    <p style="color: red; font-size: 30px;">
    ※ 업로드할 파일은 ERP의 <span style="color: blue;"><u>'AS현황 및 최종확률'</u></span>에서 다운 받은 파일을 업로드하세요!<br>
    ※ 조회할 기간을 선택하고, <span style="color: blue;"><u>'처리율 분석 실행'</u></span> 버튼을 클릭하세요.
    </p>
    """,
    unsafe_allow_html=True
)

@st.cache_data(show_spinner=False)
def classify_product_group(row):
    if row['AS접수번호'] in ['AS23020137', 'AS22110268', '606746', '606366']:
        return '설비제어'
    group1 = row['제품군1']
    group2 = row['제품군2']
    if pd.notna(group1):
        if group1 in ['설비제어', '중단사업', '가스솔루션']:
            return '설비제어'
        elif group1 == '배전반':
            return '배전반'
        elif group1 in ['평형수처리', '연구개발']:
            return 'BWMS'
    if pd.notna(group2):
        if group2 in ['ICMS', 'MAPS', 'IAS', '제어기타', '항해제어', 'FGSS', 'OFFSHORE', '발전기모터', 'A/S']:
            return '설비제어'
        elif group2 in ['고압', '저압']:
            return '배전반'
        elif group2 == 'BWMS':
            return 'BWMS'
    return '기타'

@st.cache_data(show_spinner=False)
def classify_as구분(as구분):
    if as구분 in ['위탁AS', '무상']:
        return '무상'
    elif as구분 in ['유상', '단품판매']:
        return '유상'
    return '기타'

@st.cache_data(show_spinner=False)
def classify_진행상태(상태):
    if 상태 in ['접수', '조치중']:
        return '조치중'
    elif 상태 in ['기술적종료', '공사완료', '최종완료']:
        return '조치완료'
    return '기타'

def process_data(df, start_ym, end_ym):
    df = df.copy()
    df['AS접수일자'] = pd.to_datetime(df['AS접수일자'], format='%Y/%m/%d', errors='coerce')
    df['접수년월_dt'] = df['AS접수일자'].dt.to_period('M').dt.to_timestamp()
    df['접수년월'] = df['접수년월_dt'].dt.strftime('%m-%Y')

    df = df[
        (df['전자결재번호상태'] == '종결') &
        (df['AS진행상태'] != '접수취소') &
        (df['AS접수번호'] != 'AS21120297')
    ]
    df = df[(df['접수년월_dt'] >= start_ym) & (df['접수년월_dt'] <= end_ym)]

    df['제품군'] = df.apply(classify_product_group, axis=1)
    df['AS구분'] = df['AS구분'].apply(classify_as구분)
    df['진행상태분류'] = df['AS진행상태'].apply(classify_진행상태)

    df['AS접수건수'] = 1
    df['조치완료건수'] = df['진행상태분류'].apply(lambda x: 1 if x == '조치완료' else 0)

    result = df.groupby(['AS구분', '제품군', '접수년월', '접수년월_dt']).agg({
        'AS접수건수': 'sum',
        '조치완료건수': 'sum'
    }).reset_index()

    result['AS처리율'] = (result['조치완료건수'] / result['AS접수건수'] * 100).round(2)
    result = result.sort_values(['AS구분', '제품군', '접수년월_dt'])

    def make_summary_row(구분):
        temp = result[result['AS구분'] == 구분]
        접수합 = temp['AS접수건수'].sum()
        완료합 = temp['조치완료건수'].sum()
        return pd.DataFrame([{
            'AS구분': f'{구분} 합계',
            '제품군': '',
            '접수년월': '',
            '접수년월_dt': pd.NaT,
            'AS접수건수': 접수합,
            '조치완료건수': 완료합,
            'AS처리율': round(완료합 / 접수합 * 100, 2) if 접수합 > 0 else 0
        }])

    합계1 = make_summary_row('무상')
    합계2 = make_summary_row('유상')
    total접수 = result['AS접수건수'].sum()
    total완료 = result['조치완료건수'].sum()
    전체합계 = pd.DataFrame([{
        'AS구분': '전체 합계',
        '제품군': '',
        '접수년월': '',
        '접수년월_dt': pd.NaT,
        'AS접수건수': total접수,
        '조치완료건수': total완료,
        'AS처리율': round(total완료 / total접수 * 100, 2) if total접수 > 0 else 0
    }])

    result = pd.concat([result, 합계1, 합계2, 전체합계], ignore_index=True)

    graph_df = df.groupby(['제품군', '접수년월', '접수년월_dt']).agg(
        {'AS접수건수': 'sum', '조치완료건수': 'sum'}
    ).reset_index()
    graph_df['AS처리율'] = (graph_df['조치완료건수'] / graph_df['AS접수건수'] * 100).round(2)
    graph_df = graph_df.sort_values('접수년월_dt')

    return result, df, graph_df

def plot_interactive_chart(df):
    fig = px.bar(
        df, x='접수년월', y='AS처리율', color='제품군',
        barmode='group', text='AS처리율',
        title='월별 AS 처리율 (제품군별 합산 기준)'
    )
    fig.update_traces(texttemplate='%{text}%', textposition='outside')
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
    fig.update_yaxes(range=[0, 100])
    st.plotly_chart(fig, use_container_width=True)
    return fig

def save_chart_to_image(df):
    fig, ax = plt.subplots(figsize=(16, 8))
    제품군_목록 = df['제품군'].unique()
    x_labels = sorted(df['접수년월'].unique(), key=lambda x: pd.to_datetime('01-' + x, format='%d-%m-%Y'))
    width = 0.8 / len(제품군_목록)

    for i, 제품군 in enumerate(제품군_목록):
        sub_df = df[df['제품군'] == 제품군]
        grouped = sub_df.groupby('접수년월')['AS처리율'].mean().reindex(x_labels).fillna(0)
        x_pos = np.arange(len(x_labels)) + (i - len(제품군_목록)/2)*width + width/2
        bars = ax.bar(x_pos, grouped.values, width=width, label=제품군)
        for bar, height in zip(bars, grouped.values):
            if height > 0:
                ax.text(bar.get_x() + bar.get_width()/2, height + 1, f'{height:.1f}%',
                        ha='center', va='bottom', fontsize=9)

    ax.set_ylabel('AS 처리율 (%)')
    ax.set_xlabel('접수년월')
    ax.set_title('월별 AS 처리율 (엑셀용 이미지)')
    ax.set_xticks(np.arange(len(x_labels)))
    ax.set_xticklabels(x_labels, rotation=45, ha='right')
    ax.set_ylim(0, 110)
    ax.legend()
    plt.tight_layout()
    img_data = BytesIO()
    fig.savefig(img_data, format='png')
    img_data.seek(0)
    return img_data

def to_excel(result_df, original_df, image_data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.drop(columns='접수년월_dt').to_excel(writer, index=False, sheet_name='AS처리율')
        original_df.to_excel(writer, index=False, sheet_name='원본데이터')
        worksheet = writer.sheets['AS처리율']
        worksheet.insert_image('K2', 'chart.png', {'image_data': image_data})
    output.seek(0)
    return output

uploaded_file = st.file_uploader("📎 AS 데이터 엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, skiprows=1)  # Skip header row from ERP file
        df['AS접수일자'] = pd.to_datetime(df['AS접수일자'], format='%Y/%m/%d', errors='coerce')
    except Exception as e:
        st.error(
            f"""
            ❌ 엑셀 파일을 열 수 없습니다.  
            🔒 DRM(디지털 권한 관리)으로 보호된 파일일 수 있습니다.  
            👉 오류 내용: {e}
            """
        )
        st.stop()

    if 'AS접수일자' in df.columns:
        year_options = sorted(df['AS접수일자'].dt.year.dropna().astype(int).unique())
        month_options = list(range(1, 13))

        col1, col2 = st.columns(2)
        with col1:
            start_year = st.selectbox("시작 년도", year_options)
            start_month = st.selectbox("시작 월", month_options)
        with col2:
            end_year = st.selectbox("종료 년도", year_options, index=len(year_options)-1)
            end_month = st.selectbox("종료 월", month_options, index=11)

        if 'result_df' not in st.session_state:
            st.session_state.result_df = None
            st.session_state.filtered_df = None
            st.session_state.graph_df = None

        if st.button("📊 처리율 분석 실행"):
            try:
                start_ym = datetime(start_year, start_month, 1)
                end_ym = datetime(end_year, end_month, 28)
                result_df, filtered_df, graph_df = process_data(df, start_ym, end_ym)

                st.session_state.result_df = result_df
                st.session_state.filtered_df = filtered_df
                st.session_state.graph_df = graph_df

                st.success("✅ 처리 완료!")
            except Exception as e:
                st.error(f"❌ 오류 발생: {e}")

        if st.session_state.result_df is not None:
            st.dataframe(st.session_state.result_df.drop(columns='접수년월_dt'))
            if st.session_state.graph_df is not None:
                fig = plot_interactive_chart(st.session_state.graph_df)
                image_data = save_chart_to_image(st.session_state.graph_df)
                st.download_button(
                    label="📥 결과 엑셀 다운로드 (그래프 포함)",
                    data=to_excel(st.session_state.result_df, st.session_state.filtered_df, image_data),
                    file_name=f"AS처리율_{start_year}{start_month:02d}_{end_year}{end_month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.error("❗ 'AS접수일자' 컬럼이 파일에 존재하지 않습니다.")