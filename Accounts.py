import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import plotly.express as px
import xlsxwriter

# ---------------------- Helper Functions ---------------------- #
def clean_column_names(columns):
    return [col.replace('\n', '').strip() for col in columns]

def calculate_overdue_days(row, today):
    if row['통화'] in ['USD', 'EUR']:
        base_date = row['INVOICE발행일자'] if pd.notnull(row['INVOICE발행일자']) else row['청구일자']
    else:
        base_date = row['청구일자']
    if pd.notnull(base_date):
        return (today - base_date).days
    return None

def classify_product_group(row):
    prod1 = str(row['제품군(1)']).strip()
    prod2 = str(row['제품군(2)']).strip()

    if prod1 in ['가스솔루션', '설비제어', '중단사업']:
        return '설비제어'
    elif prod1 == '평형수처리':
        return 'BWMS'
    elif prod1 == '배전반':
        return '배전반'
    elif prod1 in ['필드 값 없음', 'nan', 'NaN', 'None', ''] or prod1.lower() in ['nan', 'none']:
        if prod2 in ['IAS', 'ICMS', 'MAPS', '발전기모터', '제어기타', '항해제어']:
            return '설비제어'
        elif prod2 in ['BWMS', 'OFFSHORE']:
            return 'BWMS'
        elif prod2 in ['저압', '고압']:
            return '배전반'
        elif prod2 == 'A/S':
            return None
        else:
            return None
    return None

def filter_data(df):
    df = df.rename(columns={'AS구분명': 'AS구분'})
    df = df[df['AS구분'].isin(['유상', '단품판매', '위탁AS'])]
    df = df[(df['청구상태'] == '청구완료') & (df['입금상태'].isin(['미입금', '부분입금']))]
    return df

def process_dates(df):
    df['청구일자'] = pd.to_datetime(df['청구일자'], errors='coerce')
    df['INVOICE발행일자'] = pd.to_datetime(df['INVOICE발행일자'], errors='coerce')
    return df

def calculate_summary(df):
    summary_data = []
    currencies = df['통화'].unique()
    for cur in currencies:
        sub = df[df['통화'] == cur]
        if cur in ['USD', 'EUR']:
            row = [cur] + [
                round(sub['청구금액(통화)'].sum(), 2),
                round(sub['입금총액(통화)'].sum(), 2),
                round(sub['미입금잔액(통화)'].sum(), 2),
                0, 0, 0
            ]
        else:
            row = [cur] + [
                0, 0, 0,
                round(sub['청구금액(원화)'].sum(), 2),
                round(sub['입금총액(원화)'].sum(), 2),
                round(sub['미입금잔액(원화)'].sum(), 2)
            ]
        summary_data.append(row)
    return pd.DataFrame(summary_data, columns=['통화', '청구금액(통화)', '입금총액(통화)', '미입금잔액(통화)', '청구금액(원화)', '입금총액(원화)', '미입금잔액(원화)'])

def create_interactive_chart(df, currency, amount_column):
    filtered = df[df['통화'] == currency]
    agg = filtered.groupby('발주처명')[amount_column].sum().sort_values(ascending=False).head(20)

    if agg.empty:
        return None

    chart_df = agg.reset_index()
    fig = px.bar(
        chart_df,
        x='발주처명',
        y=amount_column,
        title=f"{currency} 기준 발주처별 미입금잔액 (상위 20개)",
        labels={'발주처명': '발주처명', amount_column: '미입금잔액'},
        text_auto='.2s',
    )
    fig.update_layout(
        xaxis_tickangle=-45,
        margin=dict(l=40, r=40, t=60, b=120),
        height=500,
        font=dict(size=10),
    )
    return fig

def to_excel(dataframe, summary_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, sheet_name='미수금 현황', index=False)
        worksheet = writer.sheets['미수금 현황']
        for i, width in enumerate([20]*len(dataframe.columns)):
            worksheet.set_column(i, i, width)

        summary_df.to_excel(writer, sheet_name='통화별 요약', index=False)
        worksheet2 = writer.sheets['통화별 요약']
        for i, width in enumerate([20]*summary_df.shape[1]):
            worksheet2.set_column(i, i, width)
    output.seek(0)
    return output

# ---------------------- Streamlit UI ---------------------- #
st.set_page_config(page_title="미수채권 분석 및 관리 시스템", layout="wide")

st.markdown(
    """
    <h1 style='display: inline;'>📊 미수채권 분석 및 관리 시스템</h1>
    <span style='color: red; font-size: 30px;'>
        ※ 업로드할 파일은 ERP의 
        <span style="color: blue;"><u>'채권관리'</u></span> 메뉴의 
        <span style="color: blue;"><u>'대금청구현황'</u></span>에서 다운 받은 파일을 업로드하세요!
    </span>
    """,
    unsafe_allow_html=True
)

st.markdown("""---  
**사용 방법**  
1. 미수금 데이터가 포함된 `.xlsx` 파일을 업로드하세요.  
2. 담당자 및 제품군을 선택하거나 전체 데이터를 분석하세요.  
3. 30/60/90/120일 이상 경과된 채권도 필터링할 수 있어요.  
""")

uploaded_file = st.file_uploader("Excel 파일 업로드", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
    except Exception as e:
        st.error(
            f"""
            ❌ 엑셀 파일을 열 수 없습니다.  
            🔒 DRM(디지털 권한 관리)으로 보호된 파일일 수 있습니다.  
            👉 오류 내용: {e}
            """
        )
        st.stop()

    df.columns = clean_column_names(df.columns)
    df = process_dates(df)
    df = filter_data(df)

    df['제품군'] = df.apply(classify_product_group, axis=1)
    df = df[df['제품군'].notna()]  # 제품군 분류 불가 항목 제외

    today = datetime.datetime.today()
    df['입금지연일수'] = df.apply(lambda row: calculate_overdue_days(row, today), axis=1)

    담당자_list = df['접수담당자'].dropna().unique().tolist()
    담당자_list.insert(0, '전체')
    selected_user = st.selectbox("담당자 선택", 담당자_list)

    product_group_list = sorted(df['제품군'].unique().tolist())
    product_group_list.insert(0, '전체')
    selected_group = st.selectbox("제품군 선택", product_group_list)

    overdue_days = st.selectbox("경과일 필터", ['전체', '30일 이상', '60일 이상', '90일 이상', '120일 이상'])

    df_filtered = df.copy()
    if selected_user != '전체':
        df_filtered = df_filtered[df_filtered['접수담당자'] == selected_user]

    if selected_group != '전체':
        df_filtered = df_filtered[df_filtered['제품군'] == selected_group]

    if overdue_days != '전체':
        threshold = int(overdue_days.replace('일 이상', ''))
        df_filtered = df_filtered[df_filtered['입금지연일수'] >= threshold]

    df_filtered = df_filtered[[ 
        'AS접수번호', '접수상태', 'INVOICE발행일자', '청구일자', '입금지연일수', '청구상태', '입금상태',
        '접수담당자', '발주처명', '통화', '도급금(통화)', '청구금액(통화)', '입금총액(통화)', '미입금잔액(통화)',
        '도급금(원화)', '청구금액(원화)', '입금총액(원화)', '미입금잔액(원화)', '판매구분', 'AS구분', '제목',
        '제품군(1)', '제품군(2)', '제품군'
    ]]

    summary_df = calculate_summary(df_filtered)

    st.markdown("---")
    st.subheader("📉 발주처별 미입금잔액 인터랙티브 그래프")

    st.markdown("### 💵 USD 기준")
    usd_fig = create_interactive_chart(df_filtered, 'USD', '미입금잔액(통화)')
    if usd_fig:
        st.plotly_chart(usd_fig, use_container_width=True)
    else:
        st.info("USD 기준 데이터가 없습니다.")

    st.markdown("### 🇰🇷 원화(KRW) 기준")
    krw_fig = create_interactive_chart(df_filtered, 'KRW', '미입금잔액(원화)')
    if krw_fig:
        st.plotly_chart(krw_fig, use_container_width=True)
    else:
        st.info("KRW 기준 데이터가 없습니다.")

    excel_file = to_excel(df_filtered, summary_df)

    st.success(f"분석 완료! 총 {len(df_filtered)}건의 미수채권이 확인되었습니다.")
    st.download_button(
        label="📥 미수채권 분석 결과 다운로드 (Excel 포함)",
        data=excel_file,
        file_name="미수금_현황_분석.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
