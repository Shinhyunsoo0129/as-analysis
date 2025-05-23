import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
import io

st.title("AS프로젝트 대상 선정 시스템")

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # 줄바꿈 문자가 포함된 컬럼명 정규화
    df.columns = df.columns.str.replace("\r|\n", "", regex=True)

    # 1. '프로젝트상태' 필터
    exclude_status = ['계약취소', '프로젝트중단']
    df = df[~df['프로젝트상태'].isin(exclude_status)]

    # 2. '제품군(1)' 분류
    def map_product(value):
        if value == '공용':
            return None
        elif value == '평형수처리':
            return '평형수처리'
        elif value in ['배전반', '육상배전', '에너지솔루션']:
            return '배전반'
        elif value in ['설비제어', '가스솔루션', '중단사업']:
            return '설비제어'
        else:
            return value

    df['제품군(1)'] = df['제품군(1)'].apply(map_product)
    df = df[df['제품군(1)'].notna()]

    # 3. '계약구분재구분' 컬럼 생성 (원본 계약구분 유지)
    df['계약구분'] = df['계약구분']  # 원본 유지 명시적 처리
    def map_contract(value):
        if value == '해당없음':
            return None
        elif value in ['자체수주(삼성중공업 거제조선)', '자체수주(국내)', '자체수주(해외)']:
            return '자체수주'
        elif value in ['99', 'U', 'S']:
            return 'SHI수주'
        else:
            return value

    df['계약구분재구분'] = df['계약구분'].apply(map_contract)
    df = df[df['계약구분재구분'].notna()]

    # 4. '인도년', '인도년월' 생성
    def extract_delivery_date(row):
        date = row['인도일자'] if pd.notna(row['인도일자']) else row['인도예정일자']
        if pd.notna(date):
            if isinstance(date, str):
                try:
                    date = pd.to_datetime(date)
                except:
                    return pd.Series([None, None])
            return pd.Series([date.year, date.strftime('%Y-%m')])
        else:
            return pd.Series([None, None])

    df[['인도년', '인도년월']] = df.apply(extract_delivery_date, axis=1)

    # 5. 제외 대상 프로젝트명 필터
    exclude_keywords = [
        "PDP 3기 DCS 설치공사", "원격제어장비구입설치", "해운대 IO SPARE PART", "CPU & POWER UNIT",
        "경산시 하수종말처리장 메인컴퓨터 점검 및 업그레이드", "IPU노후교체", "PIC2 CARD",
        "수원연구소 U/T 제어 개보수", "BMS & Governer 시스템 개발",
        "2005년도 김해지사 분산제어 설비 보수공사", "PDP크린룸설치계장공사",
        "복합동4Line Scrubber DCS 공사", "ESW2 Backup Card",
        "인천수산정수장 중앙제어실 원격제어설비 보완 수리", "FPUS 배터리 납품",
        "Rack I/O Card 납품", "삼성SDI(천안) Spareparts 납품", "삼성SDI(천안) Spareparts(FPUS) 납품",
        "사급품 파손", "단락전류"
    ]
    exclude_contains = [
        "정산", "중단", "취소", "보완작업", "보완", "보수", "수정작업", "수정작", "수정",
        "수리 작업", "수정 건", "추가", "운송", "시설재", "Spare Part", "SparePart", "교체 작업",
        "화재 보수", "화재건", "BC 추", "STARTER 추", "BOARD 추", "BOARD추", "RPB 추", "장비 견적",
        "재제작", "S/W Modifi", "점검 및 업그레이드", "전원공급기", "scanning", "3D scan", "스캔",
        "C/O", "시운전", "수정작업", "sampling", "위탁 운영"
    ]

    mask_exclude_exact = df['프로젝트명'].isin(exclude_keywords)
    mask_exclude_partial = df['프로젝트명'].apply(lambda x: any(keyword in str(x) for keyword in exclude_contains))
    df = df[~(mask_exclude_exact | mask_exclude_partial)]

    # 6. '보증종료년' 계산
    def compute_warranty_year(row):
        end_date_val = row.get('최종수요처보증종료일')
        if pd.notna(end_date_val):
            try:
                return pd.to_datetime(end_date_val).year
            except:
                return None
        else:
            try:
                base_date = pd.to_datetime(row.get('인도예정일자'))
                months = row.get('최종수요처보증개월')
                if pd.notna(base_date) and pd.notna(months):
                    delta_days = int(months * 365 / 12)
                    estimated_date = base_date + timedelta(days=delta_days)
                    return estimated_date.year
            except:
                return None
        return None

    df['보증종료년'] = df.apply(compute_warranty_year, axis=1)

    # 7. 'AS구분' 컬럼 추가
    today = pd.to_datetime(datetime.today().date())
    def classify_as(row):
        try:
            end_date_val = row.get('최종수요처보증종료일')
            if pd.notna(end_date_val):
                end_date = pd.to_datetime(end_date_val)
            else:
                base_date = pd.to_datetime(row.get('인도예정일자'))
                months = row.get('최종수요처보증개월')
                if pd.notna(base_date) and pd.notna(months):
                    delta_days = int(months * 365 / 12)
                    end_date = base_date + timedelta(days=delta_days)
                else:
                    end_date = None
            if pd.notna(end_date) and end_date < today:
                return '유상'
            else:
                return '무상'
        except:
            return '무상'

    df['AS구분'] = df.apply(classify_as, axis=1)

    # 선정 프로젝트 수 표시
    st.success("선정된 프로젝트 수: {}건".format(len(df)))

    # 교차표 생성
    pivot_table = pd.pivot_table(df, index='인도년', columns='보증종료년', values='프로젝트명', aggfunc='count', fill_value=0)
    st.subheader("[인도년 vs 보증종료년 프로젝트 건수 요약표]")
    st.dataframe(pivot_table)

    # 다운로드용 Excel 준비 (2시트)
    @st.cache_data
    def convert_excel(df1, summary):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='선정 프로젝트 리스트', index=False)
            summary.to_excel(writer, sheet_name='프로젝트 건수 요약표')
        return output.getvalue()

    st.download_button(
        label="다운로드 (Excel)",
        data=convert_excel(df, pivot_table),
        file_name="AS_프로젝트_선정_결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )