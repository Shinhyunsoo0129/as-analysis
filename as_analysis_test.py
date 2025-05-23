import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="AS 상태 업데이트 및 결산 마감 대상 선정", layout="wide")

st.markdown(
    "<h1 style='display: inline;'>📂 AS 상태 업데이트 및 결산 마감 대상 선정</h1> "
    "<br><span style='color: red; font-size: 30px;'>※ 업로드할 파일은 ERP의 "
    "<span style='color: blue;'><u>'AS현황 및 최종완료'</u></span>에서 다운 받은 파일을 업로드하세요!</span>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("✅ 분석할 Excel 파일을 업로드하세요 (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # ✅ DRM 방지를 위한 예외 처리
        df = pd.read_excel(uploaded_file, engine='openpyxl', header=[0, 1])
    except Exception as e:
        st.error(
            f"""
            ❌ 엑셀 파일을 열 수 없습니다.  
            🔒 DRM(디지털 권한 관리)으로 보호된 파일일 수 있습니다.  
            👉 오류 내용: {e}
            """
        )
        st.stop()

    # 1. 멀티인덱스 컬럼 합치기
    df.columns = ['_'.join([str(i).strip() for i in col if pd.notna(i)]) for col in df.columns]

    # 2. 컬럼명 정리
    rename_dict = {
        'AS접수번호_AS접수번호': 'AS접수번호',
        '제목_제목': '제목',
        'AS접수일자_AS접수일자': 'AS접수일자',
        '인보이스발행일자_인보이스발행일자': '인보이스발행일자',
        '전자결재번호상태_전자결재번호상태': '전자결재번호상태',
        'AS진행상태_AS진행상태': 'AS진행상태',
        'AS구분_AS구분': 'AS구분',
        '청구상태_청구상태': '청구상태',
        '입금상태_입금상태': '입금상태',
        '접수담당자_접수담당자': '접수담당자',
        '접수정보_투입자재계획': '투입자재계획',
        '접수정보_외주계획': '외주계획',
        '접수정보_기타계획': '기타계획',
        '접수정보_출장계획': '출장계획',
        '조치내역_투입자재계획': '투입자재계획.1',
        '조치내역_외주계획': '외주계획.1',
        '조치내역_기타계획': '기타계획.1',
        '조치내역_출장계획': '출장계획.1',
    }
    df = df.rename(columns=rename_dict)

    # 3. 필수 컬럼 체크
    required_columns = ['전자결재번호상태', 'AS진행상태', 'AS구분', '청구상태', '입금상태']
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.error(f"❗ 필수 컬럼 누락: {missing_cols}")
        st.stop()

    # 4. 기본 필터링
    df = df[df['전자결재번호상태'] == '종결']
    df = df[df['AS진행상태'].isin(['접수', '조치중', '기술적종료'])]
    df = df[
        (
            df['AS구분'].isin(['유상', '위탁AS', '단품판매']) & (df['청구상태'] == '청구완료')
        ) | (
            df['AS구분'] == '무상'
        )
    ]

    # 5. 복합 조건 필터링
    def match_rule3(row):
        def is_O(*cols): return all(row.get(col) == 'O' for col in cols)
        def is_X(*cols): return all(row.get(col) == 'X' for col in cols)
        return any([
            is_O('투입자재계획', '투입자재계획.1') and is_X('외주계획', '기타계획', '출장계획', '외주계획.1'),
            is_O('투입자재계획', '기타계획', '투입자재계획.1') and is_X('외주계획', '출장계획', '외주계획.1'),
            is_O('외주계획', '외주계획.1') and is_X('투입자재계획', '기타계획', '출장계획', '투입자재계획.1'),
            row.get('기타계획') == 'O' and is_X('투입자재계획', '외주계획', '출장계획', '투입자재계획.1', '외주계획.1'),
            is_O('기타계획', '출장계획') and is_X('투입자재계획', '외주계획', '투입자재계획.1', '외주계획.1'),
            is_O('투입자재계획', '출장계획', '투입자재계획.1') and is_X('외주계획', '기타계획', '외주계획.1'),
            is_O('투입자재계획', '외주계획', '투입자재계획.1', '외주계획.1') and is_X('기타계획', '출장계획'),
            is_O('투입자재계획', '외주계획', '기타계획', '출장계획', '투입자재계획.1') and row.get('기타계획') == 'X',
            row.get('출장계획') == 'O' and is_X('투입자재계획', '외주계획', '기타계획', '투입자재계획.1', '외주계획.1'),
            is_O('외주계획', '출장계획', '외주계획.1') and is_X('투입자재계획', '기타계획', '투입자재계획.1'),
            is_O('외주계획', '기타계획', '외주계획.1') and is_X('투입자재계획', '출장계획', '투입자재계획.1'),
            is_O('투입자재계획', '외주계획', '기타계획', '출장계획', '투입자재계획.1', '외주계획.1'),
            is_X('투입자재계획', '외주계획', '기타계획', '출장계획', '투입자재계획.1', '외주계획.1'),
        ])
    df = df[df.apply(match_rule3, axis=1)]

    # 6. 점검사항 생성
    def generate_checklist(row):
        if row['AS구분'] == '무상':
            return "원가 투입 완료 여부 점검 / AS상태 업데이트 점검"
        elif row['AS구분'] in ['유상', '단품판매'] and row['입금상태'] == '입금완료':
            return "원가 투입 완료 여부 점검 / AS상태 업데이트 점검 / 최종완료 처리 점검"
        elif row['AS구분'] in ['유상', '단품판매'] and row['입금상태'] in ['미입금', '부분입금']:
            return "원가 투입 완료 여부 점검 / AS상태 업데이트 점검 / 공사완료 처리 점검"
        else:
            return ""

    df['점검사항'] = df.apply(generate_checklist, axis=1)

    # 7. 접수담당자 필터
    st.sidebar.header("🧑‍💼 접수담당자 선택")
    담당자_목록 = df['접수담당자'].dropna().unique().tolist()
    selected_person = st.sidebar.selectbox("접수담당자를 선택하세요", options=["전체"] + 담당자_목록)

    filtered_df = df if selected_person == "전체" else df[df['접수담당자'] == selected_person]

    # 8. 결과 출력 및 다운로드
    columns_to_save = [
        'AS접수번호', '제목', 'AS접수일자', '인보이스발행일자', 'AS구분', 'AS진행상태',
        '입금상태', '청구상태', '투입자재계획', '외주계획', '기타계획', '출장계획',
        '투입자재계획.1', '외주계획.1', '기타계획.1', '출장계획.1', '접수담당자', '점검사항'
    ]

    result_df = filtered_df[columns_to_save]

    st.success(f"✅ '{selected_person}' 데이터 {len(result_df)}건 필터링 완료")
    st.dataframe(result_df, use_container_width=True)

    def convert_df_to_excel(df):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return buffer.getvalue()

    excel_data = convert_df_to_excel(result_df)

    st.download_button(
        label="📥 결과 엑셀 다운로드",
        data=excel_data,
        file_name=f"{selected_person}_선정결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("★★다른 담당자의 결과값을 받기 위해서는 담당자의 이름을 다시 선택하세요.★★")
