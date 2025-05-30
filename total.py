import streamlit as st

# 🔁 기능별 파일 import
import as_analysis_test
import AS_SALES
import accounts
import AS_PROCESS
import project
import PRPO
import AS_summary
import accounts_summary

# ✅ 페이지 설정
st.set_page_config(page_title="AS 통합 분석 시스템", layout="wide")

# ✅ 타이틀
st.markdown("""
    <h1 style='text-align:center;'>📊 AS 통합 분석 시스템</h1>
    <p style='text-align:center; font-size:18px;'>아래에서 원하는 분석 기능을 선택하세요.</p>
    <hr style="border: 1px solid #eee;">
""", unsafe_allow_html=True)

# ✅ 기능 이름과 연결할 app 함수 매핑
app_list = {
    "🔍 AS상태 업데이트 대상 및 결산마감 대상 선정 시스템": as_analysis_test.app,
    "💰 유상 AS 매출 집계 시스템": AS_SALES.app,
    "📂 미수채권(미입금) 집계 시스템": accounts.app,
    "📈 AS 처리율 계산 시스템": AS_PROCESS.app,
    "🗂️ AS프로젝트 대상 선정 시스템": project.app,
    "📝 발주 및 입고지연 집계 시스템": PRPO.app,
    "📊 AS 접수/조치/조치일 집계 시스템": AS_summary.app,
    "📋 AS채권현황 분석 및 점검 시스템": accounts_summary.app
}

# ✅ 선택 박스 (radio를 사용해 명확한 선택 UI)
selected_app = st.radio("👇 실행할 기능을 선택하세요:", list(app_list.keys()))

# ✅ 선택된 기능 실행
try:
    app_list[selected_app]()  # 선택된 기능 함수 실행
except Exception as e:
    st.error(f"🚨 앱 실행 중 오류가 발생했습니다:\n\n{e}")
