import streamlit as st
import subprocess
import os

# ✅ 가장 위에 배치해야 함!
st.set_page_config(page_title="AS 통합 분석 시스템", layout="centered")

# 🎨 타이틀 영역
st.markdown("""
    <h1 style='text-align:center; color:#2c3e50;'>📊 AS 통합 분석 시스템</h1>
    <p style='text-align:center; font-size:18px;'>아래 버튼을 눌러 원하는 분석 기능을 실행하세요.</p>
    <hr style="border: 1px solid #eee;">
""", unsafe_allow_html=True)

# 버튼 스타일 정의
button_style = """
    <style>
    div.stButton > button {
        background-color: #f0f8ff;
        color: blue;
        font-size: 20px;
        height: 60px;
        width: 100%;
        border: 1px solid #d0d0d0;
        border-radius: 8px;
        white-space: pre-line;  /* 줄바꿈 적용 */
    }
    </style>
"""
st.markdown(button_style, unsafe_allow_html=True)

# 📌 첫 번째 줄 버튼 UI
col1, col2 = st.columns(2)

with col1:
    if st.button("🔍 AS상태 업데이트 대상 및\n결산마감 대상 선정 시스템"):
        st.info("새 창에서 'AS 상태 업데이트 대상 선정 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "as_analysis_test.py"], shell=True)

with col2:
    if st.button("💰 유상 AS 매출 집계 시스템"):
        st.info("새 창에서 '유상 AS 매출 집계 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "AS_SALES.py"], shell=True)

# 📌 두 번째 줄 버튼 UI
col3, col4 = st.columns(2)

with col3:
    if st.button("📂 미수채권(미입금) 집계 시스템"):
        st.info("새 창에서 '미수채권(미입금) 집계 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "accounts.py"], shell=True)

with col4:
    if st.button("📈 AS 처리율 계산 시스템"):
        st.info("새 창에서 'AS 처리율 계산 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "AS_PROCESS.py"], shell=True)

# 📌 세 번째 줄 버튼 UI
col5, col6 = st.columns(2)

with col5:
    if st.button("🗂️ AS프로젝트 대상 선정 시스템"):
        st.info("새 창에서 'AS프로젝트 대상 선정 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "project.py"], shell=True)

with col6:
    if st.button("📝 발주 및 입고지연 집계 시스템"):
        st.info("새 창에서 '발주 및 입고지연 집계 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "PRPO.py"], shell=True)

# 📌 네 번째 줄 버튼 UI
col7, col8 = st.columns(2)

with col7:
    if st.button("📊 AS 접수/조치/조치일 집계 시스템"):
        st.info("새 창에서 'AS 접수/조치/조치일 집계 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "AS_summary.py"], shell=True)

with col8:
    if st.button("📋 AS채권현황 분석 및 점검 시스템"):
        st.info("새 창에서 'AS채권현황 분석 및 점검 시스템' 앱을 실행합니다.")
        subprocess.Popen(["streamlit", "run", "accounts_summary.py"], shell=True)

# ℹ️ 안내 메시지
st.markdown("""
<hr>
<p style="text-align:center; color:gray;">선택한 기능은 새 창에서 실행되며, 이 페이지는 계속 열려 있습니다.</p>
""", unsafe_allow_html=True)
