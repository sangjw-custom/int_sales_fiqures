import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import io

# 1. Firebase 설정 (최초 1회)
if not firebase_admin._apps:
    cred = credentials.Certificate("int-sales-figures_01.json")
    firebase_admin.initialize_app(cred)
db = firestore.client()

st.title("📂 손익 데이터 업로드 시스템")
st.markdown("13만 건의 대용량 엑셀 데이터를 클라우드 DB로 전송합니다.")

# 2. 파일 업로드 UI
uploaded_file = st.file_uploader("엑셀 파일을 선택하세요 (.xlsb 또는 .xlsx)", type=["xlsb", "xlsx"])

if uploaded_file is not None:
    try:
        # 대용량 처리를 위해 chunk 단위가 아닌 메모리 직접 로드 시도
        with st.spinner('데이터를 분석 중입니다...'):
            df = pd.read_excel(uploaded_file, engine='pyxlsb' if 'xlsb' in uploaded_file.name else None)
            
            # 데이터 전처리 (불필요한 행/열 제거 및 결측치 처리)
            df = df.where(pd.notnull(df), None)
            total_rows = len(df)
            st.success(f"✅ 분석 완료: 총 {total_rows:,}행 데이터를 확인했습니다.")

        # 3. 전송 버튼
        if st.button("🚀 Firestore로 전송 시작"):
            batch = db.batch()
            progress_bar = st.progress(0)
            
            for index, row in df.iterrows():
                doc_ref = db.collection('int-sales-figures').document()
                batch.set(doc_ref, row.to_dict())
                
                # 500개 단위로 묶어서 전송 (성능 최적화)
                if (index + 1) % 500 == 0:
                    batch.commit()
                    batch = db.batch()
                    progress_bar.progress((index + 1) / total_rows)
            
            batch.commit() # 남은 데이터 전송
            st.success("✨ 모든 데이터가 성공적으로 업로드되었습니다!")

    except Exception as e:
        st.error(f"오류 발생: {e}")
