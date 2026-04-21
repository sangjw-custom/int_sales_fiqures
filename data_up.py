import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd

# 1. Firebase 초기화 (중복 실행 방지)
if not firebase_admin._apps:
    # 서비스 계정 키 설정 (Secrets 활용 권장)
    cred = credentials.Certificate("int-sales-figures_01.json")
    firebase_admin.initialize_app(cred)
db = firestore.client()

st.title("📊 영업 실적 데이터 업로드")

# 2. 웹 브라우저용 파일 업로드 UI (tkinter 대체)
uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=["xlsb", "xlsx"])

if uploaded_file is not None:
    collection_name = st.text_input("저장할 컬렉션 이름", value="Sales_Profit_202603")
    
    if st.button("클라우드 업로드 시작"):
        try:
            with st.spinner('데이터 처리 중...'):
                # 파일 확장자에 따른 처리
                if uploaded_file.name.endswith('xlsb'):
                    df = pd.read_excel(uploaded_file, engine='pyxlsb')
                else:
                    df = pd.read_excel(uploaded_file)
                
                df = df.where(pd.notnull(df), None) # 빈값 처리

                # 배치 업로드 (최대 500개씩)
                batch = db.batch()
                for index, row in df.iterrows():
                    doc_ref = db.collection(collection_name).document()
                    batch.set(doc_ref, row.to_dict())
                    
                    if (index + 1) % 500 == 0:
                        batch.commit()
                        batch = db.batch()
                
                batch.commit()
                st.success(f"✅ {len(df)}건 업로드 완료!")
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
