import streamlit as st
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import requests
import io

# 1. Firebase 설정
if not firebase_admin._apps:
    cred = credentials.Certificate("int-sales-figures_01.json")
    firebase_admin.initialize_app(cred)
db = firestore.client()

st.title("📊 손익 데이터 클라우드 동기화")

# 2. 셰어포인트 경로 입력 (보안 문서 직접 읽기용)
st.info("사내 보안(DRM) 문제를 피하기 위해 셰어포인트 직접 링크를 사용합니다.")
sharepoint_url = st.text_input("https://kccglass01.sharepoint.com/:x:/s/TORG_A0117D0/IQAvkAvPyIo9TqL78ZNaoYxnAWaQ8xx8O30KmUgsBaQ067k?e=NHb5lz?download=1")

if st.button("🚀 데이터 전송 시작"):
    if not sharepoint_url:
        st.warning("링크를 입력해주세요.")
    else:
        try:
            with st.spinner('클라우드에서 13만 건 데이터를 로드 중...'):
                # 셰어포인트에서 데이터 바로 읽기 (로컬 DRM 우회)
                response = requests.get(sharepoint_url)
                response.raise_for_status()
                
                # 메모리 내에서 엑셀 변환
                excel_file = io.BytesIO(response.content)
                # xlsb라면 engine='pyxlsb' 추가, 일반 xlsx라면 생략
                df = pd.read_excel(excel_file, engine='pyxlsb') 
                
                # 데이터 정제 (NaN 제거 및 필요한 컬럼만 필터링)
                df = df.where(pd.notnull(df), None)
                total_rows = len(df)
                
            # 3. Firestore 배치 업로드
            batch = db.batch()
            progress_bar = st.progress(0)
            
            for index, row in df.iterrows():
                # 위에서 설계한 테이블 구조대로 저장
                doc_ref = db.collection('int-sales-figures').document()
                batch.set(doc_ref, row.to_dict())
                
                if (index + 1) % 500 == 0:
                    batch.commit()
                    batch = db.batch()
                    progress_bar.progress((index + 1) / total_rows)
            
            batch.commit()
            st.success(f"✨ 완료! {total_rows:,}건의 데이터가 업로드되었습니다.")

        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.write("힌트: 셰어포인트 링크 권한이 '링크가 있는 모든 사용자'로 되어 있는지 확인하세요.")
