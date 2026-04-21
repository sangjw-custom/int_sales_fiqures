import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import requests
import io

# 1. Firebase 초기화 (Secrets 활용 권장)
if not firebase_admin._apps:
    # 팁: 보안을 위해 my_key.json 내용을 Streamlit Secrets에 넣는 것을 추천합니다.
    cred = credentials.Certificate("int-sales-figures_01.json")
    firebase_admin.initialize_app(cred)
db = firestore.client()

st.title("☁️ 클라우드 데이터 전송 시스템")
st.info("MS OneDrive/SharePoint의 엑셀 주소를 이용해 Firestore로 직접 업로드합니다.")

# 2. MS 엑셀 파일 URL 입력
# 공유 링크 예시: https://company-my.sharepoint.com/:x:/g/personal/.../file.xlsb?download=1
excel_url = st.text_input(https://kccglass01.sharepoint.com/:x:/s/TORG_A0117D0/IQCytyktOmYCTprk8pZb3Pz6AZ-_H7c6bD2nNC9YfVoV348?e=XG5kwT ?download=1 포함)", placeholder="https://...")

if st.button("🚀 클라우드 데이터 동기화 시작"):
    if not excel_url:
        st.warning("링크를 입력해주세요.")
    else:
        try:
            with st.spinner('클라우드에서 데이터를 가져오는 중...'):
                # 1. MS 클라우드에서 파일 읽기
                response = requests.get(excel_url)
                response.raise_for_status()
                
                # 2. 판다스로 변환 (xlsb 엔진 사용)
                file_content = io.BytesIO(response.content)
                if excel_url.contains(".xlsb"):
                    df = pd.read_excel(file_content, engine='pyxlsb')
                else:
                    df = pd.read_excel(file_content)
                
                df = df.where(pd.notnull(df), None)
                total_rows = len(df)
                
                st.write(f"📊 총 {total_rows:,}건의 데이터를 확인했습니다. 업로드를 시작합니다.")

                # 3. Firestore 배치 업로드 (13만 건 대응)
                batch = db.batch()
                progress_bar = st.progress(0)
                
                for index, row in df.iterrows():
                    doc_ref = db.collection('int-sales-figures').document()
                    batch.set(doc_ref, row.to_dict())
                    
                    if (index + 1) % 500 == 0:
                        batch.commit()
                        batch = db.batch()
                        progress_bar.progress((index + 1) / total_rows)
                
                batch.commit() # 남은 데이터
                st.success(f"✨ 완료! {total_rows:,}건이 성공적으로 저장되었습니다.")
                
        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.write("힌트: 링크가 직접 다운로드 가능한 형태인지 확인하세요.")
