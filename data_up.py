import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd

# 1. Firebase 초기화 (중복 실행 방지)
if not firebase_admin._apps:
    # Streamlit Cloud의 Secrets 기능을 사용하는 것이 보안상 가장 좋으나,
    # 우선 기존 방식인 json 파일 경로로 설정합니다.
    try:
        cred = credentials.Certificate("int-sales-figures_01.json")
        firebase_admin.initialize_app(cred)
    except Exception as e:
        st.error(f"인증 파일(my_key.json)을 찾을 수 없거나 오류가 발생했습니다: {e}")

db = firestore.client()

st.set_page_config(page_title="Sales Data Uploader", layout="centered")
st.title("📊 영업 실적 데이터 업로드")
st.info("데이터는 Firestore의 **'int-sales-figures'** 컬렉션에 저장됩니다.")

# 2. 파일 업로드 UI
uploaded_file = st.file_uploader("엑셀 파일을 선택하세요 (.xlsb, .xlsx)", type=["xlsb", "xlsx"])

if uploaded_file is not None:
    # 사용자에게 진행 여부 확인
    if st.button("🚀 int-sales-figures에 업로드 시작"):
        try:
            with st.spinner('데이터를 처리하고 클라우드로 전송 중입니다...'):
                # 파일 확장자에 따른 엔진 설정
                if uploaded_file.name.endswith('xlsb'):
                    df = pd.read_excel(uploaded_file, engine='pyxlsb')
                else:
                    df = pd.read_excel(uploaded_file)
                
                # 데이터 전처리: 결측치(NaN)를 None으로 변환해야 Firestore 업로드 가능
                df = df.where(pd.notnull(df), None)

                # Firestore 컬렉션 지정
                collection_ref = db.collection('int-sales-figures')
                
                # 대량 업로드를 위한 Batch 작업 (500개 단위)
                batch = db.batch()
                total_rows = len(df)
                
                for index, row in df.iterrows():
                    # 각 행을 딕셔너리로 변환하여 문서 추가
                    doc_ref = collection_ref.document()
                    batch.set(doc_ref, row.to_dict())
                    
                    # 500개마다 커밋
                    if (index + 1) % 500 == 0:
                        batch.commit()
                        batch = db.batch()
                
                # 남은 데이터 커밋
                batch.commit()
                
                st.success(f"✅ 성공! 총 {total_rows}건의 데이터가 'int-sales-figures'에 저장되었습니다.")
                st.balloons() # 축하 효과
                
        except Exception as e:
            st.error(f"업로드 중 오류가 발생했습니다: {e}")
            st.write("상세 에러 내역:", e)

# 하단 도움말
with st.expander("도움말 및 주의사항"):
    st.write("- 파일 내 보안(DRM)이 적용된 경우 읽기 오류가 날 수 있습니다.")
    st.write("- 동일한 파일을 여러 번 올리면 데이터가 중복 생성되니 주의하세요.")
