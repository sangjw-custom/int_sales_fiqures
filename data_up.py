import win32com.client as win32
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import os
import time

# --- [사용자 설정 영역] ---
# 파일명을 요청하신 대로 수정하였습니다.
KEY_PATH = "int-sales-figures_01.json"  
# 엑셀 파일의 전체 경로를 입력하세요.
EXCEL_FILE_PATH = r"C:\경로를_수정하세요\▣ 영업본부 26년 3월 손익 DATA (260403)_상품.xlsb"
COLLECTION_NAME = "int-sales-figures"
# -----------------------

def initialize_firebase():
    if not firebase_admin._apps:
        # 변경된 키 파일명을 사용하여 인증합니다.
        try:
            cred = credentials.Certificate(KEY_PATH)
            firebase_admin.initialize_app(cred)
            print(f"✅ Firebase 인증 성공: {KEY_PATH}")
        except Exception as e:
            print(f"❌ 인증 파일 로드 실패: {e}")
            return None
    return firestore.client()

def read_secure_excel(file_path):
    if not os.path.exists(file_path):
        print(f"❌ 파일을 찾을 수 없습니다: {file_path}")
        return None

    print(f"📌 보안 문서 읽기 시도 (Excel 제어): {os.path.basename(file_path)}")
    # 윈도우 설치 엑셀 앱을 구동하여 DRM 권한을 활용
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False 
    
    try:
        wb = excel.Workbooks.Open(file_path)
        ws = wb.ActiveSheet
        
        # 데이터 영역 전체 로드 (메모리 효율을 위해 가급적 데이터가 있는 범위만 선택)
        raw_data = ws.UsedRange.Value
        
        # 데이터프레임 변환 (첫 줄은 제목 열)
        df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
        
        wb.Close(False)
        print(f"✅ 데이터 로드 완료: {len(df):,} 행 감지")
        return df
    except Exception as e:
        print(f"❌ 엑셀 읽기 오류: {e}")
        return None
    finally:
        excel.Quit()

def upload_to_firestore(df, collection_name):
    db = initialize_firebase()
    if db is None: return

    collection_ref = db.collection(collection_name)
    
    # NaN(빈값) 처리 (Firestore 전송 오류 방지)
    df = df.where(pd.notnull(df), None)
    
    batch = db.batch()
    total_rows = len(df)
    start_time = time.time()

    print(f"🚀 {total_rows:,}건의 데이터 업로드를 시작합니다...")

    for index, row in df.iterrows():
        doc_ref = collection_ref.document() # 자동 ID 문서 생성
        batch.set(doc_ref, row.to_dict())
        
        # 500개 단위 배치 커밋 (Firestore 제한 준수)
        if (index + 1) % 500 == 0:
            batch.commit()
            batch = db.batch()
            elapsed = time.time() - start_time
            print(f"📦 {index + 1:,} / {total_rows:,} 행 완료 ({elapsed:.1f}초 경과)")

    # 남은 잔여 데이터 최종 커밋
    batch.commit()
    print(f"\n✨ 업로드 성공! 모든 데이터가 '{collection_name}'에 저장되었습니다.")

if __name__ == "__main__":
    sales_df = read_secure_excel(EXCEL_FILE_PATH)
    
    if sales_df is not None:
        upload_to_firestore(sales_df, COLLECTION_NAME)
