import win32com.client as win32
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
import os
import time

# --- [사용자 설정] ---
KEY_PATH = "int-sales-figures_01.json" 
# 실제 엑셀 파일 경로 (예: r"C:\Users\User\Documents\손익데이터.xlsb")
EXCEL_FILE_PATH = r"D:\72. AI TEST\Firebase\sales_fiqures\2603 sales_test.xlsb"
COLLECTION_NAME = "int-sales-figures"
# --------------------

def initialize_firebase():
    if not firebase_admin._apps:
        cred = credentials.Certificate(KEY_PATH)
        firebase_admin.initialize_app(cred)
    return firestore.client()

def upload_data():
    try:
        db = initialize_firebase()
        
        # 1. 엑셀 앱 제어 (내 PC의 엑셀 권한으로 보안 문서 열기)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        
        print(f"📌 엑셀 파일을 읽는 중입니다... (약간의 시간이 소요될 수 있습니다)")
        wb = excel.Workbooks.Open(os.path.abspath(EXCEL_FILE_PATH))
        ws = wb.ActiveSheet
        
        # 데이터 영역 전체 로드
        raw_data = ws.UsedRange.Value
        
        # 2. 데이터프레임 변환
        df = pd.DataFrame(raw_data[1:], columns=raw_data[0])
        wb.Close(False)
        excel.Quit()
        
        # 3. Firestore 업로드 (13만 건 대용량 처리)
        df = df.where(pd.notnull(df), None) # 빈값 처리
        total = len(df)
        print(f"🚀 총 {total:,}건 업로드 시작 (500개씩 묶어서 전송)")
        
        batch = db.batch()
        start_time = time.time()

        for index, row in df.iterrows():
            doc_ref = db.collection(COLLECTION_NAME).document()
            batch.set(doc_ref, row.to_dict())
            
            # 500개마다 전송 (Firestore 제한)
            if (index + 1) % 500 == 0:
                batch.commit()
                batch = db.batch()
                elapsed = time.time() - start_time
                print(f"📦 {index + 1:,} / {total:,} 완료 (진행시간: {elapsed:.1f}초)")
        
        batch.commit() # 남은 데이터 전송
        print(f"✨ 성공! 모든 데이터가 '{COLLECTION_NAME}'에 업로드되었습니다.")
        
    except Exception as e:
        print(f"❌ 에러 발생: {e}")
        # 오류 시 엑셀 프로세스 강제 종료 방지
        if 'excel' in locals(): excel.Quit()

if __name__ == "__main__":
    upload_data()
