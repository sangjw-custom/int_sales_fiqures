import firebase_admin
from firebase_admin import credentials, firestore
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# 1. Firebase 접속 설정
def initialize_db():
    if not firebase_admin._apps:
        # 서비스 계정 키 파일명이 다르면 여기서 수정하세요.
        cred = credentials.Certificate("int-sales-figures_01.json")
        firebase_admin.initialize_app(cred)
    return firestore.client()

# 2. 파일 선택 창 띄우기
def select_file():
    root = tk.Tk()
    root.withdraw() # 메인 윈도우 숨기기
    
    # .xlsb, .xlsx 파일만 선택하도록 설정
    file_path = filedialog.askopenfilename(
        title="업로드할 엑셀 파일을 선택하세요",
        filetypes=[("Excel files", "*.xlsb *.xlsx *.xls")]
    )
    return file_path

# 3. 업로드 실행 함수
def start_upload():
    file_path = select_file()
    
    if not file_path:
        print("파일 선택이 취소되었습니다.")
        return

    try:
        db = initialize_db()
        
        # 컬렉션 이름을 파일명이나 날짜로 자동 생성 (예: Sales_Data)
        collection_name = "Sales_Profit_Upload" 
        
        print(f"[{os.path.basename(file_path)}] 읽는 중...")
        df = pd.read_excel(file_path, engine='pyxlsb')
        df = df.where(pd.notnull(df), None) # 빈값 처리

        batch = db.batch()
        total_count = len(df)
        
        for index, row in df.iterrows():
            doc_data = row.to_dict()
            doc_ref = db.collection(collection_name).document()
            batch.set(doc_ref, doc_data)
            
            if (index + 1) % 500 == 0:
                batch.commit()
                batch = db.batch()
                print(f"{index + 1} / {total_count}개 완료...")

        batch.commit()
        messagebox.showinfo("성공", f"{total_count}건의 데이터가 성공적으로 업로드되었습니다!")
        
    except Exception as e:
        messagebox.showerror("오류", f"업로드 중 문제가 발생했습니다:\n{e}")

if __name__ == "__main__":
    start_upload()
