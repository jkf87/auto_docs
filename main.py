import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def main(csv_file_path, hwp_file_path):
    # CSV 파일을 읽기
    data = pd.read_csv(csv_file_path)

    # 한글 객체 생성 및 파일 열기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(hwp_file_path)
    hwp.XHwpWindows.Item(0).Visible = True

    # 필드 목록 가져오기
    field_list = hwp.GetFieldList().split('\x02')
    field_names = [field.split('\x00')[0] for field in field_list]

    # 데이터를 한글 파일의 필드에 삽입하기
    for idx, row in data.iterrows():
        for field_name in field_names:
            if field_name in data.columns:
                hwp.PutFieldText(field_name, str(row[field_name]))

        # PDF로 저장하기
        output_pdf_name = f'{row["이름"]}_결과_{idx}.pdf'
        output_pdf_path = os.path.join(os.path.dirname(hwp_file_path), output_pdf_name)

        hwp.SaveAs(output_pdf_path, "PDF")

    # 한글 파일 닫기
    hwp.Quit()

def select_file(file_type):
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window
    messagebox.showinfo("Select File", f"데이터가 들어있는 {file_type} file을 선택해주세요.")
    file_path = filedialog.askopenfilename(title=f"한글 문서를 선택해주세요. 확장자 {file_type}", filetypes=[(f"{file_type} files", f"*.{file_type.lower()}")])
    return file_path

# GUI를 통해 파일 선택
csv_file = select_file("CSV")  # CSV 파일 선택
hwp_file = select_file("HWP")  # HWP 파일 선택

# 메인 함수 실행
main(csv_file, hwp_file)
