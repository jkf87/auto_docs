import os
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, Label, Button, StringVar, Entry, Radiobutton, Checkbutton
import win32clipboard
from io import BytesIO
from PIL import Image

def get_image_files(folder_path):
    # 지정된 폴더에서 이미지 파일 목록을 가져옵니다.
    supported_formats = [".jpg", ".jpeg", ".png", ".bmp", ".gif"]
    files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if os.path.splitext(f)[1].lower() in supported_formats]
    return sorted(files)

def main(csv_file_path, hwp_file_path, image_folder, save_format):
    # CSV 파일을 읽기
    data = pd.read_csv(csv_file_path)

    # 이미지 폴더가 제공되었는지 확인하고 이미지 파일 목록을 가져옵니다.
    image_files = get_image_files(image_folder) if image_folder else []

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

        # 이미지 처리
        if image_files and idx < len(image_files):
            image_path = image_files[idx]
            if os.path.exists(image_path):
                # 이전에 삽입된 이미지 삭제
                hwp.MoveToField("이미지")
                hwp.Run("SelectAll")
                hwp.Run("Delete")

                # 이미지를 클립보드에 복사
                image = Image.open(image_path)
                output = BytesIO()
                image.convert("RGB").save(output, "BMP")
                clipboard_data = output.getvalue()[14:]
                output.close()
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, clipboard_data)
                win32clipboard.CloseClipboard()

                # 한글 문서에서 이미지를 붙여넣기
                hwp.MoveToField("이미지")
                hwp.HAction.Run("Paste")

        # 파일 저장
        output_file_name = f'{row.iloc[0]}_결과_{idx}.{save_format.lower()}'
        output_file_path = os.path.join(os.path.dirname(hwp_file_path), output_file_name)
        if save_format == "HWP":
            hwp.SaveAs(output_file_path, "HWP")
        elif save_format == "PDF":
            hwp.SaveAs(output_file_path, "PDF")

        # 저장된 파일에 대한 로그 출력
        print(f'File saved: {output_file_path}')

    # 한글 파일 닫기
    hwp.Quit()



def select_file(entry, file_type):
    file_path = filedialog.askopenfilename(filetypes=[(f"{file_type} files", f"*.{file_type.lower()}")])
    entry.delete(0, tk.END)  # Clear the entry
    entry.insert(0, file_path)  # Insert the selected file path

def select_folder(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)  # Clear the entry
    entry.insert(0, folder_path)  # Insert the selected folder path

def start_main_process(csv_entry, hwp_entry, folder_entry, save_format_var, insert_images): # 시작버튼 클릭시 실행
    csv_file = csv_entry.get() # csv파일 경로 가져오기
    hwp_file = hwp_entry.get() # hwp파일 경로 가져오기
    image_folder = None if not insert_images.get() else folder_entry.get() # 이미지 폴더 경로 가져오기
    save_format = save_format_var.get() # 저장 형식 가져오기 (hwp, pdf)

    # Validate if the fields are not empty
    if csv_file and hwp_file :
        # Here you can call your main function with the paths
        main(csv_file, hwp_file, image_folder, save_format)
        print("Starting main process...")
    else:
        print("Please select all the files and folder before proceeding.")

# GUI Setup
root = tk.Tk()
root.title("한글 문서 텍스트, 이미지 삽입 및 저장 프로그램")
insert_images_var = tk.BooleanVar()

# Define the StringVar
csv_var = StringVar()
hwp_var = StringVar()
folder_var = StringVar()
save_format_var = StringVar(value="pdf") # 기본값 pdf

# Define the Entry widgets
csv_entry = Entry(root, textvariable=csv_var, width=50)
hwp_entry = Entry(root, textvariable=hwp_var, width=50)
folder_entry = Entry(root, textvariable=folder_var, width=50)

csv_entry.grid(row=0, column=1, padx=2, pady=2)
hwp_entry.grid(row=1, column=1, padx=2, pady=2)
folder_entry.grid(row=2, column=1, padx=2, pady=2)

# Define the Labels
Label(root, text="CSV파일을 선택하세요").grid(row=0, column=0, sticky='w', padx=10, pady=10)
Label(root, text="HWP파일을 선택하세요").grid(row=1, column=0, sticky='w', padx=10, pady=10)
Label(root, text="이미지가 저장된 폴더를 선택하세요").grid(row=2, column=0, sticky='w', padx=10, pady=10)

# Define the Entry widgets
csv_entry.grid(row=0, column=1, padx=10, pady=10)
hwp_entry.grid(row=1, column=1, padx=10, pady=10)
folder_entry.grid(row=2, column=1, padx=10, pady=10)

# Define the Checkbutton for image attachment
Checkbutton(root, text="이미지첨부하기", variable=insert_images_var).grid(row=3, column=0, columnspan=2, sticky='w', padx=10, pady=10)

# Define Radio Buttons for file save format
Radiobutton(root, text="HWP", variable=save_format_var, value="HWP").grid(row=4, column=0, sticky='w', padx=10, pady=10)
Radiobutton(root, text="PDF", variable=save_format_var, value="PDF").grid(row=4, column=1, sticky='w', padx=10, pady=10)

# Define the Buttons
Button(root, text="Browse", command=lambda: select_file(csv_entry, "CSV")).grid(row=0, column=2, padx=10, pady=10)
Button(root, text="Browse", command=lambda: select_file(hwp_entry, "HWP")).grid(row=1, column=2, padx=10, pady=10)
Button(root, text="Browse", command=lambda: select_folder(folder_entry)).grid(row=2, column=2, padx=10, pady=10)
Button(root, text="Start", command=lambda: start_main_process(csv_entry, hwp_entry, folder_entry, save_format_var, insert_images_var)).grid(row=5, column=1, pady=10)
root.mainloop()
