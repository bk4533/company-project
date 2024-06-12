import tkinter as tk
from tkinter import filedialog, messagebox
import re
import os

def select_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        input_entry_folder.delete(0, tk.END)
        input_entry_folder.insert(0, folder_selected)

def select_file():
    file_selected = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file_selected:
        input_entry_file.delete(0, tk.END)
        input_entry_file.insert(0, file_selected)

def process_files_in_folder():
    folder_path = input_entry_folder.get()
    if not folder_path:
        messagebox.showwarning("경고", "폴더 경로를 입력하세요!")
        return
    
    # 폴더 내의 모든 파일 처리
    for filename in os.listdir(folder_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()
            processed_text = further_process_text(text)
            output_file_path = os.path.join(folder_path, f"re_{filename}")
            with open(output_file_path, 'w', encoding='utf-8') as file:
                file.write(processed_text)
    
    messagebox.showinfo("완료", "폴더 내 모든 파일 처리가 완료되었습니다!")

def process_selected_file():
    file_path = input_entry_file.get()
    if not file_path:
        messagebox.showwarning("경고", "파일 경로를 입력하세요!")
        return

    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
    processed_text = further_process_text(text)
    output_file_path = os.path.join(os.path.dirname(file_path), f"re_{os.path.basename(file_path)}")
    with open(output_file_path, 'w', encoding='utf-8') as file:
        file.write(processed_text)

    messagebox.showinfo("완료", "파일 처리가 완료되었습니다!")

def further_process_text(text):
    result = []

    # 텍스트를 마침표와 물음표로 나누어 문장 리스트로 만듦
    sentences = re.split(r'(?<=\.)|(?<=\?)', text)

    # 각 문장을 처리
    for sentence in sentences:
        # 문장의 앞뒤 공백 제거
        sentence = sentence.strip()

        # 빈 문장은 건너뜀
        if not sentence:
            continue

        # 문장의 길이가 70자 미만인 경우
        if len(sentence) < 70:
            # 70자에 가까운 마지막 마침표나 물음표 위치 찾기
            punctuation_index = max(sentence.rfind('.', 0, 70), sentence.rfind('?', 0, 70))
            if punctuation_index != -1:
                sentence = sentence[:punctuation_index + 1] + '\n\n' + sentence[punctuation_index + 1:]
            result.append(sentence)


        # 문장의 길이가 70자 이상 90자 미만인 경우
        elif 70 <= len(sentence) < 90:
            sentence = sentence.replace('. ', '.\n\n').replace('? ', '?\n\n')
            result.append(sentence)

        # 남은 문장은 앞에 '**\n'을 추가하여 유지
        else:
            result.append('**\n' + sentence)

    # 결과 리스트를 두 번 줄바꿈으로 연결하여 반환
    return '\n\n'.join(result)

def center_window(window, width=600, height=150):
    # 창의 너비와 높이 설정
    window.geometry(f'{width}x{height}')
    # 화면의 너비와 높이 가져오기
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    # 창의 위치 계산
    x = int((screen_width / 2) - (width / 2))
    y = int((screen_height / 2) - (height / 2))
    # 창의 위치 설정
    window.geometry(f'+{x}+{y}')

# GUI 설정
root = tk.Tk()
root.title("문단 정리")

center_window(root)

frame = tk.Frame(root)
frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

font_large = ("Arial", 10)  # 기본 폰트 크기

# 폴더 선택
input_label_folder = tk.Label(frame, text="폴더 경로:", font=font_large)
input_label_folder.grid(row=0, column=0, padx=5, pady=(20, 10))

input_entry_folder = tk.Entry(frame, width=50, font=font_large)
input_entry_folder.grid(row=0, column=1, padx=5, pady=(20, 10))

select_button_folder = tk.Button(frame, text="폴더 선택", command=select_folder, font=font_large)
select_button_folder.grid(row=0, column=2, padx=5, pady=(20, 10))

process_button_folder = tk.Button(frame, width=3, text="▶", command=process_files_in_folder, font=font_large, anchor='center')
process_button_folder.grid(row=0, column=4, padx=15, pady=(20, 10))  # 오른쪽으로 이동

# 파일 선택
input_label_file = tk.Label(frame, text="파일 경로:", font=font_large)
input_label_file.grid(row=1, column=0, padx=5, pady=(20, 10))

input_entry_file = tk.Entry(frame, width=50, font=font_large)
input_entry_file.grid(row=1, column=1, padx=5, pady=(20, 10))

select_button_file = tk.Button(frame, text="파일 선택", command=select_file, font=font_large)
select_button_file.grid(row=1, column=2, padx=5, pady=(20, 10))

process_button_file = tk.Button(frame, width=3, text="▶", command=process_selected_file, font=font_large, anchor='center')
process_button_file.grid(row=1, column=4, padx=15, pady=(20, 10))  # 오른쪽으로 이동

root.mainloop()
