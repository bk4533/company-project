import difflib
import tkinter as tk
from tkinter import messagebox

# 파일 경로 설정
input_file_path = 'C:\\Users\\미림미디어랩\\Desktop\\chatgpt1\\새 폴더\\test.txt'
output_file_path = 'C:\\Users\\미림미디어랩\\Desktop\\chatgpt1\\새 폴더\\re_test.txt'
output_diff_file_path = 'C:\\Users\\미림미디어랩\\Desktop\\chatgpt1\\새 폴더\\diff.txt'

# 파일 읽기 함수
def read_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except FileNotFoundError:
        print(f"파일을 찾을 수 없습니다: {file_path}")
        exit(1)

# 원본 텍스트와 결과 텍스트 읽기
original_text = read_file(input_file_path)
processed_text = read_file(output_file_path)

# 줄바꿈 문자 제거 함수
def remove_newlines(text):
    return text.replace('\n\n', '').replace('\r', '')

# 줄바꿈 문자를 제거한 텍스트
original_text_no_newlines = remove_newlines(original_text)
processed_text_no_newlines = remove_newlines(processed_text)

# 원본 텍스트와 결과 텍스트 비교 및 차이 출력
def compare_texts_and_show_diff(original, processed, output_diff_file):
    if original == processed:
        message = "원본 텍스트와 결과 텍스트가 동일합니다."
    else:
        diff = list(difflib.ndiff(original, processed))
        added = [line[2:] for line in diff if line.startswith('+ ')]
        removed = [line[2:] for line in diff if line.startswith('- ')]
        
        added_message = "추가된 내용: " + ' '.join(added) if added else "추가된 내용이 없습니다."
        removed_message = "삭제된 내용: " + ' '.join(removed) if removed else "삭제된 내용이 없습니다."
        
        message = f"원본 텍스트와 결과 텍스트가 다릅니다. 차이점은 다음과 같습니다:\n\n{added_message}\n\n{removed_message}"
        
        # 차이점을 파일로 저장
        with open(output_diff_file, 'w', encoding='utf-8') as file:
            file.write(message)
    
    # 메시지 박스 표시
    root = tk.Tk()
    root.withdraw()  # 숨김 상태의 Tkinter 루트 윈도우
    messagebox.showinfo("텍스트 비교 결과", message)
    root.destroy()

# 비교 수행 및 차이점 출력
compare_texts_and_show_diff(original_text_no_newlines, processed_text_no_newlines, output_diff_file_path)
