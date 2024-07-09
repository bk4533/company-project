import openai
import os
API_KEY= os.getenv("FLASK_API_KEY")
# OpenAI API 키 설정
openai.api_key = API_KEY  # 실제 API 키로 대체하세요.

# 파일 경로 설정
input_file_path = 'C:\\Users\\bk453\\Desktop\\새 폴더\\test.txt'
output_file_path = 'C:\\Users\\bk453\\Desktop\\새 폴더\\processed_test.txt'

# 첫 번째 파일 읽기
try:
    with open(input_file_path, 'r', encoding='utf-8') as file:
        text = file.read()
except FileNotFoundError:
    print(f"파일을 찾을 수 없습니다: {input_file_path}")
    exit(1)

# 텍스트를 분할하는 함수 (예: 단락 단위로 분할)
def split_text(text, max_tokens=4500):
    words = text.split()
    chunks = []
    current_chunk = []
    current_length = 0

    for word in words:
        word_length = len(word)
        if current_length + word_length > max_tokens:
            chunks.append(' '.join(current_chunk))
            current_chunk = [word]
            current_length = word_length
        else:
            current_chunk.append(word)
            current_length += word_length + 1  # +1 for space

    if current_chunk:
        chunks.append(' '.join(current_chunk))
    
    return chunks

# 텍스트를 분할
text_chunks = split_text(text)

# 분할된 텍스트를 처리하고 결과를 저장할 리스트
processed_chunks = []

# 각 청크를 처리
for i, chunk in enumerate(text_chunks):
    print(f"Processing chunk {i+1}/{len(text_chunks)}")
    response = openai.ChatCompletion.create(
        model="gpt-4o",
        messages=[
            {"role": "user", "content": f"다음 텍스트를 문단으로 나누고 줄바꿈을 추가해줘. 텍스트 내용을 변경하지않기, 내용 줄이지 않기, 소제목을 넣지 말기, 원본 내용을 유지해.: {chunk}"}
        ]
    )
    processed_chunk = response['choices'][0]['message']['content']
    processed_chunks.append(processed_chunk)

# 모든 처리된 청크를 결합
final_text = '\n\n'.join(processed_chunks)

# 결과를 같은 파일에 저장
with open(output_file_path, 'w', encoding='utf-8') as file:
    file.write(final_text)

print(f"파일이 성공적으로 저장되었습니다: {output_file_path}")
