import streamlit as st
import anthropic
import os
import PyPDF2
import io
import zlib
import struct
import re
import pandas as pd
import olefile
import warnings
import zipfile
import xml.etree.ElementTree as ET
import unicodedata

# 시스템 환경 변수에서 API 키 가져오기
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")

# API 키가 설정되어 있는지 확인
if not ANTHROPIC_API_KEY:
    st.error("ANTHROPIC_API_KEY 환경 변수가 설정되지 않았습니다.")
    st.stop()

# Anthropic 클라이언트 초기화
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

class HwpTextExtractor:
    def __init__(self):
        self.stopwords = []

    def get_hwp_text(self, filename, target_page):
        with olefile.OleFileIO(filename) as f:
            dirs = f.listdir()

            if ["FileHeader"] not in dirs or ["\x05HwpSummaryInformation"] not in dirs:
                raise Exception("Not Valid HWP.")

            header_data = f.openstream("FileHeader").read()
            is_compressed = (header_data[36] & 1) == 1

            sections = [d for d in dirs if d[0] == "BodyText"]
            sections.sort(key=lambda x: int(x[1][len("Section"):]))
            
            text = ""
            for section in sections:
                section_data = f.openstream(section).read()
                if is_compressed:
                    unpacked_data = zlib.decompress(section_data, -15)
                else:
                    unpacked_data = section_data

                i = 0
                size = len(unpacked_data)
                page_count = 0
                while i < size:
                    header = struct.unpack_from("<I", unpacked_data, i)[0]
                    rec_type = header & 0x3ff
                    rec_len = (header >> 20) & 0xfff

                    if rec_type == 67:
                        rec_data = unpacked_data[i+4:i+4+rec_len]
                        try:
                            text += rec_data.decode('utf-16')
                        except UnicodeDecodeError:
                            print(f"Failed to decode. Skipping...")
                        text += "\n"

                        if b'\x14' in rec_data:
                            page_count += 1
                            if page_count >= target_page:
                                break

                    i += 4 + rec_len

                text += "\n"

            return text

    def extract_text_from_hwp(self, hwp_file, target_page):
        extracted_text = self.get_hwp_text(hwp_file, target_page)
        processed_text = self.process_text(extracted_text)
        processed_text = self.remove_special_chars(processed_text)
        return processed_text

    def remove_chinese_characters(self, s: str):
        return re.sub(r'[\u4e00-\u9fff]+', '', s)
        
    def remove_control_characters(self, s):
        return "".join(ch for ch in s if unicodedata.category(ch)[0]!="C")

    def process_text(self, text):
        # 기존 처리
        processed_text = text.replace('\x02', ' ')
        processed_text = processed_text.replace('\x03', ' ')
        processed_text = processed_text.replace('\x0b', ' ')
        processed_text = processed_text.replace('\x0c', ' ')
        processed_text = re.sub(r'\s+', ' ', processed_text)
        
        # 새로운 정제 과정 추가
        processed_text = self.remove_chinese_characters(processed_text)
        processed_text = self.remove_control_characters(processed_text)        
        return processed_text.strip()

    def remove_special_chars(self, text):
        special_chars = r'[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\\\|\\(\\)\\[\\]\\<\\>`\'…》]'
        text = re.sub(special_chars, '', text)
        return text
    
class HwpxTextExtractor:
    def __init__(self):
        self.stopwords = []

    def convert_hwpx_to_txt(self, hwpx_file_path):
        base, _ = os.path.splitext(hwpx_file_path)
        zip_file_path = base + ".zip"
        os.rename(hwpx_file_path, zip_file_path)
        all_texts = []
        
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            all_files = zip_ref.namelist()
            section_files = [f for f in all_files if f.startswith('Contents/section') and f.endswith('.xml')]
            
            for section_file in section_files:
                zip_ref.extract(section_file, '.')
                section_text = self.extract_text_from_xml(section_file)
                all_texts.extend(section_text)
                os.remove(section_file)
        
        os.remove(zip_file_path)
        combined_text = '\n'.join(all_texts)
        return combined_text

    def extract_text_from_xml(self, xml_file_path):
        texts = []
        try:
            tree = ET.parse(xml_file_path)
            root = tree.getroot()
            for elem in root.iter('{http://www.hancom.co.kr/hwpml/2011/paragraph}t'):
                if elem.text:
                    texts.append(elem.text.strip())
        except ET.ParseError as e:
            return [f"XML 파싱 에러: {e}"]
        return texts

    def extract_text_from_hwpx(self, hwpx_file_path):
        extracted_text = self.convert_hwpx_to_txt(hwpx_file_path)
        processed_text = self.remove_special_chars(extracted_text)
        return processed_text

    def remove_special_chars(self, text):
        special_chars = r'[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\\\|\\(\\)\\[\\]\\<\\>`\'…》]'
        text = re.sub(special_chars, '', text)
        return text

def remove_chinese_characters(s: str):
    return re.sub(r'[\u4e00-\u9fff]+', '', s)

def remove_control_characters(s):
    return "".join(ch for ch in s if unicodedata.category(ch)[0]!="C")

def remove_special_chars(text):
    special_chars = r'[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\\\|\\(\\)\\[\\]\\<\\>`\'…》]'
    return re.sub(special_chars, '', text)

def refine_text(text):
    text = remove_chinese_characters(text)
    text = remove_control_characters(text)
    text = remove_special_chars(text)
    return text.strip()

def read_txt(file):
    try:
        content = file.getvalue().decode("utf-8")
        return refine_text(content)
    except Exception as e:
        st.error(f"TXT 파일 읽기 오류: {str(e)}")
        return None

def read_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
        return refine_text(text)
    except Exception as e:
        st.error(f"PDF 파일 읽기 오류: {str(e)}")
        return None

def read_hwp(file):
    try:
        hwp_extractor = HwpTextExtractor()
        content = hwp_extractor.extract_text_from_hwp(file, 15)  # 15 페이지까지 읽기
        return refine_text(content)
    except Exception as e:
        st.error(f"HWP 파일 읽기 오류: {str(e)}")
        return None

def read_hwpx(file):
    try:
        hwpx_extractor = HwpxTextExtractor()
        content = hwpx_extractor.extract_text_from_hwpx(file)
        return refine_text(content)
    except Exception as e:
        st.error(f"HWPX 파일 읽기 오류: {str(e)}")
        return None

def read_file(file):
    if file is None:
        st.error("파일이 업로드되지 않았습니다.")
        return None
    
    file_extension = os.path.splitext(file.name)[1].lower()
    
    content = None
    if file_extension == '.txt':
        content = read_txt(file)
    elif file_extension == '.pdf':
        content = read_pdf(file)
    elif file_extension == '.hwp':
        content = read_hwp(file)
    elif file_extension == '.hwpx':
        content = read_hwpx(file)
    else:
        st.error(f"지원되지 않는 파일 형식입니다: {file_extension}")
        return None
    
    if content:
        st.success(f"파일 읽기 성공: {len(content)} 문자")
    else:
        st.error("파일 내용을 읽을 수 없습니다.")
    
    return content

def generate_content(prompt):
    try:
        response = client.messages.create(
            model="claude-3-5-sonnet-20240620",
            max_tokens=4000,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        return response.content[0].text
    except Exception as e:
        st.error(f"API 호출 오류: {str(e)}")
        return None

def display_strategy_slide(title, content):
    slide_html = """
    <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 100vh; background-color: #f3f4f6; padding: 2rem;">
        <div style="width: 100%; max-width: 48rem; background-color: white; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05); border-radius: 0.5rem; overflow: hidden;">
            <div style="padding: 1.5rem; border-bottom: 1px solid #e5e7eb;">
                <h1 style="font-size: 1.5rem; font-weight: bold; text-align: center; color: #2563eb;">
                    {title}
                </h1>
            </div>
            <div style="padding: 1.5rem;">
                <div style="font-size: 1rem; text-align: left;">
                    {content}
                </div>
            </div>
        </div>
    </div>
    """
    formatted_content = content.replace('\n', '<br>')
    formatted_slide = slide_html.format(title=title, content=formatted_content)
    st.markdown(formatted_slide, unsafe_allow_html=True)

def main():
    st.title("RFP 분석 및 전략 생성 도구")

    uploaded_file = st.file_uploader("RFP 파일을 업로드하세요", type=["txt", "pdf", "hwp", "hwpx"])

    if uploaded_file is not None:
        content = read_file(uploaded_file)
        
        if content is None:
            st.error("파일 내용을 읽을 수 없습니다. 파일 형식을 확인해주세요.")
            return

        st.subheader("파일 내용 미리보기")
        st.text_area("", content[:500] + "...", height=200)

        if st.button("분석 시작"):
            with st.spinner("분석 중..."):
                # RFP 요약 (생성은 하되 표시하지 않음)
                summary_prompt = f"""다음은 RFP의 내용입니다:

                {content[:4000]}  # API 제한으로 인해 내용을 잘랐습니다.

                위 사업의 내용을 사업명, 발주처, 사업기간, 세부 사업기간, 장소, 내용, 사업목적, 추진방향을 요약하고, 메인 과업과 주가 되는 주요 과업을 요약해주세요. 선정 방식, 방법, 일반 사항 등 사업과 직접 관련이 없는 내용은 제외해주세요."""

                rfp_summary = generate_content(summary_prompt)

                # 커뮤니케이션 전략 생성
                comm_prompt = f"""다음은 RFP의 요약 내용입니다:

                {rfp_summary}

                위의 사업 내용을 바탕으로 커뮤니케이션 전략을 생성해주세요. 다음 구조화된 형식을 따라 작성해주세요:

                헤드라인 메시지: [전체 전략을 대표하는 핵심 메시지]

                거버닝 메시지: [전략의 핵심을 요약하는 한 문장]

                본문:
                • 메시지 1 (공식적이고 포멀한 메시지):
                  - [메시지 내용]
                  - [배경 설명 및 실행 계획]

                • 메시지 2 (언어 유희가 포함된 키치하고 재미있는, 젊은 세대에 어필하는 메시지):
                  - [메시지 내용]
                  - [배경 설명 및 실행 계획]

                • 메시지 3 (메시지 1과 2의 중간 정도 느낌의 메시지):
                  - [메시지 내용]
                  - [배경 설명 및 실행 계획]

                각 메시지에 대한 실행 계획은 구체적이고 상세해야 합니다. 기획, 분석 등의 일반적인 내용이 아닌, 실제 수행 방법과 단계를 자세히 기술해주세요."""

                comm_strategy = generate_content(comm_prompt)
                if comm_strategy:
                    display_strategy_slide("커뮤니케이션 전략", comm_strategy)

                # 성공 전략 생성
                success_prompt = f"""다음은 RFP의 요약 내용입니다:

                {rfp_summary}

                위의 사업 내용을 바탕으로 성공 전략을 생성해주세요. 다음 구조화된 형식을 따라 작성해주세요:

                헤드라인 메시지: [전체 전략을 대표하는 핵심 메시지]

                거버닝 메시지: [전략의 핵심을 요약하는 한 문장]

                본문:
                • 핵심 전략 키워드: [5글자 영어 단어]
                  - [각 글자에 해당하는 전략 단어 설명]
                  - [각 전략 단어를 포함한 헤드라인 메시지]
                  - [각 전략에 대한 상세 실행 계획]

                실행 계획은 구체적이고 상세해야 합니다. 기획, 분석 등의 일반적인 내용이 아닌, 실제 수행 방법과 단계를 자세히 기술해주세요."""

                success_strategy = generate_content(success_prompt)
                if success_strategy:
                    display_strategy_slide("성공 전략", success_strategy)

                # 성공 전략 상세 생성
                detail_prompt = f"""다음은 RFP의 요약 내용입니다:

                {rfp_summary}

                그리고 이는 앞서 생성한 성공 전략입니다:

                {success_strategy}

                위의 내용을 바탕으로 각 전략에 대한 상세 내용을 작성해주세요. 다음 구조화된 형식을 따라 작성해주세요:

                헤드라인 메시지: [전체 전략을 대표하는 핵심 메시지]

                거버닝 메시지: [전략의 핵심을 요약하는 한 문장]

                본문:
                • [첫 번째 전략 단어]:
                  - [헤드라인 메시지]
                  - [중요성 설명]
                  - [구체적인 실행 방안]
                  - [기대 효과]

                • [두 번째 전략 단어]:
                  - [헤드라인 메시지]
                  - [중요성 설명]
                  - [구체적인 실행 방안]
                  - [기대 효과]

                [나머지 전략 단어들에 대해서도 같은 형식으로 작성]

                각 전략에 대한 설명과 실행 방안은 매우 구체적이고 상세해야 합니다. 실제 수행 방법, 단계, 필요한 자원 등을 자세히 기술해주세요."""

                strategy_details = generate_content(detail_prompt)
                if strategy_details:
                    display_strategy_slide("성공 전략 상세", strategy_details)

if __name__ == "__main__":
    main()
