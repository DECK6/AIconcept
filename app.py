import streamlit as st
import anthropic
import os
from dotenv import load_dotenv
import PyPDF2
import io
from pyhwp import hwp5odt
from pyhwp.hwp5 import xmlmodel
from lxml import etree
import olefile
import zlib

# .env 파일에서 환경 변수 로드
load_dotenv()

# Anthropic API 키 설정
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")

# Anthropic 클라이언트 초기화
client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

def read_pdf(file):
    pdf_reader = PyPDF2.PdfReader(file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text() + "\n"
    return text

def read_hwp(file):
    try:
        with hwp5odt.ODTTransform() as transform:
            os_odt_file = transform(file)
            with os_odt_file.open('content.xml') as f:
                xml_content = f.read()
                root = etree.fromstring(xml_content)
                paragraphs = root.findall('.//{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p')
                text = '\n'.join([p.text for p in paragraphs if p.text])
        return text
    except Exception as e:
        st.error(f"HWP 파일 읽기 오류: {str(e)}")
        return None

def read_hwpx(file):
    try:
        ole = olefile.OleFileIO(file)
        encoded_text = ole.openstream('BodyText/Section0').read()
        decoded_text = zlib.decompress(encoded_text, -15).decode('utf-16')
        return decoded_text
    except Exception as e:
        st.error(f"HWPX 파일 읽기 오류: {str(e)}")
        return None

def read_file(file):
    if file.type == "application/pdf":
        return read_pdf(file)
    elif file.type == "text/plain":
        return file.getvalue().decode("utf-8")
    elif file.type == "application/x-hwp":
        return read_hwp(file)
    elif file.type == "application/hwp":
        return read_hwpx(file)
    else:
        return None

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

def main():
    st.title("RFP 분석 및 전략 생성 도구")

    # 파일 업로드
    uploaded_file = st.file_uploader("RFP 파일을 업로드하세요", type=["txt", "pdf", "hwp", "hwpx"])

    if uploaded_file is not None:
        # 파일 내용 읽기
        content = read_file(uploaded_file)

        if content is None:
            st.error("파일을 읽을 수 없습니다. txt, pdf, hwp 또는 hwpx 파일만 지원합니다.")
            return

        # 파일 내용 미리보기
        st.subheader("파일 내용 미리보기")
        st.text_area("", content[:500] + "...", height=200)

        # 분석 시작 버튼
        if st.button("분석 시작"):
            with st.spinner("분석 중..."):
                # RFP 요약
                summary_prompt = f"""다음은 RFP의 내용입니다:

                {content}

                위 사업의 내용을 사업명, 발주처, 사업기간, 세부 사업기간, 장소, 내용, 사업목적, 추진방향을 요약하고, 메인 과업과 주가 되는 주요 과업을 요약해주세요. 선정 방식, 방법, 일반 사항 등 사업과 직접 관련이 없는 내용은 제외해주세요."""

                rfp_summary = generate_content(summary_prompt)
                if rfp_summary:
                    st.subheader("RFP 요약")
                    st.write(rfp_summary)

                # 커뮤니케이션 전략 생성
                comm_prompt = f"""다음은 RFP의 요약 내용입니다:

                {rfp_summary}

                위의 사업 내용을 바탕으로 커뮤니케이션 전략을 생성해주세요. 다음 구조화된 형식을 따라 작성해주세요:

                페이지 제목: 커뮤니케이션 전략

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
                    st.subheader("커뮤니케이션 전략")
                    st.write(comm_strategy)

                # 성공 전략 생성
                success_prompt = f"""다음은 RFP의 요약 내용입니다:

                {rfp_summary}

                위의 사업 내용을 바탕으로 성공 전략을 생성해주세요. 다음 구조화된 형식을 따라 작성해주세요:

                페이지 제목: 성공 전략

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
                    st.subheader("성공 전략")
                    st.write(success_strategy)

                # 성공 전략 상세 생성
                detail_prompt = f"""다음은 RFP의 요약 내용입니다:

                {rfp_summary}

                그리고 이는 앞서 생성한 성공 전략입니다:

                {success_strategy}

                위의 내용을 바탕으로 각 전략에 대한 상세 내용을 작성해주세요. 다음 구조화된 형식을 따라 작성해주세요:

                페이지 제목: 성공 전략 상세

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
                    st.subheader("성공 전략 상세")
                    st.write(strategy_details)

if __name__ == "__main__":
    main()
