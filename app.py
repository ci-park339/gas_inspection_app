import base64
import json
import os
import re

import gspread
import pandas as pd
import fitz
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from PIL import Image

load_dotenv()

PROMPT = """이 이미지는 가스 기밀시험 자동 기록지(레코스타)입니다. 이미지 안의 텍스트를 분석하여 아래 JSON 배열 구조로만 응답하세요. 코드블록이나 설명 없이 JSON 배열만 출력하세요.

[
{
"검사넘버": "NO: 뒤의 숫자",
"기밀_유지_시간": "첫 측정 시각과 마지막 측정 시각의 차이(예: 4분)",
"최종_합격_여부": "판정대기",
"측정기록": [
{
"번호": "1~8 (용지 맨 왼쪽 물리 인덱스 0~7에 1을 더한 값)",
"검사일자": "YYYY-MM-DD",
"검사시각": "HH:MM:SS",
"압력": "단위 포함(예: 2.2kPa)",
"온도": "단위 포함(예: 22.8c)"
}
]
}
]

규칙 1 (다중 용지): 사진에 기밀시험 용지가 여러 장 붙어 있다면, 각각 별도의 객체로 분리하여 배열에 담으세요. 용지가 한 장이어도 배열로 응답하세요.
규칙 2 (누락 금지): 용지 1장당 반드시 8개의 측정 기록을 빠짐없이 측정기록 배열에 담으세요.
규칙 3 (판정 금지): 합격/불합격을 판정하지 마세요. 최종_합격_여부는 항상 "판정대기"로만 출력하고, 압력·온도·시각 등 숫자와 시간을 정확히 읽는 데만 집중하세요.
규칙 4 (시간 단위 내림 처리): 기밀_유지_시간을 계산할 때 '4분 1초'나 '4분 2초' 등 1~2초 단위의 자투리 시간이 포함되어 있다면, 초 단위를 버림(내림) 처리하여 "4분"과 같이 깔끔한 분 단위로만 기록하세요.
규칙 5 (데이터 유추 금지): 글자가 흐려 확신할 수 없다면 절대 지어내지 말고 해당 값을 무조건 "인식불가"로 출력하세요.
규칙 6 (기밀 유지 시간 보전): 중간이나 마지막 측정 시각이 "인식불가"로 판독되더라도, 기록 간격(보통 30초 내외)의 흐름을 유추하여 전체 검사 진행 시간을 짧게 깎아내리지 마세요. 전체 흐름상 시작과 끝의 시간 차이가 3분 50초 ~ 4분 59초 사이로 유추된다면 무조건 "4분"으로 통일해서 기록하세요.
규칙 7 (오독 주의 2/7, 0/6): 열전사 영수증의 특성상 숫자 2와 7, 0과 6이 매우 비슷하게 보입니다. 앞뒤 온도/압력의 흐름과 소수점 위치를 파악하여 절대 혼동하지 마세요.
규칙 8 (순서 밀림 절대 금지, 번호 1~8 통일): 용지 맨 왼쪽에 찍힌 물리적 인덱스 번호(0~7)를 기준으로 정확히 1:1 매핑하세요. 단, JSON 배열에 담을 때는 각 기록의 "번호"를 1번부터 8번까지(1~8)로 통일하고, 배열도 번호 순서(1→8)대로 담으세요. 만약 물리적 번호 3번(4번째 줄)이 안 보인다면 그 자리(번호 4)에 "인식불가"를 넣어야 하며, 절대 밑의 데이터를 위로 당겨서 배열을 압축하지 마세요."""


def get_tab(spreadsheet, title, header):
    """탭이 있으면 가져오고, 없으면 헤더와 함께 생성한다."""
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        tab = spreadsheet.add_worksheet(title=title, rows=1000, cols=len(header))
        tab.append_row(header)
        return tab


def is_blank(value):
    return str(value).strip().lower() in ("", "인식불가", "<na>", "null", "none", "nan")


def to_float(text):
    """'5.7kPa' 같은 문자열에서 숫자 부분만 float로 변환 (없으면 None)."""
    match = re.search(r"-?\d+(?:\.\d+)?", str(text))
    return float(match.group()) if match else None


NUMBER = r"-?\d+(?:\.\d+)?"


def num_text(value):
    """'5.7kPa' -> '5.7' (숫자 부분만, 없으면 '')."""
    match = re.search(NUMBER, str(value))
    return match.group() if match else ""


def join_row(n, time_in, pressure, temp):
    """폼의 입력 3개를 원래 포맷으로 결합. 잘못된 입력은 ValueError."""
    time_in, pressure, temp = (v.strip() for v in (time_in, pressure, temp))
    if time_in:
        match = re.fullmatch(r"(\d{1,2}):(\d{1,2}):(\d{1,2})", time_in)
        if not match or int(match[1]) > 23 or int(match[2]) > 59 or int(match[3]) > 59:
            raise ValueError(f"{n}번: 검사시각은 HH:MM:SS 형식(예: 16:09:37)으로 입력하세요.")
        time_text = ":".join(f"{int(g):02d}" for g in match.groups())
    else:
        time_text = ""
    for label, v in (("압력", pressure), ("온도", temp)):
        if v and not re.fullmatch(NUMBER, v):
            raise ValueError(f"{n}번: {label} 칸에는 숫자만 입력하세요.")
    return {
        "검사시각": time_text,
        "압력": f"{pressure}kPa" if pressure else "",
        "온도": f"{temp}c" if temp else "",
    }


def judge(paper):
    """첫 번째 압력을 기준선으로, 하나라도 그보다 작으면 '불합격', 아니면 '합격'.
    숫자를 읽을 수 없어 확정할 수 없으면 '판정불가'."""
    pressures = [to_float(r.get("압력")) for r in paper.get("측정기록") or []]
    if not pressures or pressures[0] is None:
        return "판정불가"
    baseline = pressures[0]
    if any(p is not None and p < baseline for p in pressures):
        return "불합격"
    return "판정불가" if None in pressures else "합격"


def is_valid(paper):
    """측정기록이 정확히 8개이고 검사시각/압력/온도 모두에 인식불가·빈값이 없으면 True."""
    records = paper.get("측정기록") or []
    return len(records) == 8 and not any(
        is_blank(r.get(k)) for r in records for k in ("검사시각", "압력", "온도")
    )


def save_paper(paper):
    """용지 1장을 마스터/디테일 탭에 적재한다."""
    sheet_url = os.getenv("GOOGLE_SHEET_URL")
    if not sheet_url:
        raise ValueError(".env 파일에 GOOGLE_SHEET_URL이 없습니다.")
    spreadsheet = gspread.service_account(filename="google_secret.json").open_by_url(sheet_url)
    records = paper.get("측정기록") or []
    if not records:
        raise ValueError(f"검사넘버 {paper.get('검사넘버')}: 측정기록이 비어 있습니다.")
    exam_no = paper.get("검사넘버")
    first_date = records[0].get("검사일자")
    month = str(first_date)[:7]  # YYYY-MM
    if not re.fullmatch(r"\d{4}-\d{2}", month):
        raise ValueError(f"검사일자를 확인할 수 없습니다: {first_date}")

    master_tab = get_tab(
        spreadsheet, f"마스터_{month}",
        ["검사넘버(PK)", "검사일자", "기밀_유지_시간", "최종_합격_여부"],
    )
    detail_tab = get_tab(
        spreadsheet, f"디테일_{month}",
        ["검사넘버(FK)", "검사시각", "압력", "온도"],
    )
    master_row = [exam_no, first_date, paper.get("기밀_유지_시간"), paper.get("최종_합격_여부")]
    detail_rows = [[exam_no, r.get("검사시각"), r.get("압력"), r.get("온도")] for r in records]
    master_tab.append_row(master_row, table_range="A1")
    detail_tab.append_rows(detail_rows, table_range="A1")

st.set_page_config(
    page_title="가스 기밀시험 용지 업로드",
    page_icon="📋",
    layout="centered",
)

st.title("📋 가스 기밀시험 용지 업로드")
st.markdown(
    """
    기밀시험 용지 사진을 올려 주세요. (PC·모바일 공통)
    1. **파일 업로드** 버튼으로 이미지를 선택합니다.
    2. 미리보기를 확인한 뒤 **데이터 추출하기**를 누릅니다.
    """
)

uploaded_file = st.file_uploader(
    "기밀시험 용지 이미지를 선택하세요",
    type=["jpg", "jpeg", "png", "pdf"],
)

st.markdown("---")
st.subheader("미리보기")

if uploaded_file is None:
    st.info("이미지를 업로드하면 여기에 미리보기가 표시됩니다.")
else:
    if uploaded_file.name.lower().endswith(".pdf"):
        # PDF는 비전 모델이 직접 읽지 못하므로 페이지별 PNG로 변환
        with fitz.open(stream=uploaded_file.getvalue(), filetype="pdf") as pdf:
            images = [page.get_pixmap(dpi=150).tobytes("png") for page in pdf]
        mime = "image/png"
        for i, img in enumerate(images, 1):
            st.image(img, caption=f"{uploaded_file.name} - {i}쪽", use_column_width=True)
    else:
        images = [uploaded_file.getvalue()]
        mime = uploaded_file.type
        st.image(Image.open(uploaded_file), caption=uploaded_file.name, use_column_width=True)

    if st.button("데이터 추출하기"):
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            st.error(".env 파일에 OPENAI_API_KEY가 없습니다.")
        else:
            image_parts = [
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:{mime};base64,{base64.b64encode(img).decode('utf-8')}",
                        "detail": "high",
                    },
                }
                for img in images
            ]
            with st.spinner("데이터를 추출하는 중..."):
                try:
                    response = OpenAI(api_key=api_key).chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {
                                "role": "user",
                                "content": [{"type": "text", "text": PROMPT}] + image_parts,
                            }
                        ],
                    )
                    content = response.choices[0].message.content
                    # 마크다운 코드블록 등을 걷어내고 JSON 배열 부분만 파싱
                    papers = json.loads(content[content.index("["):content.rindex("]") + 1])
                    for paper in papers:  # AI 판정 대신 파이썬 계산값으로 덮어쓰기
                        paper["최종_합격_여부"] = judge(paper)
                    # 화면 새로고침에 대비해 session_state에 보관 (저장 결과 saved/errors 포함)
                    st.session_state.update(
                        papers=papers, papers_file=uploaded_file.name, saved={}, errors={},
                        run=st.session_state.get("run", 0) + 1,
                    )
                    st.success("추출 완료")
                except Exception as e:
                    st.session_state.pop("papers", None)
                    st.error(f"추출 실패: {e}")

            # 갈래 A: 검증 통과 용지는 즉시 자동 저장 (갈래 B는 저장 차단)
            if "papers" in st.session_state:
                with st.spinner("구글 시트에 데이터를 저장하는 중..."):
                    for idx, paper in enumerate(st.session_state["papers"]):
                        if is_valid(paper):
                            try:
                                save_paper(paper)
                                st.session_state["saved"][idx] = "✅ 자동 저장 완료"
                            except Exception as e:
                                st.session_state["errors"][idx] = f"구글 시트 저장 실패: {e}"

    # 추출 결과 표시 (버튼을 누르지 않은 rerun에서도 유지)
    papers = st.session_state.get("papers")
    if papers and st.session_state.get("papers_file") == uploaded_file.name:
        saved, errors, run = st.session_state["saved"], st.session_state["errors"], st.session_state["run"]
        for idx, paper in enumerate(papers):
            records = paper.get("측정기록") or []
            exam_no = paper.get("검사넘버")
            st.subheader(f"📝 기밀시험 기록: {exam_no}")
            col1, col2, col3 = st.columns(3)
            col1.metric("검사일자", (records[0].get("검사일자") if records else None) or "-")
            col2.metric("기밀_유지_시간", paper.get("기밀_유지_시간") or "-")
            col3.metric("최종_합격_여부", paper.get("최종_합격_여부") or "-")

            if idx in saved:
                st.success(saved[idx])
                df = pd.DataFrame(records, columns=["검사시각", "압력", "온도"])
                df.index = range(1, len(df) + 1)  # 줄 번호를 1부터 표시
                st.table(df)
                continue

            st.error(errors.get(idx) or "🚨 인식불가 데이터가 있거나 8개가 아닙니다. 아래 칸을 직접 수정한 후 저장하세요.")
            # 갈래 B: st.data_editor 대체 (Streamlit 1.12) - 폼 + 텍스트 입력
            with st.form(key=f"edit_form_{exam_no}_{idx}"):
                date_in = st.text_input(
                    "검사일자 (YYYY-MM-DD)",
                    value="" if not records or is_blank(records[0].get("검사일자")) else str(records[0].get("검사일자")),
                    key=f"{run}_{idx}_date",
                )
                edited = []
                for i in range(max(8, len(records))):
                    rec = records[i] if i < len(records) else {}
                    key = f"{run}_{idx}_{i}"
                    st.markdown(f"**[{i + 1}회차 측정기록]**")  # 화면 번호는 1~8
                    edited.append([
                        st.text_input(
                            "검사시각 (HH:MM:SS)",
                            value="" if is_blank(rec.get("검사시각")) else str(rec.get("검사시각")),
                            key=f"{key}_time",
                        ),
                        st.text_input("압력 (kPa)", value=num_text(rec.get("압력")), key=f"{key}_p"),
                        st.text_input("온도 (c)", value=num_text(rec.get("온도")), key=f"{key}_t"),
                    ])
                submitted = st.form_submit_button("수정된 데이터 구글 시트에 강제 저장")

            if submitted:
                try:  # 입력값 3개를 원래 포맷("16:09:37", "5.7kPa", "27.7c")으로 결합
                    fixed = dict(paper, 측정기록=[
                        {"검사일자": date_in.strip(), **join_row(i + 1, *cells)}
                        for i, cells in enumerate(edited) if i < 8 or any(v.strip() for v in cells)  # 1~8번은 빈칸이어도 자리 유지
                    ])
                except ValueError as e:
                    st.error(f"입력 오류: {e}")
                else:
                    fixed["최종_합격_여부"] = judge(fixed)  # 수정된 압력값 기준으로 재판정
                    try:
                        with st.spinner("구글 시트에 데이터를 저장하는 중..."):
                            save_paper(fixed)
                    except Exception as e:
                        errors[idx] = f"구글 시트 저장 실패: {e}"
                        st.error(errors[idx])
                    else:
                        papers[idx] = fixed
                        saved[idx] = "✅ 수정본 저장 완료"
                        st.experimental_rerun()
