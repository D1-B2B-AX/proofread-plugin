"""
PDF 교안 텍스트 추출 스크립트
- .pdf 파일에서 페이지별 텍스트를 추출
- 일반 텍스트와 표를 분리 추출
- 각 텍스트의 페이지 내 위치(상/중/하) 정보 포함
- 반복 텍스트 변형 감지 (유사 문장 중 소수 변형 자동 플래그)
- 이미지 위주 페이지 감지 및 안내
"""

import sys
import os
import json
import re
import pdfplumber

# 공통 감지 유틸리티 import
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from detection_utils import detect_text_variants, detect_suspicious_words, detect_suspicious_english, detect_broken_text, detect_suspected_spacing

# 브라우저 인쇄 시 자동 삽입되는 머리글/바닥글 패턴 (URL, 날짜, 페이지번호 등)
BROWSER_HEADER_FOOTER = re.compile(
    r'^https?://\S+$'           # URL만 있는 줄
    r'|^\d{1,4}/\d{1,2}/\d{1,2}$'  # 날짜만 있는 줄 (2026/03/16)
    r'|^\d+\s*/\s*\d+$'        # 페이지번호만 있는 줄 (1 / 5)
)


def get_position_label(y_top, page_height):
    """텍스트의 y좌표와 페이지 전체 높이로 상/중/하 판별"""
    ratio = y_top / page_height if page_height > 0 else 0
    if ratio < 0.33:
        return "상"
    elif ratio < 0.66:
        return "중"
    else:
        return "하"


def extract_pdf(file_path):
    """
    PDF 파일에서 페이지별 텍스트를 추출

    Returns:
        dict: {
            "file_name": 파일명,
            "total_pages": 총 페이지 수,
            "image_only_pages": 이미지 위주 페이지 수,
            "pages": [
                {
                    "page_number": 페이지 번호,
                    "texts": [
                        {"text": 텍스트, "position": "상/중/하"},
                        ...
                    ],
                    "is_image_only": 이미지 위주 여부
                },
                ...
            ]
        }
    """
    result = {
        "file_name": file_path.split("/")[-1].split("\\")[-1],
        "total_pages": 0,
        "image_only_pages": 0,
        "pages": []
    }

    with pdfplumber.open(file_path) as pdf:
        result["total_pages"] = len(pdf.pages)

        for i, page in enumerate(pdf.pages, 1):
            page_items = []
            is_image_only = False
            page_height = page.height

            # 표 추출 (텍스트보다 먼저 추출하여 중복 제거에 활용)
            tables = page.extract_tables()
            table_bboxes = page.find_tables()

            # 표 영역을 제외한 텍스트 추출 (중복 방지)
            if table_bboxes:
                try:
                    filtered_page = page
                    for table_obj in table_bboxes:
                        # bbox가 페이지 경계를 벗어나는 경우 클리핑
                        x0, y0, x1, y1 = table_obj.bbox
                        px0, py0, px1, py1 = page.bbox
                        clipped_bbox = (
                            max(x0, px0), max(y0, py0),
                            min(x1, px1), min(y1, py1)
                        )
                        filtered_page = filtered_page.outside_bbox(clipped_bbox)
                    text = filtered_page.extract_text()
                except (ValueError, Exception):
                    # 표 영역 제외 실패 시 전체 텍스트 추출 (중복 감수)
                    text = page.extract_text()
            else:
                text = page.extract_text()

            # 이미지 위주 페이지 판별
            # 텍스트가 거의 없고 (30자 미만) 표도 없으면 이미지 위주로 판단
            text_length = len(text.strip()) if text else 0
            if text_length < 30 and not tables:
                is_image_only = True
                result["image_only_pages"] += 1

            # 텍스트 처리 (줄 단위 위치 추정)
            if text and text.strip():
                lines = [line.strip() for line in text.split("\n")
                         if line.strip() and not BROWSER_HEADER_FOOTER.match(line.strip())]
                # 각 줄의 위치를 words 좌표에서 추정
                line_positions = _estimate_line_positions(page, table_bboxes)
                for line in lines:
                    position = _find_line_position(line, line_positions, page_height)
                    page_items.append({"text": line, "position": position})

            # 표 데이터 처리
            if tables and table_bboxes:
                for table_idx, table in enumerate(tables):
                    # 표의 y좌표로 위치 판별
                    if table_idx < len(table_bboxes):
                        table_top = table_bboxes[table_idx].bbox[1]
                        position = get_position_label(table_top, page_height)
                    else:
                        position = "중"

                    for row in table:
                        row_texts = [str(cell).strip() for cell in row if cell is not None and str(cell).strip()]
                        if row_texts:
                            page_items.append({"text": "[표] " + " | ".join(row_texts), "position": position})

            result["pages"].append({
                "page_number": i,
                "texts": page_items,
                "is_image_only": is_image_only
            })

    # 반복 텍스트 변형 감지 (줄 단위)
    result["suspected_variants"] = detect_text_variants(result["pages"], key="page_number")

    # 형태소 분석 기반 의심 단어 감지
    result["suspicious_words"] = detect_suspicious_words(result["pages"], key="page_number")

    # 영문 스펠 체크 기반 의심 단어 감지
    result["suspicious_english"] = detect_suspicious_english(result["pages"], key="page_number")

    # 텍스트 깨짐 감지
    result["broken_texts"] = detect_broken_text(result["pages"], key="page_number")

    # 띄어쓰기 오류 감지
    result["suspected_spacing"] = detect_suspected_spacing(result["pages"], key="page_number")

    return result




def _estimate_line_positions(page, table_bboxes):
    """페이지 내 텍스트의 줄별 y좌표를 추정하여 반환"""
    positions = {}
    try:
        words = page.extract_words()
        if not words:
            return positions

        # 표 영역 내 words 제외
        table_rects = []
        if table_bboxes:
            for tb in table_bboxes:
                table_rects.append(tb.bbox)

        for word in words:
            # 표 영역 안에 있는 word는 건너뛰기
            in_table = False
            for (tx0, ty0, tx1, ty1) in table_rects:
                if tx0 <= word["x0"] <= tx1 and ty0 <= word["top"] <= ty1:
                    in_table = True
                    break
            if in_table:
                continue

            # 텍스트를 key로, y좌표를 값으로 저장 (첫 등장 기준)
            w_text = word["text"].strip()
            if w_text and w_text not in positions:
                positions[w_text] = word["top"]

    except Exception:
        pass

    return positions


def _find_line_position(line, line_positions, page_height):
    """줄 텍스트에 포함된 단어의 y좌표로 위치 판별"""
    # 줄에 포함된 단어 중 하나라도 좌표가 있으면 사용
    for word_text, y_top in line_positions.items():
        if word_text in line:
            return get_position_label(y_top, page_height)
    # 좌표를 찾지 못하면 기본값
    return "중"


def format_output(result):
    """추출 결과를 읽기 쉬운 텍스트 형식으로 변환"""
    output_lines = []
    output_lines.append(f"파일명: {result['file_name']}")
    output_lines.append(f"총 페이지: {result['total_pages']}p")

    if result["image_only_pages"] > 0:
        output_lines.append(f"이미지 위주 페이지: {result['image_only_pages']}p (텍스트 검수 불가)")

    output_lines.append("")

    for page in result["pages"]:
        page_num = page["page_number"]
        texts = page["texts"]
        is_image_only = page["is_image_only"]

        output_lines.append(f"=== 페이지 {page_num} ===")

        if is_image_only:
            output_lines.append("(이미지 위주 페이지 — 텍스트 검수 불가)")
        elif texts:
            for item in texts:
                output_lines.append(f"[{item['position']}] {item['text']}")
        else:
            output_lines.append("(텍스트 없음)")

        output_lines.append("")

    return "\n".join(output_lines)




if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python extract_pdf.py <파일경로.pdf> [--json]")
        sys.exit(1)

    file_path = sys.argv[1]
    use_json = "--json" in sys.argv

    result = extract_pdf(file_path)

    if use_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(format_output(result))
