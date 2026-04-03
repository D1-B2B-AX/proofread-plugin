"""
교안 검수 결과 HTML 생성 스크립트 v2.1
- v2.1: 체크박스 UI + 선택 기반 메일 구성 + 데이터 수집 + 소요시간 측정
"""

import webbrowser
import os
import re
from datetime import datetime

# ── 설정 ──
PLUGIN_VERSION = "2.1"
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyVDzrw7ERUc6wmtfDepa0xZEklU_dTjoDpQRQP7rXjtJIxENPiC1AcK3N159dGvrtHkg/exec"


def escape_html(text):
    """HTML 특수문자 이스케이프"""
    if not text:
        return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def safe_html(text):
    """HTML 이스케이프 후, 검수 결과용 허용 태그만 복원"""
    if not text:
        return ""
    escaped = escape_html(text)
    escaped = escaped.replace('&lt;span class=&quot;highlight&quot;&gt;', '<span class="highlight">')
    escaped = escaped.replace('&lt;/span&gt;', '</span>')
    escaped = escaped.replace('&lt;b&gt;', '<b>')
    escaped = escaped.replace('&lt;/b&gt;', '</b>')
    return escaped


def strip_tags(text):
    """HTML 태그 제거하여 순수 텍스트 반환 (데이터 속성용)"""
    if not text:
        return ""
    return re.sub(r'<[^>]+>', '', str(text))


def auto_highlight(original, suggestion):
    """original/suggestion에 태그가 없으면 글자 diff로 자동 추가"""
    if not original or not suggestion:
        return original or "", suggestion or ""

    orig_plain = strip_tags(str(original))
    sugg_plain = strip_tags(str(suggestion))

    # 이미 태그가 있으면 그대로 반환
    if "<span" in str(original) or "<b>" in str(suggestion):
        return original, suggestion

    import difflib
    sm = difflib.SequenceMatcher(None, orig_plain, sugg_plain)
    orig_parts = []
    sugg_parts = []

    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            orig_parts.append(escape_html(orig_plain[i1:i2]))
            sugg_parts.append(escape_html(sugg_plain[j1:j2]))
        elif op == "replace":
            orig_parts.append(f'<span class="highlight">{escape_html(orig_plain[i1:i2])}</span>')
            sugg_parts.append(f'<b>{escape_html(sugg_plain[j1:j2])}</b>')
        elif op == "delete":
            orig_parts.append(f'<span class="highlight">{escape_html(orig_plain[i1:i2])}</span>')
        elif op == "insert":
            sugg_parts.append(f'<b>{escape_html(sugg_plain[j1:j2])}</b>')

    return "".join(orig_parts), "".join(sugg_parts)


def generate_html(review_data, output_path=None):
    """
    검수 결과 데이터를 HTML로 변환

    Args:
        review_data: dict {
            "file_name": 파일명,
            "total_pages": 총 페이지수,
            "client": 고객사명,
            "course": 과정명,
            "instructor": 강사명,
            "om_name": 검수자 이름 (없으면 "OOO"),
            "file_type": "ppt" | "pdf" | "notion" | "text",
            "errors": [...],
            "warnings": [...],
            "notes": [...],
            "mail_subject": 메일 제목,
            "mail_body": 메일 본문 (v2.1에서는 동적 생성으로 대체)
        }
        output_path: HTML 저장 경로 (None이면 자동 생성)

    Returns:
        str: 생성된 HTML 파일 경로
    """
    # om_name: HTML 모달에서 입력받으므로 기본값 허용
    if not review_data.get("om_name"):
        review_data["om_name"] = ""

    # total_pages: 추출 스크립트가 저장한 임시 파일에서 보정
    import tempfile
    if not review_data.get("total_pages"):
        page_count_path = os.path.join(tempfile.gettempdir(), "proofread_page_count.txt")
        if os.path.exists(page_count_path):
            with open(page_count_path, "r") as f:
                try:
                    review_data["total_pages"] = int(f.read().strip())
                except ValueError:
                    pass

    if output_path is None:
        date_str = datetime.now().strftime("%Y%m%d")
        safe_name = review_data.get("file_name", "교안").replace(".pptx", "").replace(".pdf", "").replace(".md", "")
        home = os.path.expanduser("~")
        temp_dir = os.path.join(home, "Desktop", "claude_temp")
        os.makedirs(temp_dir, exist_ok=True)
        output_path = os.path.join(temp_dir, f"검수결과_{safe_name}_{date_str}.html")

    def _page_sort_key(item):
        nums = re.findall(r'\d+', str(item.get("page", "0")))
        return int(nums[0]) if nums else 0

    # 분류 고정 규칙 강제 적용
    _must_be_warning = {"개인정보(이메일)", "개인정보(전화번호)", "이메일", "전화번호"}
    raw_errors = review_data.get("errors", [])
    raw_warnings = review_data.get("warnings", [])
    fixed_errors = []
    for item in raw_errors:
        if item.get("type", "") in _must_be_warning:
            raw_warnings.append(item)
        else:
            fixed_errors.append(item)

    errors = sorted(fixed_errors, key=_page_sort_key)
    warnings = sorted(raw_warnings, key=_page_sort_key)
    notes = sorted(review_data.get("notes", []), key=_page_sort_key)

    error_count = len(errors)
    warning_count = len(warnings)
    note_count = len(notes)

    # 확인 필요 심각도순 정렬
    sev_order = {"상": 0, "중": 1, "하": 2}
    warnings_sorted = sorted(warnings, key=lambda x: (sev_order.get(x.get("severity", ""), 3), _page_sort_key(x)))
    warning_high_count = sum(1 for w in warnings_sorted if w.get("severity") == "상")
    selectable_count = error_count + warning_high_count

    client = escape_html(review_data.get("client", ""))
    course = escape_html(review_data.get("course", ""))
    instructor = escape_html(review_data.get("instructor", ""))
    om_name = escape_html(review_data.get("om_name", "OOO"))
    file_name = escape_html(review_data.get("file_name", ""))
    total_pages = review_data.get("total_pages", "")
    file_type = review_data.get("file_type", "ppt")
    review_date = datetime.now().strftime("%Y-%m-%d")
    page_count_js = total_pages if total_pages else 0

    type_notice = {
        "ppt": "페이지 번호는 PPT 왼쪽 슬라이드 목록 순서와 동일합니다. (PPT 교안 기준) 간혹 슬라이드 번호가 0부터 시작하는 PPT의 경우, 표시된 페이지와 1p 차이가 날 수 있습니다.",
        "pdf": "페이지 번호는 PDF 뷰어 페이지 번호와 동일합니다. PDF 특성상 띄어쓰기가 누락될 수 있으며, 이는 원본 오류가 아닙니다.",
        "notion": "텍스트 직접 입력이므로 위치 정보는 제공되지 않습니다.",
        "text": "텍스트 직접 입력이므로 위치 정보는 제공되지 않습니다."
    }.get(file_type, "")

    pages_display = f" ({total_pages}p)" if total_pages else ""

    # ── 오류 확정 테이블 행 (체크박스 포함) ──
    error_rows = ""
    for i, item in enumerate(errors, 1):
        sev = item.get("severity", "중")
        if sev == "상":
            sev_html = '<span class="severity-high">상</span>'
        elif sev == "중":
            sev_html = '<span class="severity-mid">중</span>'
        else:
            sev_html = '<span class="severity-low">하</span>'

        # 볼드/하이라이트 자동 처리
        hl_orig, hl_sugg = auto_highlight(item.get("original", ""), item.get("suggestion", ""))

        cb_page = escape_html(item.get("page", "-"))
        cb_orig = escape_html(strip_tags(item.get("original", "")))
        cb_sugg = escape_html(strip_tags(item.get("suggestion", "")))

        error_rows += f"""      <tr class="row-error"><td class="cb-cell"><input type="checkbox" class="item-check error-check" checked data-page="{cb_page}" data-original="{cb_orig}" data-suggestion="{cb_sugg}"></td><td>{i}</td><td class="page-num">{item.get("page", "-")}</td><td>{escape_html(item.get("type", ""))}</td><td>{safe_html(hl_orig)}</td><td>{safe_html(hl_sugg)}</td><td>{sev_html}</td><td>{escape_html(item.get("location", ""))}</td></tr>\n"""

    # ── 확인 필요 테이블 행 (심각도 "상"만 체크박스) ──
    warning_rows = ""
    last_was_high = False
    for i, item in enumerate(warnings_sorted, 1):
        sev = item.get("severity", "")
        is_high = (sev == "상")

        if last_was_high and not is_high:
            col_count = 8
            warning_rows += f"""      <tr class="row-om-divider">{"<td></td>" * col_count}</tr>\n"""

        if is_high:
            sev_html = '<span class="severity-high">상</span>'
            row_class = "row-om-high"
            cb_page = escape_html(item.get("page", "-"))
            cb_orig = escape_html(strip_tags(item.get("original", "")))
            cb_sugg = escape_html(strip_tags(item.get("suggestion", "")))
            cb_html = f'<td class="cb-cell"><input type="checkbox" class="item-check warning-high-check" checked data-page="{cb_page}" data-original="{cb_orig}" data-suggestion="{cb_sugg}"></td>'
        elif sev == "중":
            sev_html = '<span class="severity-mid">중</span>'
            row_class = "row-om"
            cb_html = '<td class="cb-cell"></td>'
        elif sev == "하":
            sev_html = '<span class="severity-low">하</span>'
            row_class = "row-om"
            cb_html = '<td class="cb-cell"></td>'
        else:
            sev_html = '<span class="severity-low">-</span>'
            row_class = "row-om"
            cb_html = '<td class="cb-cell"></td>'

        # 볼드/하이라이트 자동 처리
        hl_orig, hl_sugg = auto_highlight(item.get("original", ""), item.get("suggestion", ""))

        warning_rows += f"""      <tr class="{row_class}">{cb_html}<td>{i}</td><td class="page-num">{item.get("page", "-")}</td><td>{escape_html(item.get("type", ""))}</td><td>{safe_html(hl_orig)}</td><td>{safe_html(hl_sugg)}</td><td>{sev_html}</td><td>{escape_html(item.get("location", ""))}</td></tr>\n"""
        last_was_high = is_high

    # ── 참고사항 테이블 행 (체크박스 없음) ──
    note_rows = ""
    for i, item in enumerate(notes, 1):
        hl_orig, hl_sugg = auto_highlight(item.get("original", ""), item.get("suggestion", ""))
        note_rows += f"""      <tr class="row-ref"><td>{i}</td><td class="page-num">{item.get("page", "-")}</td><td>{escape_html(item.get("type", ""))}</td><td>{safe_html(hl_orig)}</td><td>{safe_html(hl_sugg)}</td><td>{escape_html(item.get("note", ""))}</td></tr>\n"""

    # ── 오류 확정 섹션 ──
    if error_count > 0:
        error_section = f"""  <div class="section-error">
    <div class="section-title">오류 확정 <button class="select-ctrl" onclick="toggleAll('error-check')">전체 선택/해제</button></div>
    <table>
      <tr><th class="cb-header"></th><th>#</th><th>페이지</th><th>유형</th><th>원문</th><th>수정 제안</th><th>심각도</th><th>위치</th></tr>
{error_rows}    </table>
  </div>"""
    else:
        error_section = """  <div class="section-error">
    <div class="section-title">오류 확정</div>
    <div class="no-error">오류가 발견되지 않았습니다.</div>
  </div>"""

    # ── 확인 필요 섹션 ──
    if warning_count > 0:
        whc_label = f' <span style="font-size:12px; color:#888; font-weight:400;">(심각도 상: 메일 포함 대상)</span>' if warning_high_count > 0 else ''
        warning_section = f"""  <div class="section-om">
    <div class="section-title">확인 필요{whc_label}</div>
    <table>
      <tr><th class="cb-header"></th><th>#</th><th>페이지</th><th>유형</th><th>발견 내용</th><th>확인 포인트</th><th>심각도</th><th>위치</th></tr>
{warning_rows}    </table>
  </div>"""
    else:
        warning_section = ""

    # ── 참고사항 섹션 ──
    if note_count > 0:
        note_section = f"""  <div class="section-ref">
    <div class="section-title">참고사항</div>
    <table>
      <tr><th>#</th><th>페이지</th><th>유형</th><th>현재 표기</th><th>권장 표기</th><th>비고</th></tr>
{note_rows}    </table>
  </div>"""
    else:
        note_section = ""

    # ── 메일 섹션 (체크박스 기반 동적 구성) ──
    if error_count > 0 or warning_high_count > 0:
        mail_subject = review_data.get("mail_subject", f"{client}_{course} 교안 확인 요청드립니다")
        mail_section = f"""  <div class="mail-section">
    <h3>강사 수정 요청 메일 초안 <span class="selected-count" id="selectedCount"></span></h3>
    <div class="mail-content mail-body" id="mailBody"></div>
    <div class="btn-row">
      <button class="copy-btn" onclick="copyMail(event)">메일 복사</button>
    </div>
  </div>"""
    else:
        mail_section = ""

    # ── HTML 생성 ──
    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>교안 검수 결과</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: "Pretendard", -apple-system, sans-serif; background: #f5f5f5; padding: 30px; color: #333; }}
  .container {{ max-width: 1000px; margin: 0 auto; background: #fff; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); overflow: hidden; }}
  .header-area {{ background: #13299f; padding: 40px 40px 28px; color: #fff; position: relative; }}
  .detail-area {{ background: #fff; padding: 28px 40px 40px; }}
  h1 {{ font-size: 22px; margin-bottom: 8px; }}
  .subtitle {{ color: rgba(255,255,255,0.7); font-size: 14px; margin-bottom: 30px; }}
  .info-box {{ background: rgba(255,255,255,0.12); border-radius: 8px; padding: 16px 20px; margin-bottom: 24px; font-size: 14px; line-height: 1.8; color: #fff; }}
  .info-box span {{ color: rgba(255,255,255,0.6); }}
  .summary-cards {{ display: flex; gap: 12px; margin-bottom: 30px; }}
  .card {{ flex: 1; border-radius: 8px; padding: 16px; text-align: center; }}
  .card .num {{ font-size: 28px; font-weight: bold; }}
  .card .label {{ font-size: 12px; margin-top: 4px; }}
  .card-error {{ background: rgba(255,255,255,0.95); color: #d32f2f; border: none; }}
  .card-om {{ background: rgba(255,255,255,0.95); color: #e67700; border: none; }}
  .card-ref {{ background: rgba(255,255,255,0.95); color: #888; border: none; }}
  .section-title {{ font-size: 16px; font-weight: bold; margin: 28px 0 12px; padding-left: 10px; display: flex; align-items: center; gap: 10px; }}
  .section-error .section-title {{ border-left: 4px solid #d32f2f; color: #d32f2f; }}
  .section-om .section-title {{ border-left: 4px solid #e67700; color: #e67700; }}
  .section-ref .section-title {{ border-left: 4px solid #999; color: #666; }}
  table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 13px; }}
  th {{ background: #f8f9fa; padding: 10px 12px; text-align: left; font-weight: 600; border-bottom: 2px solid #e0e0e0; }}
  td {{ padding: 10px 12px; border-bottom: 1px solid #eee; vertical-align: top; border-right: 1px solid #f0f0f0; }}
  td:last-child {{ border-right: none; }}
  tr:hover {{ background: #fafafa; }}
  .severity-high {{ display: inline-block; background: #ffebee; color: #d32f2f; font-weight: bold; padding: 2px 10px; border-radius: 12px; font-size: 12px; }}
  .severity-mid {{ display: inline-block; background: #fff3e0; color: #e67700; font-weight: bold; padding: 2px 10px; border-radius: 12px; font-size: 12px; }}
  .severity-low {{ display: inline-block; background: #f5f5f5; color: #999; font-weight: bold; padding: 2px 10px; border-radius: 12px; font-size: 12px; }}
  .highlight {{ background: #fff3cd; padding: 1px 4px; border-radius: 3px; }}
  .page-num {{ color: #13299f; font-weight: 600; }}
  .row-error td:first-child {{ border-left: 3px solid #d32f2f; }}
  .row-om td:first-child {{ border-left: 3px solid #e67700; }}
  .row-om-high td:first-child {{ border-left: 3px solid #e67700; }}
  .row-om-divider td {{ border-bottom: 1px solid #ccc; border-right: none; padding: 0; height: 4px; }}
  .row-om-divider td:first-child {{ border-left: 3px solid transparent; }}
  .row-ref td:first-child {{ border-left: 3px solid #ccc; }}
  h1 .divider {{ color: rgba(255,255,255,0.4); font-weight: 300; margin: 0 6px; }}
  .note {{ background: #eef2ff; border-radius: 6px; padding: 12px 16px; font-size: 13px; color: #13299f; margin-bottom: 20px; }}
  .no-error {{ background: #e8f5e9; border-radius: 8px; padding: 20px; text-align: center; color: #2e7d32; font-size: 15px; margin-bottom: 20px; }}
  .mail-section {{ background: #f8f9fa; border-radius: 8px; padding: 20px; margin-top: 30px; }}
  .mail-section h3 {{ font-size: 15px; margin-bottom: 12px; }}
  .mail-content {{ background: #fff; border: 1px solid #e0e0e0; border-radius: 6px; padding: 16px; font-size: 13px; line-height: 1.8; white-space: pre-wrap; }}
  .btn-row {{ display: flex; gap: 8px; margin-top: 10px; }}
  .copy-btn {{ padding: 8px 20px; background: #13299f; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; }}
  .copy-btn:hover {{ background: #0e1f7a; }}
  .copy-btn-sm {{ padding: 4px 12px; background: #e0e0e0; color: #333; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; }}
  .copy-btn-sm:hover {{ background: #bdbdbd; }}
  .copied {{ background: #4caf50 !important; color: #fff !important; }}
  /* ── v2.1 체크박스 스타일 ── */
  .cb-cell {{ text-align: center; width: 36px; padding: 10px 4px !important; }}
  .cb-header {{ text-align: center; width: 36px; }}
  .item-check {{ width: 16px; height: 16px; cursor: pointer; accent-color: #13299f; }}
  .select-ctrl {{ font-size: 11px; font-weight: 400; color: #888; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; padding: 2px 8px; cursor: pointer; }}
  .select-ctrl:hover {{ background: #e0e0e0; }}
  .selected-count {{ font-size: 13px; font-weight: 600; color: #13299f; }}
  .data-msg {{ font-size: 11px; color: #4caf50; margin-top: 6px; display: none; }}
  /* ── 이름 입력 모달 ── */
  .modal-overlay {{ position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.35); display: none; justify-content: center; align-items: center; z-index: 1000; }}
  .modal-box {{ background: #fff; border-radius: 10px; padding: 28px 32px; width: 340px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }}
  .modal-title {{ font-size: 15px; font-weight: 600; color: #333; margin-bottom: 16px; }}
  .modal-input {{ width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 14px; font-family: inherit; outline: none; }}
  .modal-input:focus {{ border-color: #13299f; }}
  .modal-btns {{ display: flex; gap: 8px; margin-top: 16px; justify-content: flex-end; }}
  .modal-confirm {{ padding: 8px 20px; background: #13299f; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-family: inherit; }}
  .modal-confirm:hover {{ background: #0e1f7a; }}
  .modal-cancel {{ padding: 8px 16px; background: #f0f0f0; color: #666; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; font-family: inherit; }}
  .modal-cancel:hover {{ background: #e0e0e0; }}
  @media (max-width: 768px) {{
    body {{ padding: 10px; }}
    .header-area {{ padding: 24px 20px 20px; }}
    .detail-area {{ padding: 20px; }}
    h1 {{ font-size: 17px; padding-right: 80px; }}
    .summary-cards {{ flex-direction: column; }}
    table {{ font-size: 12px; }}
    td, th {{ padding: 8px 6px; }}
  }}
</style>
</head>
<body>
<div class="container">
  <div class="header-area">
  <h1>「{course}」<span class="divider">|</span><span style="color: #01b1e3;">교안 검수 결과</span></h1>
  <p class="subtitle">{client} &middot; {instructor} 강사 &middot; {review_date} 검수 &middot; {om_name}</p>

  <div class="info-box">
    <span>파일:</span> <b>{file_name}{pages_display}</b><br>
    <span>고객사:</span> <b>{client}</b> &nbsp;|&nbsp; <span>과정:</span> <b>{course}</b> &nbsp;|&nbsp; <span>강사:</span> <b>{instructor}</b>
  </div>

  <div class="summary-cards">
    <div class="card card-error"><div class="num">{error_count}</div><div class="label">오류 확정</div></div>
    <div class="card card-om"><div class="num">{warning_count}</div><div class="label">확인 필요</div></div>
    <div class="card card-ref"><div class="num">{note_count}</div><div class="label">참고사항</div></div>
  </div>
  </div>

  <div class="detail-area">
  <div class="note">※ {type_notice}</div>

{error_section}

{warning_section}

{note_section}

{mail_section}
  <div class="data-msg" id="dataMsg"></div>
  </div>
</div>

<!-- 이름 입력 모달 -->
<div class="modal-overlay" id="nameModal">
  <div class="modal-box">
    <div class="modal-title" id="modalTitle">발신자로 기입될 이름을 입력해주세요</div>
    <input type="text" class="modal-input" id="nameInput" placeholder="예: 홍길동">
    <div class="modal-btns">
      <button class="modal-cancel" onclick="cancelName()">취소</button>
      <button class="modal-confirm" onclick="confirmName()">확인</button>
    </div>
  </div>
</div>

<script>
// ── 메타데이터 ──
var META = {{
  omName: "{om_name}",
  pluginVersion: "{PLUGIN_VERSION}",
  client: "{client}",
  course: "{course}",
  instructor: "{instructor}",
  mailSubject: "{escape_html(mail_subject)}",
  fileName: "{file_name}",
  fileType: "{escape_html(file_type)}",
  pageCount: {page_count_js},
  foundConfirmed: {error_count},
  foundReview: {warning_count},
  foundReference: {note_count},
  selectableCount: {selectable_count}
}};

var APPS_SCRIPT_URL = "{APPS_SCRIPT_URL}";
var PAGE_LOAD_TIME = Date.now();
var dataSent = false;
var nameConfirmed = false;
var cancelCount = 0;

// ── 메일 본문 동적 생성 ──
function updateMailBody() {{
  var checked = document.querySelectorAll('.item-check:checked');
  var count = checked.length;
  var countEl = document.getElementById('selectedCount');
  if (countEl) {{
    countEl.textContent = count > 0 ? '(' + count + '건 선택)' : '(선택 없음)';
  }}

  var bodyEl = document.getElementById('mailBody');
  if (!bodyEl) return;

  if (count === 0) {{
    bodyEl.textContent = '선택된 항목이 없습니다. 위 표에서 강사에게 요청할 항목을 선택해주세요.';
    return;
  }}

  var lines = [];
  lines.push('[제목] ' + META.mailSubject);
  lines.push('');
  lines.push(META.instructor + ' 강사님, 안녕하세요.');
  lines.push(META.client + ' ' + META.course + ' 교안 준비에 힘써 주셔서 감사합니다.');
  lines.push('');
  lines.push('교안 검토 중 아래 사항이 확인되어 말씀드립니다.');
  lines.push('확인 후 수정이 필요한 부분이 있으시면 반영 부탁드리겠습니다.');
  lines.push('');
  checked.forEach(function(cb, i) {{
    lines.push((i + 1) + '. ' + cb.dataset.page + ' "' + cb.dataset.original + '" \\u2192 "' + cb.dataset.suggestion + '"');
  }});
  lines.push('');
  lines.push('\\u203B 일부 항목은 의도하신 표현일 수 있어 확인 차 여쭙습니다.');
  lines.push('수정 반영 후 파일을 회신해 주시면 최종 확인 후 진행하겠습니다.');
  lines.push('바쁘신 중에 번거로우시겠지만 확인 부탁드리겠습니다.');
  lines.push('');
  lines.push('감사합니다.');
  lines.push(META.omName + ' 드림');

  bodyEl.textContent = lines.join('\\n');
}}

// ── 전체 선택/해제 토글 ──
function toggleAll(className) {{
  var cbs = document.querySelectorAll('.' + className);
  var allChecked = true;
  cbs.forEach(function(cb) {{ if (!cb.checked) allChecked = false; }});
  cbs.forEach(function(cb) {{ cb.checked = !allChecked; }});
  updateMailBody();
}}

// ── 데이터 전송 (fetch no-cors + sendBeacon 폴백) ──
function sendData() {{
  if (dataSent) return;
  dataSent = true;

  var selected = document.querySelectorAll('.item-check:checked').length;
  var durationSec = Math.round((Date.now() - PAGE_LOAD_TIME) / 1000);

  var payload = {{
    timestamp: new Date().toISOString(),
    om_name: META.omName,
    plugin_version: META.pluginVersion,
    client_name: META.client,
    course_name: META.course,
    instructor_name: META.instructor,
    file_name: META.fileName,
    file_type: META.fileType,
    page_count: META.pageCount,
    found_confirmed: META.foundConfirmed,
    found_review: META.foundReview,
    found_reference: META.foundReference,
    selectable_count: META.selectableCount,
    om_selected: selected,
    mail_copied: true,
    duration_sec: durationSec
  }};

  var jsonStr = JSON.stringify(payload);

  // 1차: fetch (text/plain으로 CORS preflight 회피)
  fetch(APPS_SCRIPT_URL, {{
    method: 'POST',
    mode: 'no-cors',
    headers: {{ 'Content-Type': 'text/plain' }},
    body: jsonStr
  }}).catch(function() {{
    // 2차 폴백: sendBeacon
    if (navigator.sendBeacon) {{
      navigator.sendBeacon(APPS_SCRIPT_URL, jsonStr);
    }}
  }});

  // 데이터 전송 완료 (UI 표시 없음)
}}

// ── 이름 모달 ──
function showNameModal() {{
  var modal = document.getElementById('nameModal');
  var input = document.getElementById('nameInput');
  var title = document.getElementById('modalTitle');
  cancelCount = 0;

  if (!nameConfirmed) {{
    title.textContent = '발신자로 기입될 이름을 입력해주세요';
    input.value = '';
    input.placeholder = '예: 홍길동';
  }} else {{
    title.textContent = '발신자: ' + META.omName;
    input.value = META.omName;
    input.placeholder = '';
  }}
  modal.style.display = 'flex';
  input.focus();
  input.select();
}}

function confirmName() {{
  var input = document.getElementById('nameInput');
  var val = input.value.trim();
  if (!val) {{
    handleEmptyName();
    return;
  }}
  META.omName = val;
  nameConfirmed = true;
  cancelCount = 0;
  dataSent = false;
  closeNameModal();
  updateMailBody();
  doCopyMail();
}}

function handleEmptyName() {{
  cancelCount++;
  if (cancelCount >= 2) {{
    META.omName = '미입력';
    nameConfirmed = true;
    dataSent = false;
    closeNameModal();
    updateMailBody();
    doCopyMail();
  }} else {{
    document.getElementById('modalTitle').textContent = '이름 없이 진행하면 미입력으로 기록됩니다';
    document.getElementById('nameInput').focus();
  }}
}}

function cancelName() {{
  handleEmptyName();
}}

function closeNameModal() {{
  document.getElementById('nameModal').style.display = 'none';
}}

// ── 메일 복사 ──
function copyMail(e) {{
  showNameModal();
}}

function doCopyMail() {{
  var text = document.getElementById('mailBody').innerText;
  var btn = document.querySelector('.copy-btn');
  navigator.clipboard.writeText(text).then(function() {{
    btn.textContent = '복사 완료!';
    btn.classList.add('copied');
    setTimeout(function() {{ btn.textContent = '메일 복사'; btn.classList.remove('copied'); }}, 1500);
    sendData();
  }});
}}

// ── 초기화 ──
document.addEventListener('DOMContentLoaded', function() {{
  document.querySelectorAll('.item-check').forEach(function(cb) {{
    cb.addEventListener('change', updateMailBody);
  }});
  updateMailBody();
  document.getElementById('nameInput').addEventListener('keydown', function(e) {{
    if (e.key === 'Enter') confirmName();
  }});
}});
</script>
</body>
</html>"""

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    # 검수 완료 후 안내 메시지 (코드 강제 출력)
    print(f"\nHTML 생성 완료: {output_path}")
    print()
    print("📋 결과 확인")
    print("- 결과 페이지는 브라우저에서 바로 확인할 수 있어요.")
    print("- 메일 복사는 결과 페이지 하단 \"메일 복사\" 버튼을 이용해 주세요.")
    print()
    print("🔄 추가 작업")
    print("- 다른 교안 검수: /교안검수를 다시 입력")
    print("- 강사 수정본 확인: \"2차 검수\"라고 입력")
    print()
    print("🚪 종료")
    print("- 결과 확인이 끝났으면 이 대화창은 닫아도 됩니다.")

    return output_path


def open_in_browser(html_path):
    """생성된 HTML을 브라우저에서 열기"""
    webbrowser.open(f"file:///{html_path.replace(os.sep, '/')}")


if __name__ == "__main__":
    sample_data = {
        "file_name": "ppt 스킬업 한번에 끝내는 제작 스킬 3차수 강의교안_오류테스트.pptx",
        "total_pages": 224,
        "client": "HL그룹",
        "course": "한번에 끝내는 파워포인트 디자인 제작 스킬",
        "instructor": "이지훈",
        "om_name": "전정현",
        "file_type": "ppt",
        "errors": [
            {"page": "p.75", "type": "오탈자", "severity": "상",
             "original": '최첨단 물류 <span class="highlight">시서템</span>',
             "suggestion": "시<b>스</b>템", "location": "[상] 슬라이드 제목"},
            {"page": "p.75", "type": "오탈자", "severity": "상",
             "original": '디지털 전환 <span class="highlight">인프로</span>',
             "suggestion": "인프<b>라</b>", "location": "[중] 카테고리 제목"},
            {"page": "p.80", "type": "오탈자", "severity": "중",
             "original": '의류용 <span class="highlight">섬요</span>로 사용하던',
             "suggestion": "섬<b>유</b>", "location": "[하] 산업용 섬유 설명"},
            {"page": "p.84", "type": "오탈자", "severity": "중",
             "original": '형태안정성 및 <span class="highlight">가경</span>',
             "suggestion": "가<b>격</b>", "location": "[하] 타이어코드 사용 이유"},
            {"page": "p.87", "type": "오탈자", "severity": "상",
             "original": '투자자와의 <span class="highlight">동반성자</span>',
             "suggestion": "동반성<b>장</b>", "location": "[상] 투자 어필 제목"},
        ],
        "warnings": [
            {"page": "p.24", "type": "유사 단어", "severity": "중",
             "original": '보안 <span class="highlight">강력</span>',
             "suggestion": '같은 페이지 "서비스강화", "생산성향상" 패턴 → "강화"가 자연스러움',
             "location": "[중]"},
            {"page": "p.183→184", "type": "문맥 모순", "severity": "상",
             "original": '<span class="highlight">4가지</span> → <span class="highlight">3가지</span> 프로세스',
             "suggestion": "같은 제목인데 수 불일치", "location": "[상]"},
        ],
        "notes": [
            {"page": "p.6~7", "type": "외래어 표기",
             "original": "프레젠테이션 / 프리젠테이션 혼용",
             "suggestion": "프레젠테이션 (국립국어원)", "note": "본문과 글꼴 예시에서 혼용"},
            {"page": "p.63", "type": "외래어 표기",
             "original": "컨텐츠", "suggestion": "콘텐츠 (국립국어원)", "note": "실무에서 혼용"},
        ],
        "mail_subject": "HL그룹_한번에 끝내는 파워포인트 디자인 제작 스킬 교안 확인 요청드립니다",
    }

    path = generate_html(sample_data)
    print(f"HTML 생성 완료: {path}")
    open_in_browser(path)
