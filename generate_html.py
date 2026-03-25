"""
교안 검수 결과 HTML 생성 스크립트
- 검수 결과 데이터를 받아 HTML 파일을 생성하고 브라우저에서 자동으로 열기
- Day1 메인 블루(#13299f) 컬러 매칭, Pretendard 웹폰트, 반응형 대응
"""

import webbrowser
import os
from datetime import datetime


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
    # 허용 태그 복원: <span class="highlight">, </span>, <b>, </b>
    escaped = escaped.replace('&lt;span class=&quot;highlight&quot;&gt;', '<span class="highlight">')
    escaped = escaped.replace('&lt;/span&gt;', '</span>')
    escaped = escaped.replace('&lt;b&gt;', '<b>')
    escaped = escaped.replace('&lt;/b&gt;', '</b>')
    return escaped


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
            "file_type": "ppt" | "pdf" | "notion" | "text",
            "errors": [  # 오류 확정
                {"page": "p.3", "type": 오류유형, "severity": "상/중/하",
                 "original": 원문(highlight 포함 가능), "suggestion": 수정안, "location": "[상] 위치설명"}
            ],
            "warnings": [  # 확인 필요
                {"page": "p.7", "type": 오류유형,
                 "original": 발견내용, "suggestion": 확인포인트, "location": "[중]"}
            ],
            "notes": [  # 참고사항
                {"page": "p.12", "type": 오류유형,
                 "original": 현재표기, "suggestion": 권장표기, "note": 비고}
            ],
            "mail_subject": 메일 제목,
            "mail_body": 메일 본문
        }
        output_path: HTML 저장 경로 (None이면 자동 생성)

    Returns:
        str: 생성된 HTML 파일 경로
    """
    if output_path is None:
        date_str = datetime.now().strftime("%Y%m%d")
        safe_name = review_data.get("file_name", "교안").replace(".pptx", "").replace(".pdf", "").replace(".md", "")
        home = os.path.expanduser("~")
        temp_dir = os.path.join(home, "Desktop", "claude_temp")
        os.makedirs(temp_dir, exist_ok=True)
        output_path = os.path.join(temp_dir, f"검수결과_{safe_name}_{date_str}.html")

    def _page_sort_key(item):
        """페이지 번호 오름차순 정렬 키 (p.3 → 3, p.12 → 12)"""
        import re as _re
        nums = _re.findall(r'\d+', str(item.get("page", "0")))
        return int(nums[0]) if nums else 0

    # 분류 고정 규칙 강제 적용: 오류 확정에 들어오면 안 되는 유형을 자동 이동
    _must_be_warning = {"개인정보(이메일)", "개인정보(전화번호)", "이메일", "전화번호"}
    raw_errors = review_data.get("errors", [])
    raw_warnings = review_data.get("warnings", [])
    fixed_errors = []
    for item in raw_errors:
        if item.get("type", "") in _must_be_warning:
            raw_warnings.append(item)  # 확인 필요로 이동
        else:
            fixed_errors.append(item)

    errors = sorted(fixed_errors, key=_page_sort_key)
    warnings = sorted(raw_warnings, key=_page_sort_key)
    notes = sorted(review_data.get("notes", []), key=_page_sort_key)

    error_count = len(errors)
    warning_count = len(warnings)
    note_count = len(notes)

    client = escape_html(review_data.get("client", ""))
    course = escape_html(review_data.get("course", ""))
    instructor = escape_html(review_data.get("instructor", ""))
    file_name = escape_html(review_data.get("file_name", ""))
    total_pages = review_data.get("total_pages", "")
    file_type = review_data.get("file_type", "ppt")
    review_date = datetime.now().strftime("%Y-%m-%d")

    # 교안 형태별 안내 문구
    type_notice = {
        "ppt": "페이지 번호는 PPT 왼쪽 슬라이드 목록 순서와 동일합니다. (PPT 교안 기준) 간혹 슬라이드 번호가 0부터 시작하는 PPT의 경우, 표시된 페이지와 1p 차이가 날 수 있습니다.",
        "pdf": "페이지 번호는 PDF 뷰어 페이지 번호와 동일합니다. PDF 특성상 띄어쓰기가 누락될 수 있으며, 이는 원본 오류가 아닙니다.",
        "notion": "텍스트 직접 입력이므로 위치 정보는 제공되지 않습니다.",
        "text": "텍스트 직접 입력이므로 위치 정보는 제공되지 않습니다."
    }.get(file_type, "")

    pages_display = f" ({total_pages}p)" if total_pages else ""

    # 오류 확정 테이블 행
    error_rows = ""
    for i, item in enumerate(errors, 1):
        sev = item.get("severity", "중")
        if sev == "상":
            sev_html = '<span class="severity-high">상</span>'
        elif sev == "중":
            sev_html = '<span class="severity-mid">중</span>'
        else:
            sev_html = '<span class="severity-low">하</span>'

        error_rows += f"""      <tr class="row-error"><td>{i}</td><td class="page-num">{item.get("page", "-")}</td><td>{escape_html(item.get("type", ""))}</td><td>{safe_html(item.get("original", ""))}</td><td>{safe_html(item.get("suggestion", ""))}</td><td>{sev_html}</td><td>{escape_html(item.get("location", ""))}</td></tr>\n"""

    # 확인 필요 테이블 행 (심각도 포함, 상→중→하→없음 순 정렬 후 페이지순)
    sev_order = {"상": 0, "중": 1, "하": 2}
    warnings_sorted = sorted(warnings, key=lambda x: (sev_order.get(x.get("severity", ""), 3), _page_sort_key(x)))
    warning_rows = ""
    last_was_high = False
    for i, item in enumerate(warnings_sorted, 1):
        sev = item.get("severity", "")
        is_high = (sev == "상")

        # 심각도 "상" → 일반으로 전환되는 지점에 구분선 삽입
        if last_was_high and not is_high:
            col_count = 7
            warning_rows += f"""      <tr class="row-om-divider">{"<td></td>" * col_count}</tr>\n"""

        if is_high:
            sev_html = '<span class="severity-high">상</span>'
            row_class = "row-om-high"
        elif sev == "중":
            sev_html = '<span class="severity-mid">중</span>'
            row_class = "row-om"
        elif sev == "하":
            sev_html = '<span class="severity-low">하</span>'
            row_class = "row-om"
        else:
            sev_html = '<span class="severity-low">-</span>'
            row_class = "row-om"

        warning_rows += f"""      <tr class="{row_class}"><td>{i}</td><td class="page-num">{item.get("page", "-")}</td><td>{escape_html(item.get("type", ""))}</td><td>{safe_html(item.get("original", ""))}</td><td>{safe_html(item.get("suggestion", ""))}</td><td>{sev_html}</td><td>{escape_html(item.get("location", ""))}</td></tr>\n"""
        last_was_high = is_high

    # 참고사항 테이블 행
    note_rows = ""
    for i, item in enumerate(notes, 1):
        note_rows += f"""      <tr class="row-ref"><td>{i}</td><td class="page-num">{item.get("page", "-")}</td><td>{escape_html(item.get("type", ""))}</td><td>{safe_html(item.get("original", ""))}</td><td>{safe_html(item.get("suggestion", ""))}</td><td>{escape_html(item.get("note", ""))}</td></tr>\n"""

    # 오류 확정 섹션
    if error_count > 0:
        error_section = f"""  <div class="section-error">
    <div class="section-title">오류 확정</div>
    <table>
      <tr><th>#</th><th>페이지</th><th>유형</th><th>원문</th><th>수정 제안</th><th>심각도</th><th>위치</th></tr>
{error_rows}    </table>
  </div>"""
    else:
        error_section = """  <div class="section-error">
    <div class="section-title">오류 확정</div>
    <div class="no-error">오류가 발견되지 않았습니다.</div>
  </div>"""

    # 확인 필요 섹션
    if warning_count > 0:
        warning_section = f"""  <div class="section-om">
    <div class="section-title">확인 필요</div>
    <table>
      <tr><th>#</th><th>페이지</th><th>유형</th><th>발견 내용</th><th>확인 포인트</th><th>심각도</th><th>위치</th></tr>
{warning_rows}    </table>
  </div>"""
    else:
        warning_section = ""

    # 참고사항 섹션
    if note_count > 0:
        note_section = f"""  <div class="section-ref">
    <div class="section-title">참고사항</div>
    <table>
      <tr><th>#</th><th>페이지</th><th>유형</th><th>현재 표기</th><th>권장 표기</th><th>비고</th></tr>
{note_rows}    </table>
  </div>"""
    else:
        note_section = ""

    # 메일 섹션 (오류 확정 1건 이상 또는 확인 필요(상) 1건 이상일 때)
    warning_high_count = sum(1 for w in warnings_sorted if w.get("severity") == "상")
    if error_count > 0 or warning_high_count > 0:
        mail_subject = review_data.get("mail_subject", f"{client}_{course} 교안 확인 요청드립니다")
        mail_body = review_data.get("mail_body", "")
        mail_section = f"""  <div class="mail-section">
    <h3>강사 수정 요청 메일 초안</h3>
    <div class="mail-title">
      <span id="mailTitle">{escape_html(mail_subject)}</span>
      <button class="copy-btn-sm" onclick="copyTitle()">제목 복사</button>
    </div>
    <div class="mail-content mail-body" id="mailBody">{escape_html(mail_body)}</div>
    <div class="btn-row">
      <button class="copy-btn" onclick="copyAll()">전체 복사 (제목+본문)</button>
      <button class="copy-btn-sm" onclick="copyBody()">본문만 복사</button>
    </div>
  </div>"""
    else:
        mail_section = ""

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
  /* 가이드 버튼 (현재 비활성 — 원복 시 주석 해제)
  .guide-btn {{ position: absolute; top: 18px; right: 24px; background: linear-gradient(180deg, #12c4f5, #01a1d0); color: #fff; border: none; border-top: 1px solid rgba(255,255,255,0.3); padding: 5px 14px; border-radius: 6px; font-size: 13px; font-weight: 600; text-decoration: none; cursor: pointer; transition: all 0.2s; display: inline-flex; align-items: center; gap: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.2), inset 0 1px 0 rgba(255,255,255,0.2); font-family: "Pretendard", -apple-system, sans-serif; }}
  .guide-btn:hover {{ background: linear-gradient(180deg, #0fb8e8, #0093be); transform: translateY(-1px); box-shadow: 0 3px 6px rgba(0,0,0,0.25), inset 0 1px 0 rgba(255,255,255,0.2); }}
  */
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
  .section-title {{ font-size: 16px; font-weight: bold; margin: 28px 0 12px; padding-left: 10px; }}
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
  .mail-title {{ background: #fff; border: 1px solid #e0e0e0; border-bottom: none; border-radius: 6px 6px 0 0; padding: 12px 16px; font-size: 13px; font-weight: 600; display: flex; justify-content: space-between; align-items: center; }}
  .mail-body {{ border-radius: 0 0 6px 6px; }}
  .btn-row {{ display: flex; gap: 8px; margin-top: 10px; }}
  .copy-btn {{ padding: 8px 20px; background: #13299f; color: #fff; border: none; border-radius: 6px; cursor: pointer; font-size: 13px; }}
  .copy-btn:hover {{ background: #0e1f7a; }}
  .copy-btn-sm {{ padding: 4px 12px; background: #e0e0e0; color: #333; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; }}
  .copy-btn-sm:hover {{ background: #bdbdbd; }}
  .copied {{ background: #4caf50 !important; color: #fff !important; }}
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
  <!-- 가이드 버튼 (현재 비활성 — 원복 시 주석 해제)
  <a class="guide-btn" href="https://www.notion.so/Guide-32711db511ea812bbb8dd4d827d38d12" target="_blank"><span style="font-size:20px;">&#128221;</span> 사용 가이드</a>
  -->
  <h1>「{course}」<span class="divider">|</span><span style="color: #01b1e3;">교안 검수 결과</span></h1>
  <p class="subtitle">{client} &middot; {instructor} 강사 &middot; {review_date} 검수</p>

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
  </div>
</div>

<script>
function copyTitle() {{
  var text = document.getElementById('mailTitle').innerText;
  navigator.clipboard.writeText(text).then(function() {{
    var btn = event.target;
    btn.textContent = '복사!';
    btn.classList.add('copied');
    setTimeout(function() {{ btn.textContent = '제목 복사'; btn.classList.remove('copied'); }}, 1500);
  }});
}}
function copyBody() {{
  var text = document.getElementById('mailBody').innerText;
  navigator.clipboard.writeText(text).then(function() {{
    var btn = event.target;
    btn.textContent = '복사 완료!';
    btn.classList.add('copied');
    setTimeout(function() {{ btn.textContent = '본문만 복사'; btn.classList.remove('copied'); }}, 1500);
  }});
}}
function copyAll() {{
  var title = document.getElementById('mailTitle').innerText;
  var body = document.getElementById('mailBody').innerText;
  navigator.clipboard.writeText(title + '\\n\\n' + body).then(function() {{
    var btn = event.target;
    btn.textContent = '복사 완료!';
    btn.classList.add('copied');
    setTimeout(function() {{ btn.textContent = '전체 복사 (제목+본문)'; btn.classList.remove('copied'); }}, 1500);
  }});
}}
</script>
</body>
</html>"""

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    return output_path


def open_in_browser(html_path):
    """생성된 HTML을 브라우저에서 열기"""
    webbrowser.open(f"file:///{html_path.replace(os.sep, '/')}")


if __name__ == "__main__":
    # 어제 샘플과 동일한 테스트 데이터
    sample_data = {
        "file_name": "ppt 스킬업 한번에 끝내는 제작 스킬 3차수 강의교안_오류테스트.pptx",
        "total_pages": 224,
        "client": "HL그룹",
        "course": "한번에 끝내는 파워포인트 디자인 제작 스킬",
        "instructor": "이지훈",
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
            {"page": "p.24", "type": "유사 단어",
             "original": '보안 <span class="highlight">강력</span>',
             "suggestion": '같은 페이지 "서비스강화", "생산성향상" 패턴 → "강화"가 자연스러움',
             "location": "[중]"},
            {"page": "p.183→184", "type": "문맥 모순",
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
        "mail_body": """이지훈 강사님, 안녕하세요.
HL그룹 한번에 끝내는 파워포인트 디자인 제작 스킬 교안 준비에 힘써 주셔서 감사합니다.

교안 검토 중 아래 사항이 확인되어 말씀드립니다.
확인 후 수정이 필요한 부분이 있으시면 반영 부탁드리겠습니다.

1. p.75 "최첨단 물류 시서템" → "시스템"
2. p.75 "디지털 전환 인프로" → "인프라"
3. p.80 "의류용 섬요로" → "섬유로"
4. p.84 "형태안정성 및 가경" → "가격"
5. p.87 "투자자와의 동반성자" → "동반성장"

※ 일부 항목은 의도하신 표현일 수 있어 확인 차 여쭙습니다.
수정 반영 후 파일을 회신해 주시면 최종 확인 후 진행하겠습니다.
바쁘신 중에 번거로우시겠지만 확인 부탁드리겠습니다.

감사합니다.
OOO 드림"""
    }

    path = generate_html(sample_data)
    print(f"HTML 생성 완료: {path}")
    open_in_browser(path)
