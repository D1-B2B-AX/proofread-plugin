"""
교안 검수 공통 감지 유틸리티
- 텍스트 깨짐 감지 (broken_texts) — □ 대체 문자, � 유니코드 대체, 자모 분리
- 반복 텍스트 변형 감지 (suspected_variants)
- 형태소 분석 기반 의심 단어 감지 (suspicious_words)
- 영문 스펠 체크 기반 의심 단어 감지 (suspicious_english)

extract_pptx.py, extract_pdf.py에서 공통으로 사용
"""

import re
from collections import defaultdict

# 형태소 분석기 (kiwipiepy)
try:
    from kiwipiepy import Kiwi
    _kiwi = Kiwi()
    _HAS_KIWI = True
except ImportError:
    _HAS_KIWI = False

# 영문 스펠체커 (pyspellchecker)
try:
    from spellchecker import SpellChecker
    _spell = SpellChecker()
    _HAS_SPELL = True
except ImportError:
    _HAS_SPELL = False


def detect_broken_text(pages_or_slides, key="slide_number"):
    """
    텍스트 깨짐을 감지 (폰트 임베딩 실패, 인코딩 깨짐)

    감지 유형:
    1. □ 대체 문자: 폰트가 없을 때 □로 치환되는 패턴
    2. � 유니코드 대체 문자(U+FFFD): 인코딩 변환 실패 시 발생
    3. 한글 자모 분리: 완성형 글자가 ㄱㅏㅁ처럼 풀어지는 패턴
    """
    # 패턴 1: □(U+25A1), ■(U+25A0), ⊠(U+22A0) 등 대체 문자가 2개 이상 연속
    replacement_box_pattern = re.compile(r'[□■⊠☐▢]{2,}')

    # 패턴 2: �(U+FFFD) 유니코드 대체 문자
    unicode_replacement_pattern = re.compile(r'\uFFFD')

    # 패턴 3: 한글 자모(ㄱ-ㅎ: U+3131-U+314E, ㅏ-ㅣ: U+314F-U+3163)가 3개 이상 연속
    # 정상 텍스트에서 자모가 연속으로 3개 이상 나오는 경우는 거의 없음
    jamo_pattern = re.compile(r'[\u3131-\u3163]{3,}')

    broken_items = []

    for page in pages_or_slides:
        page_num = page[key]
        for item in page["texts"]:
            text = item["text"].strip()
            position = item.get("position", "중")

            # 패턴 1: □ 대체 문자
            matches = replacement_box_pattern.findall(text)
            if matches:
                broken_items.append({
                    "page": page_num,
                    "position": position,
                    "text": text,
                    "type": "replacement_char",
                    "detail": f"대체 문자(□) {sum(len(m) for m in matches)}개 감지",
                    "reason": "폰트 임베딩 실패로 글자가 □로 표시됨 — 해당 폰트가 없는 환경에서 변환된 것으로 보임"
                })

            # 패턴 2: � 유니코드 대체 문자
            matches = unicode_replacement_pattern.findall(text)
            if matches:
                broken_items.append({
                    "page": page_num,
                    "position": position,
                    "text": text,
                    "type": "unicode_replacement",
                    "detail": f"유니코드 대체 문자(�) {len(matches)}개 감지",
                    "reason": "인코딩 변환 실패로 글자가 �로 표시됨 — 폰트 누락 또는 인코딩 불일치"
                })

            # 패턴 3: 한글 자모 분리
            matches = jamo_pattern.findall(text)
            if matches:
                broken_items.append({
                    "page": page_num,
                    "position": position,
                    "text": text,
                    "type": "jamo_separated",
                    "detail": f"자모 분리 감지: {'、'.join(matches)}",
                    "reason": "한글이 자모(초성/중성/종성)로 분리되어 표시됨 — 폰트 호환 문제 또는 유니코드 NFD 정규화 문제"
                })

    return broken_items


def _edit_distance_one(s1, s2):
    """두 문자열이 정확히 1글자만 다른지 확인 (길이 동일한 경우만)
    숫자↔숫자 차이는 의도적 데이터 차이일 가능성이 높으므로 제외
    구두점↔구두점 차이도 의도적 표기 차이일 가능성이 높으므로 제외"""
    if len(s1) != len(s2):
        return False
    PUNCTUATION = set('.,;:!?·…~-–—\'\"()[]{}/<>@#$%^&*+=|\\')
    diff_count = 0
    for i, (c1, c2) in enumerate(zip(s1, s2)):
        if c1 != c2:
            diff_count += 1
            if diff_count > 1:
                return False
            if c1.isdigit() and c2.isdigit():
                return False
            if c1 in PUNCTUATION and c2 in PUNCTUATION:
                return False
    return diff_count == 1


def detect_text_variants(pages_or_slides, key="slide_number"):
    """
    전체 슬라이드/페이지에서 반복 텍스트의 소수 변형을 감지

    원리: 같은 텍스트가 여러 번 등장하는데 1글자만 다른 변형이 있으면,
    다수파를 정본으로 보고 소수파를 의심 변형으로 플래그.
    빈도가 동일하면 양쪽 모두 OM 확인 필요로 플래그.
    """
    # 1단계: 모든 텍스트를 수집 (5자 이상만 대상)
    text_locations = defaultdict(list)
    for page in pages_or_slides:
        page_num = page[key]
        for item in page["texts"]:
            text = item["text"].strip()
            if len(text) >= 5:
                text_locations[text].append({
                    "page": page_num,
                    "position": item.get("position", "중")
                })

    # 2단계: 유사 텍스트 그룹 찾기 (1글자 차이)
    texts = list(text_locations.keys())
    variants = []
    checked = set()

    for i, t1 in enumerate(texts):
        if t1 in checked:
            continue
        for j, t2 in enumerate(texts):
            if i >= j or t2 in checked:
                continue
            if abs(len(t1) - len(t2)) > 0:
                continue
            if _edit_distance_one(t1, t2):
                count1 = len(text_locations[t1])
                count2 = len(text_locations[t2])

                # 차이나는 글자 위치 찾기
                diff_idx = -1
                for k, (c1, c2) in enumerate(zip(t1, t2)):
                    if c1 != c2:
                        diff_idx = k
                        break

                if count1 > count2:
                    # 다수파 = 정본, 소수파 = 변형 의심
                    variants.append({
                        "majority_text": t1,
                        "majority_count": count1,
                        "variant_text": t2,
                        "variant_count": count2,
                        "variant_locations": text_locations[t2],
                        "diff_char": f"'{t2[diff_idx]}'→'{t1[diff_idx]}' (위치 {diff_idx + 1}번째 글자)"
                    })
                    checked.add(t2)
                elif count2 > count1:
                    variants.append({
                        "majority_text": t2,
                        "majority_count": count2,
                        "variant_text": t1,
                        "variant_count": count1,
                        "variant_locations": text_locations[t1],
                        "diff_char": f"'{t1[diff_idx]}'→'{t2[diff_idx]}' (위치 {diff_idx + 1}번째 글자)"
                    })
                    checked.add(t1)
                else:
                    # 빈도 동일 → 양쪽 모두 플래그 (어느 쪽이 오류인지 OM 확인 필요)
                    variants.append({
                        "majority_text": "(빈도 동일 — 정본 판단 불가)",
                        "majority_count": count1,
                        "variant_text": t1,
                        "variant_count": count1,
                        "variant_locations": text_locations[t1],
                        "diff_char": f"'{t1[diff_idx]}' vs '{t2[diff_idx]}' (위치 {diff_idx + 1}번째 글자)",
                        "counterpart": t2
                    })
                    variants.append({
                        "majority_text": "(빈도 동일 — 정본 판단 불가)",
                        "majority_count": count2,
                        "variant_text": t2,
                        "variant_count": count2,
                        "variant_locations": text_locations[t2],
                        "diff_char": f"'{t2[diff_idx]}' vs '{t1[diff_idx]}' (위치 {diff_idx + 1}번째 글자)",
                        "counterpart": t1
                    })
                    checked.add(t1)
                    checked.add(t2)

    return variants


def detect_suspicious_words(pages_or_slides, key="slide_number"):
    """
    형태소 분석기로 사전에 없는 한글 단어를 감지

    원리: 텍스트를 공백/구두점으로 분리하여 개별 어절을 추출한 뒤,
    순수 한글 어절을 형태소 분석했을 때 조사 없이 낱글자로 쪼개지면
    사전에 없는 단어일 가능성이 높음
    """
    if not _HAS_KIWI:
        return []

    word_split_pattern = re.compile(r'[가-힣]+')
    PARTICLE_TAGS = {'JKS', 'JKC', 'JKG', 'JKO', 'JKB', 'JKV', 'JKQ',
                     'JX', 'JC', 'EC', 'EF', 'ETM', 'ETN', 'EP',
                     'XSV', 'XSA', 'XSN', 'XPN', 'VCP', 'VCN'}

    word_locations = defaultdict(list)
    for page in pages_or_slides:
        page_num = page[key]
        for item in page["texts"]:
            text = item["text"].strip()
            for token_raw in text.split():
                matches = word_split_pattern.findall(token_raw)
                for word in matches:
                    if 2 <= len(word) <= 7:
                        word_locations[word].append({
                            "page": page_num,
                            "position": item.get("position", "중"),
                            "context": text
                        })

    suspicious = []
    checked_words = set()

    for word, locations in word_locations.items():
        if word in checked_words:
            continue

        tokens = _kiwi.tokenize(word)

        if len(tokens) == 1:
            continue

        has_particle = any(t.tag in PARTICLE_TAGS for t in tokens)
        if has_particle:
            last_token = tokens[-1]
            if last_token.tag == 'EF' and len(last_token.form) == 1:
                nng_count = sum(1 for t in tokens if t.tag in ('NNG', 'NNP'))
                if nng_count >= 2:
                    pass
                else:
                    continue
            else:
                continue

        tags = [t.tag for t in tokens]
        forms = [t.form for t in tokens]

        if all(t in ('NNG', 'NNP') for t in tags):
            continue

        if any(t == 'NR' for t in tags):
            continue

        is_suspicious = False

        if any(t in ('VV', 'VA', 'EF') for t in tags):
            is_suspicious = True

        if any(t in ('MM', 'MAG') for t in tags):
            has_single_char = any(len(f) == 1 for f in forms)
            if has_single_char and len(word) <= 3:
                is_suspicious = True

        if not is_suspicious and len(tokens) == 2:
            if tags == ['NNG', 'NNB'] and len(forms[0]) == 1 and len(word) <= 3:
                is_suspicious = True

        if not is_suspicious:
            continue

        token_info = [(t.form, t.tag) for t in tokens]
        suspicious.append({
            "word": word,
            "tokens": str(token_info),
            "occurrences": len(locations),
            "locations": locations[:5],
            "reason": f"형태소 분석 결과 '{word}'이(가) 비정상적 태그 조합으로 분리됨 → 사전에 없는 단어 의심"
        })
        checked_words.add(word)

    return suspicious


def detect_suspicious_english(pages_or_slides, key="slide_number"):
    """
    영문 스펠체커로 사전에 없는 영어 단어를 감지

    필터:
    - 숫자 포함 단어 제외 (VARCHAR2, H2O 등)
    - 대문자 약어 2~5글자 제외 (DBMS, SQL, API 등)
    - camelCase/PascalCase 제외 (DBeaver, GitHub 등)
    - URL, 이메일, 코드 패턴 제외
    - 3글자 미만 제외
    """
    if not _HAS_SPELL:
        return []

    eng_word_pattern = re.compile(r'[A-Za-z]{3,}')
    code_context_pattern = re.compile(r'[A-Za-z_]+\.[A-Za-z_]+|[A-Za-z]+_[A-Za-z]+')

    word_locations = defaultdict(list)
    for page in pages_or_slides:
        page_num = page[key]
        for item in page["texts"]:
            text = item["text"].strip()
            if code_context_pattern.search(text):
                continue
            words = eng_word_pattern.findall(text)
            for word in words:
                if any(c.isdigit() for c in word):
                    continue
                if word.isupper():
                    continue
                has_upper = any(c.isupper() for c in word)
                has_lower = any(c.islower() for c in word)
                if has_upper and has_lower and word[0].isupper() and any(c.isupper() for c in word[1:]):
                    continue
                if word.islower():
                    continue
                if word.lower() not in _spell:
                    candidates = _spell.candidates(word.lower())
                    candidate_str = ", ".join(list(candidates)[:3]) if candidates else "없음"
                    word_locations[word].append({
                        "page": page_num,
                        "position": item.get("position", "중"),
                        "context": text,
                        "candidates": candidate_str
                    })

    suspicious = []
    for word, locations in word_locations.items():
        suspicious.append({
            "word": word,
            "occurrences": len(locations),
            "locations": locations[:5],
            "candidates": locations[0]["candidates"],
            "reason": f"영문 스펠체커에 '{word}'이(가) 미등록 → 스펠링 오류 의심"
        })

    return suspicious
