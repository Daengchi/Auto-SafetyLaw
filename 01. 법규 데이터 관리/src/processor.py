"""
데이터 처리 모듈.

build_3stage_table():
    시행령/시행규칙 조문 본문에서 '법 제X조', '영 제X조' 참조 패턴을 추출하여
    실제 위임 관계에 기반한 3단비교표를 생성합니다.

find_missing_articles():
    상위 법령에 대응 조문이 없는 하위 법령 조문 리스트를 반환합니다.
"""
import re
from collections import defaultdict
import pandas as pd
from src.parser import to_raw

# 시행령 조문 본문에서 '법 제X조' 참조를 찾는 패턴
# 예: "법 제5조", "이 법 제5조", "같은 법 제5조"
_LAW_REF_RE = re.compile(r'(?:이\s+|같은\s+)?법\s+제(\d+)조')

# 시행규칙 조문 본문에서 '영 제X조' 참조를 찾는 패턴
# 예: "영 제3조", "이 영 제3조", "같은 영 제3조"
_ENF_REF_RE = re.compile(r'(?:이\s+|같은\s+)?영\s+제(\d+)조')

COLUMNS = [
    "법률 조문번호", "법률 조문제목",
    "시행령 조문번호", "시행령 조문제목",
    "시행규칙 조문번호", "시행규칙 조문제목",
]


def _fmt(num: str) -> str:
    """조문번호 표시 형식 통일: '5' → '제5조'"""
    if not num or num.startswith("제"):
        return num
    return f"제{num}조"


def _sort_key(num: str) -> tuple:
    """'5', '5의2', '12' 등을 올바른 순서로 정렬하기 위한 키."""
    parts = re.split(r"의", num)
    return tuple(int(p) if p.isdigit() else 0 for p in parts)


def _make_row(l_num, l_title, e_num, e_title, r_num, r_title) -> dict:
    return {
        "법률 조문번호":     _fmt(l_num),
        "법률 조문제목":     l_title,
        "시행령 조문번호":   _fmt(e_num) if e_num else "",
        "시행령 조문제목":   e_title,
        "시행규칙 조문번호": _fmt(r_num) if r_num else "",
        "시행규칙 조문제목": r_title,
    }


def build_3stage_table(
    law_arts: list[dict],
    enf_arts: list[dict],
    rul_arts: list[dict],
) -> pd.DataFrame:
    """
    위임 참조 텍스트 분석 기반 3단비교표 생성.

    알고리즘:
    1. 시행령 각 조문의 '내용'에서 '법 제X조' 패턴 추출
       → enf_to_law  : {시행령조문번호: set(참조하는 법률조문번호)}
    2. 시행규칙 각 조문의 '내용'에서 '영 제X조' 패턴 추출
       → rul_to_enf  : {시행규칙조문번호: set(참조하는 시행령조문번호)}
    3. 역방향 맵 구성
       → law_to_enf  : {법률조문번호: set(참조하는 시행령조문번호)}
       → enf_to_rul  : {시행령조문번호: set(참조하는 시행규칙조문번호)}
    4. 법률 조문 기준으로 행 생성

    법률/시행령/시행규칙 중 어느 쪽이 비어있어도 빈 열로 처리합니다.
    """
    # ── 조회 사전 ──────────────────────────────────────────────────────────────
    enf_dict = {a["번호"]: a for a in enf_arts}
    rul_dict = {a["번호"]: a for a in rul_arts}

    # ── 위임 참조 맵 구성 ──────────────────────────────────────────────────────
    # 시행령 조문 → 참조하는 법률 조문번호 집합
    enf_to_law: dict[str, set] = defaultdict(set)
    for e in enf_arts:
        for ref in _LAW_REF_RE.findall(e.get("내용", "")):
            enf_to_law[e["번호"]].add(ref)

    # 시행규칙 조문 → 참조하는 시행령 조문번호 집합
    rul_to_enf: dict[str, set] = defaultdict(set)
    for r in rul_arts:
        for ref in _ENF_REF_RE.findall(r.get("내용", "")):
            rul_to_enf[r["번호"]].add(ref)

    # 역방향: 법률 조문번호 → 이를 참조하는 시행령 조문번호 집합
    law_to_enf: dict[str, set] = defaultdict(set)
    for enf_num, law_nums in enf_to_law.items():
        for law_num in law_nums:
            law_to_enf[law_num].add(enf_num)

    # 역방향: 시행령 조문번호 → 이를 참조하는 시행규칙 조문번호 집합
    enf_to_rul: dict[str, set] = defaultdict(set)
    for rul_num, enf_nums in rul_to_enf.items():
        for enf_num in enf_nums:
            enf_to_rul[enf_num].add(rul_num)

    # ── 3단비교표 행 생성 ──────────────────────────────────────────────────────
    rows: list[dict] = []

    for l in law_arts:
        l_num = l["번호"]
        enf_nums = sorted(law_to_enf.get(l_num, set()), key=_sort_key)

        if not enf_nums:
            # 이 법률 조문을 위임받은 시행령 조문 없음
            rows.append(_make_row(l_num, l["제목"], "", "", "", ""))
        else:
            for e_num in enf_nums:
                e = enf_dict.get(e_num, {})
                rul_nums = sorted(enf_to_rul.get(e_num, set()), key=_sort_key)

                if not rul_nums:
                    rows.append(_make_row(
                        l_num, l["제목"],
                        e_num, e.get("제목", ""),
                        "", "",
                    ))
                else:
                    for r_num in rul_nums:
                        r = rul_dict.get(r_num, {})
                        rows.append(_make_row(
                            l_num, l["제목"],
                            e_num, e.get("제목", ""),
                            r_num, r.get("제목", ""),
                        ))

    return pd.DataFrame(rows, columns=COLUMNS) if rows else pd.DataFrame(columns=COLUMNS)


def find_missing_articles(
    all_articles: list[dict],
    mapped_nums: set[str],
) -> pd.DataFrame:
    """
    상위 법령 조문번호 집합에 없는 하위 법령 조문을 반환 (누락 조문).

    Args:
        all_articles: 하위 법령의 전체 조문 [{번호, 제목}, ...]
        mapped_nums:  상위 법령 조문번호 집합 (법률 or 시행령)

    Returns:
        DataFrame(columns=["조문번호", "조문제목"])
    """
    missing = [
        {"조문번호": _fmt(a["번호"]), "조문제목": a["제목"]}
        for a in all_articles
        if to_raw(a["번호"]) not in mapped_nums
    ]
    return pd.DataFrame(missing, columns=["조문번호", "조문제목"])
