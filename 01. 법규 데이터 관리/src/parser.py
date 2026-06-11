"""
XML 파싱 모듈.

thdCmp API 응답 실제 구조:
  <LspttnThdCmpLawXService>
    <위임조문삼단비교>
      <법률조문>
        <조번호>0003</조번호>          ← 4자리 제로패딩
        <조가지번호>00</조가지번호>    ← '00'=없음, '02'=의2
        <조제목>제3조(적용 범위)</조제목>
        <시행령조문>                   ← 있을 수도 없을 수도
          <법령명>산업안전보건법 시행령</법령명>
          <조번호>0002</조번호>
          <조가지번호>00</조가지번호>
          <조제목>제2조(적용범위 등)</조제목>
        </시행령조문>
        <시행규칙조문>                 ← 있을 수도 없을 수도
          ...
        </시행규칙조문>
      </법률조문>
      ...
    </위임조문삼단비교>
"""
import re
import xml.etree.ElementTree as ET
from typing import Optional
import pandas as pd

_CONTENT_TAGS = {"조문내용", "항내용", "호내용", "목내용", "단서내용"}

COLUMNS_3STAGE = [
    "법률 조문번호", "법률 조문제목",
    "시행령 조문번호", "시행령 조문제목",
    "시행규칙 조문번호", "시행규칙 조문제목",
]

COLUMNS_OTHER = [
    "법률 조문번호", "법률 조문제목",
    "위임규칙 법령명", "위임규칙 조문번호", "위임규칙 조문제목",
]

# "제3조(적용 범위)" → ("제3조", "적용 범위")
_TITLE_RE = re.compile(r'^(제\d+조(?:의\d+)?)\((.+)\)\s*$')


# ── 내부 유틸 ─────────────────────────────────────────────────────────────────

def _text(elem: Optional[ET.Element]) -> str:
    if elem is None or elem.text is None:
        return ""
    return elem.text.strip()


def _parse_xml(xml_text: str) -> ET.Element:
    xml_text = xml_text.lstrip("﻿")  # UTF-8 BOM 제거
    try:
        return ET.fromstring(xml_text)
    except ET.ParseError:
        return ET.fromstring(xml_text.encode("utf-8"))


def _build_num(jo_num: str, jo_ga: str) -> str:
    """
    '0003' + '00' → '제3조'
    '0004' + '02' → '제4조의2'
    """
    try:
        n = int(jo_num)
    except (ValueError, TypeError):
        return ""
    ga = int(jo_ga) if jo_ga and jo_ga != "00" else 0
    return f"제{n}조의{ga}" if ga else f"제{n}조"


def _split_title(jo_title: str) -> tuple[str, str]:
    """
    '제3조(적용 범위)' → ('제3조', '적용 범위')
    '제3조'            → ('제3조', '')
    ''                 → ('', '')
    """
    if not jo_title:
        return "", ""
    m = _TITLE_RE.match(jo_title)
    if m:
        return m.group(1), m.group(2)
    return jo_title, ""


def to_raw(article_num: str) -> str:
    """
    갭 분석용: 조문번호를 raw 숫자 형식으로 변환.
    '제3조'   → '3'
    '제4조의2' → '4의2'
    """
    m = re.match(r'^제(\d+(?:의\d+)?)조$', article_num)
    return m.group(1) if m else article_num


_to_raw = to_raw  # 내부 호환성 유지


def _collect_content(unit: ET.Element) -> str:
    return " ".join(
        elem.text.strip()
        for elem in unit.iter()
        if elem.tag in _CONTENT_TAGS and elem.text
    )


# ── 공개 파싱 함수 ─────────────────────────────────────────────────────────────

def parse_law_search(xml_text: str) -> list[dict]:
    """
    법령 검색 결과 XML 파싱.
    반환: [{"name": 법령명, "lst": 법령일련번호, "id": 법령ID}, ...]
    """
    root = _parse_xml(xml_text)
    laws: list[dict] = []
    for item in root.iter("law"):
        name = _text(_find(item, "법령명한글", "법령명_한글", "법령명"))
        lst  = _text(_find(item, "법령일련번호", "MST", "LST", "일련번호"))
        lid  = _text(_find(item, "법령ID", "ID"))
        if name:
            laws.append({"name": name, "lst": lst, "id": lid})
    return laws


def parse_thdcmp_search(xml_text: str) -> list[dict]:
    """
    3단비교 법령 목록 검색 결과 파싱 (target=thdCmp).
    법률·시행령·시행규칙이 공식 연결된 세트로 반환된다.
    반환: [{"name": 법령명, "lst": 법령일련번호, "type": 법령구분명}, ...]
    """
    root = _parse_xml(xml_text)
    laws: list[dict] = []
    for item in root.iter("thdCmp"):
        name = _text(_find(item, "법령명한글"))
        lst  = _text(_find(item, "삼단비교일련번호"))   # thdCmp 검색 전용 태그명
        kind = _text(_find(item, "법령구분명"))
        if name and lst:
            laws.append({"name": name, "lst": lst, "type": kind})
    return laws


def parse_3stage_comparison(
    xml_text: str,
    debug_path: Optional[str] = None,
    rul_name: Optional[str] = None,
) -> tuple:
    """
    thdCmp API 응답에서 위임조문 3단비교표 파싱.

    rul_name 이 주어지면 시행규칙조문 중 해당 법령명과 일치하는 것만
    시행규칙 열에 포함하고, 나머지는 other_df(기타 위임규칙)로 분리한다.

    반환:
        (main_df,
         other_df        : DataFrame,  # 기타 위임규칙 (rul_name 불일치 행)
         mapped_enf_nums : set[str],
         mapped_rul_nums : set[str],
         enf_mst         : None,
         rul_mst         : None,
         law_name        : str)
    """
    if debug_path:
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(xml_text)
        print(f"  원본 XML 저장: {debug_path}")

    root = _parse_xml(xml_text)

    # 법령명 추출
    law_name = _text(_find(root, "법령명", "법령명한글", "기준법령명"))

    container = root.find(".//위임조문삼단비교")
    if container is None:
        empty = pd.DataFrame(columns=COLUMNS_3STAGE)
        empty_other = pd.DataFrame(columns=COLUMNS_OTHER)
        return empty, empty_other, set(), set(), None, None, law_name

    rows: list[dict] = []
    other_rows: list[dict] = []
    mapped_enf_nums: set[str] = set()
    mapped_rul_nums: set[str] = set()

    for item in container.findall("법률조문"):
        jo_title = _text(item.find("조제목"))

        # 빈 조제목 = 장·절 구분 헤더 → 건너뜀
        if not jo_title:
            continue

        l_num, l_title = _split_title(jo_title)
        if not l_num:
            l_num = _build_num(_text(item.find("조번호")), _text(item.find("조가지번호")))

        # 시행령 조문 (없으면 빈 문자열)
        enf_elem = item.find("시행령조문")
        if enf_elem is not None:
            ej_title = _text(enf_elem.find("조제목"))
            e_num, e_title = _split_title(ej_title)
            if not e_num:
                e_num = _build_num(_text(enf_elem.find("조번호")), _text(enf_elem.find("조가지번호")))
        else:
            e_num, e_title = "", ""

        # 시행규칙 조문 — rul_name 과 일치하면 시행규칙 열, 아니면 기타 위임규칙
        r_num, r_title = "", ""
        rul_elem = item.find("시행규칙조문")
        if rul_elem is not None:
            rule_law_nm = _text(rul_elem.find("법령명"))
            rj_title    = _text(rul_elem.find("조제목"))
            o_num, o_title = _split_title(rj_title)
            if not o_num:
                o_num = _build_num(_text(rul_elem.find("조번호")), _text(rul_elem.find("조가지번호")))

            if rul_name is None or rule_law_nm == rul_name:
                r_num, r_title = o_num, o_title
            elif o_num:
                other_rows.append({
                    "법률 조문번호":    l_num,
                    "법률 조문제목":    l_title,
                    "위임규칙 법령명":  rule_law_nm,
                    "위임규칙 조문번호": o_num,
                    "위임규칙 조문제목": o_title,
                })

        rows.append({
            "법률 조문번호":     l_num,
            "법률 조문제목":     l_title,
            "시행령 조문번호":   e_num,
            "시행령 조문제목":   e_title,
            "시행규칙 조문번호": r_num,
            "시행규칙 조문제목": r_title,
        })

        if e_num:
            mapped_enf_nums.add(_to_raw(e_num))
        if r_num:
            mapped_rul_nums.add(_to_raw(r_num))

    df = pd.DataFrame(rows, columns=COLUMNS_3STAGE) if rows else pd.DataFrame(columns=COLUMNS_3STAGE)
    df = df.drop_duplicates(subset=["법률 조문번호", "시행령 조문번호", "시행규칙 조문번호"])

    other_df = pd.DataFrame(other_rows, columns=COLUMNS_OTHER) if other_rows else pd.DataFrame(columns=COLUMNS_OTHER)
    return df, other_df, mapped_enf_nums, mapped_rul_nums, None, None, law_name


def parse_law_articles(xml_text: str, include_content: bool = False) -> list[dict]:
    """
    법령 전체 조문 파싱 (갭 분석용).
    '조문여부'가 '조문'인 항목만 포함 (장·절 구분 제외).
    반환: [{"번호": str, "제목": str [, "내용": str]}, ...]
    """
    root = _parse_xml(xml_text)
    articles: list[dict] = []

    for unit in root.iter("조문단위"):
        kind = _text(unit.find("조문여부"))
        if kind and kind != "조문":
            continue
        jo_num = unit.find("조번호")
        jo_ga  = unit.find("조가지번호")
        if jo_num is not None and _text(jo_num):
            num = _build_num(_text(jo_num), _text(jo_ga))
        else:
            num = _text(unit.find("조문번호"))
        title = _text(unit.find("조문제목"))
        if not num:
            continue
        art: dict = {"번호": num, "제목": title}
        if include_content:
            art["내용"] = _collect_content(unit)
        articles.append(art)

    return articles


# ── 내부 헬퍼 ─────────────────────────────────────────────────────────────────

def _find(root: ET.Element, *tags: str) -> Optional[ET.Element]:
    for tag in tags:
        elem = root.find(f".//{tag}")
        if elem is not None:
            return elem
    return None
