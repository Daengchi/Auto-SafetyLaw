"""
법규준수평가표 자동 생성기 - CLI 진입점

사용 예시:
  python main.py                          # laws.json 전체 법령 처리
  python main.py --name 산업안전보건법     # 단일 법령만 처리
  python main.py --output ./output
  python main.py --debug
  python main.py --update                 # 기존 파일 업데이트 (파일 선택창)
"""
import argparse
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv

from src.api_client import APIError, LawAPIClient
from src import exporter, parser, processor


# ── 헬퍼 ──────────────────────────────────────────────────────────────────────

def _find_queried_law(all_laws: list[dict], query: str) -> dict | None:
    """검색 결과에서 입력 법령명과 가장 일치하는 항목 반환."""
    for law in all_laws:
        if law["name"] == query:
            return law
    for law in all_laws:
        if query in law["name"]:
            return law
    return all_laws[0] if all_laws else None


def _choose_law(laws: list[dict]) -> dict:
    """기본 법률 후보가 여러 개일 때 사용자에게 선택 요청."""
    print(f"\n  {len(laws)}개의 법령이 검색되었습니다:")
    for i, law in enumerate(laws, 1):
        print(f"    {i}. {law['name']}  (MST: {law['lst']})")
    print("\n  ※ 기본 법률(시행령/시행규칙이 아닌 것)을 선택하세요.")
    while True:
        try:
            choice = int(input("  선택 (번호): ")) - 1
            if 0 <= choice < len(laws):
                return laws[choice]
            print(f"  1~{len(laws)} 사이의 번호를 입력하세요.")
        except (ValueError, KeyboardInterrupt):
            print("\n취소되었습니다.")
            sys.exit(0)


def _get_articles(
    client: LawAPIClient,
    mst: str,
    label: str,
    debug: bool,
    include_content: bool = False,
) -> list[dict]:
    """법령 MST로 전체 조문 조회·파싱."""
    print(f"  {label} 전체 조문 조회 중...  (MST: {mst})")
    try:
        xml = client.get_law_articles(mst)
        if debug:
            fname = f"debug_{label}_{mst}.xml"
            with open(fname, "w", encoding="utf-8") as f:
                f.write(xml)
            print(f"    XML 저장: {fname}")
        arts = parser.parse_law_articles(xml, include_content=include_content)
        print(f"    {len(arts)}개 조문 추출")
        return arts
    except APIError as e:
        print(f"  ⚠  {label} 조회 실패: {e}")
        return []


# ── 행정규칙 처리 ──────────────────────────────────────────────────────────────

def _process_admrul(
    client: LawAPIClient,
    law_name: str,
) -> tuple[str, object] | None:
    """
    행정규칙(고시·훈령·예규)을 처리해 (행정규칙명, admrul_df)를 반환.
    검색되지 않으면 None.
    """
    print(f"  '{law_name}' 행정규칙으로 재검색 중...")
    try:
        search_xml = client.search_admrul(law_name)
    except APIError as e:
        print(f"  오류: {e}")
        return None

    rules = parser.parse_admrul_search(search_xml)
    if not rules:
        return None

    rule = _find_queried_law(rules, law_name) or rules[0]
    print(f"  행정규칙: {rule['name']}  (종류: {rule['종류']}, 일련번호: {rule['lst']})")

    try:
        articles_xml = client.get_admrul_articles(rule["lst"])
    except APIError as e:
        print(f"  오류: {e}")
        return None

    articles = parser.parse_admrul_articles(articles_xml)
    print(f"  {len(articles)}개 조문 추출")
    if not articles:
        return None

    combined_df = exporter._build_admrul_df(articles)
    return (rule["name"], combined_df)


# ── 단일 법령 처리 ─────────────────────────────────────────────────────────────

def _process_one_law(
    client: LawAPIClient,
    law_name: str,
    debug: bool = False,
    auto_select: bool = False,
) -> tuple[str, object] | None:
    """
    단일 법령을 처리해 (실제법령명, combined_df)를 반환.
    실패 시 None 반환.
    단독 부령의 경우 시행규칙 C열에만 데이터를 채운 combined_df 반환.
    """
    empty_gap = __import__("pandas").DataFrame(columns=["조문번호", "조문제목"])

    # ── Step 1: thdCmp 검색으로 법령 목록 조회 + 유형 판별 ───────────────────
    print(f"\n  [1/3] '{law_name}' 3단비교 법령 목록 검색 중...")
    try:
        search_xml = client.search_3stage_laws(law_name)
    except APIError as e:
        print(f"  오류: {e}")
        return None

    all_laws = parser.parse_thdcmp_search(search_xml)
    if not all_laws:
        # 법령에서 못 찾으면 행정규칙(고시·훈령·예규)으로 폴백
        admrul = _process_admrul(client, law_name)
        if admrul is not None:
            return admrul
        print(f"  '{law_name}' 법령을 찾을 수 없습니다.")
        return None

    queried_law = _find_queried_law(all_laws, law_name)
    is_buryeong = queried_law is not None and "부령" in queried_law.get("type", "")

    if is_buryeong:
        print(f"  '{law_name}' ({queried_law['type']}) 감지 - 상위 법률 탐색 중...")
        try:
            probe_xml = client.get_3stage_comparison(queried_law["lst"])
        except APIError as e:
            print(f"  오류: {e}")
            return None

        _, _, _, _, _, _, probe_law_name = parser.parse_3stage_comparison(probe_xml)

        if probe_law_name and probe_law_name != law_name:
            # 상위 법률 기준으로 재검색
            print(f"  -> 상위 법률 '{probe_law_name}' 기준으로 3단비교를 생성합니다.")
            try:
                search_xml = client.search_3stage_laws(probe_law_name)
            except APIError as e:
                print(f"  오류: {e}")
                return None
            all_laws = parser.parse_thdcmp_search(search_xml)
            if not all_laws:
                print(f"  '{probe_law_name}' 법령을 찾을 수 없습니다.")
                return None
            law_name = probe_law_name
        else:
            # 단독 부령 → 조문 목록을 시행규칙 C열에만 채운 combined_df 반환
            print(f"  단독 부령입니다. 시행규칙 조문으로 처리합니다.")
            print(f"\n  [2/3] 전체 조문 조회 중...  (MST: {queried_law['lst']})")
            try:
                articles_xml = client.get_law_articles(queried_law["lst"])
            except APIError as e:
                print(f"  오류: {e}")
                return None
            articles = parser.parse_law_articles(articles_xml)
            print(f"  {len(articles)}개 조문 추출")

            import pandas as pd
            combined_df = pd.DataFrame(
                [{"No.": i + 1, "법률 조문": "", "시행령 조문": "",
                  "시행규칙 조문": f"{a['번호']}\n{a['제목']}" if a['제목'] else a['번호'],
                  "해당여부": ""}
                 for i, a in enumerate(articles)],
                columns=["No.", "법률 조문", "시행령 조문", "시행규칙 조문", "해당여부"],
            )
            return (queried_law["name"], combined_df)

    # 법률 / 시행령 / 시행규칙 분류
    base_candidates = [l for l in all_laws if "시행령" not in l["name"] and "시행규칙" not in l["name"]]
    enf_candidates  = [l for l in all_laws if "시행령" in l["name"]]
    rul_candidates  = [l for l in all_laws if "시행규칙" in l["name"]]

    if not base_candidates:
        base_candidates = all_laws

    if len(base_candidates) == 1:
        selected = base_candidates[0]
        print(f"  법률:    {selected['name']}  (MST: {selected['lst']})")
    elif auto_select:
        selected = _find_queried_law(base_candidates, law_name) or base_candidates[0]
        print(f"  법률:    {selected['name']}  (MST: {selected['lst']})  [자동 선택]")
    else:
        selected = _choose_law(base_candidates)

    law_name = selected["name"]
    law_mst  = selected["lst"]
    enf_mst  = enf_candidates[0]["lst"]  if enf_candidates else None
    rul_mst  = rul_candidates[0]["lst"]  if rul_candidates else None
    rul_name = rul_candidates[0]["name"] if rul_candidates else None

    if enf_mst:
        print(f"  시행령:  {enf_candidates[0]['name']}  (MST: {enf_mst})")
    if rul_mst:
        print(f"  시행규칙: {rul_candidates[0]['name']}  (MST: {rul_mst})")

    # ── Step 2: 위임조문 3단비교표 ────────────────────────────────────────────
    print(f"\n  [2/3] 위임조문 3단비교표 조회 중... (MST: {law_mst})")
    try:
        comparison_xml = client.get_3stage_comparison(law_mst)
    except APIError as e:
        print(f"  오류: {e}")
        return None

    debug_path = f"debug_thdCmp_{law_mst}.xml" if debug else None
    (main_df, _, mapped_enf_nums, mapped_rul_nums,
     _, _, law_name_api) = parser.parse_3stage_comparison(
        comparison_xml, debug_path=debug_path, rul_name=rul_name,
    )

    if law_name_api:
        law_name = law_name_api

    if main_df.empty:
        print("  thdCmp 파싱 결과가 비어있습니다.")
        return None

    print(f"  3단비교표: {len(main_df)}행")

    # ── Step 3: 갭 분석용 시행령·시행규칙 전체 조문 조회 ─────────────────────
    print(f"\n  [3/3] 갭 분석용 조문 조회 중...")
    missing_enf_df = None
    missing_rul_df = None

    if enf_mst:
        enf_arts = _get_articles(client, enf_mst, "시행령", debug)
        if enf_arts:
            missing_enf_df = processor.find_missing_articles(enf_arts, mapped_enf_nums)
            print(f"  누락된 시행령: {len(missing_enf_df)}개")
    else:
        print("  시행령 없음 - 갭 분석 건너뜀")

    if rul_mst:
        rul_arts = _get_articles(client, rul_mst, "시행규칙", debug)
        if rul_arts:
            missing_rul_df = processor.find_missing_articles(rul_arts, mapped_rul_nums)
            print(f"  누락된 시행규칙: {len(missing_rul_df)}개")
    else:
        print("  시행규칙 없음 - 갭 분석 건너뜀")

    enf_df  = missing_enf_df if missing_enf_df is not None else empty_gap
    rule_df = missing_rul_df if missing_rul_df is not None else empty_gap
    combined_df = exporter._build_combined_df(main_df, enf_df, rule_df)
    return (law_name, combined_df)


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main() -> None:
    load_dotenv()

    arg_parser = argparse.ArgumentParser(
        description="국가법령정보센터 법규준수평가표 생성기",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    arg_parser.add_argument("--name", "-n", metavar="법령명", default=None,
                            help="단일 법령만 처리 (미지정 시 laws.json 전체)")
    arg_parser.add_argument("--output", "-o", metavar="경로", default=".",
                            help="출력 디렉토리 (기본: 현재 디렉토리)")
    arg_parser.add_argument("--debug", action="store_true",
                            help="원본 XML 저장 (파싱 문제 진단용)")
    arg_parser.add_argument("--update", action="store_true",
                            help="기존 법규준수평가표 파일 업데이트 (파일 선택창)")
    args = arg_parser.parse_args()

    api_key = os.getenv("LAW_API_KEY")
    if not api_key:
        print("오류: LAW_API_KEY 환경변수가 없습니다.")
        print("  .env 파일에 다음을 추가하세요:  LAW_API_KEY=인증키")
        sys.exit(1)

    client = LawAPIClient(api_key)

    # ── 업데이트 모드 ───────────────────────────────────────────────────────────
    if args.update:
        import tkinter as tk
        from tkinter import filedialog
        from src.updater import update_file

        root = tk.Tk()
        root.withdraw()
        xlsx_path = filedialog.askopenfilename(
            title="업데이트할 법규준수평가표를 선택하세요",
            filetypes=[("Excel 파일", "*.xlsm *.xlsx"), ("모든 파일", "*.*")],
        )
        root.destroy()

        if not xlsx_path:
            print("파일을 선택하지 않았습니다.")
            sys.exit(0)

        def _process(law_name: str):
            return _process_one_law(client, law_name, debug=args.debug, auto_select=True)

        update_file(xlsx_path, _process)
        return

    # 처리할 법령 목록 결정
    if args.name:
        law_names = [args.name]
    else:
        laws_json = os.path.join(os.path.dirname(__file__), "..", "laws.json")
        if not os.path.exists(laws_json):
            print(f"오류: laws.json 파일을 찾을 수 없습니다: {laws_json}")
            sys.exit(1)
        with open(laws_json, encoding="utf-8") as f:
            law_names = json.load(f)
        print(f"laws.json에서 {len(law_names)}개 법령 로드")

    # 법령별 처리
    results: list[tuple[str, object]] = []
    for law_name in law_names:
        print(f"\n{'='*60}")
        print(f"처리: {law_name}")
        result = _process_one_law(client, law_name, debug=args.debug, auto_select=args.name is None)
        if result:
            results.append(result)
        else:
            print(f"  ⚠  '{law_name}' 처리 실패 — 건너뜀")

    if not results:
        print("\n처리된 법령이 없습니다.")
        sys.exit(1)

    # Excel 저장
    os.makedirs(args.output, exist_ok=True)
    date_str    = datetime.now().strftime("%Y%m%d")
    output_path = os.path.join(args.output, f"법규준수평가표_{date_str}.xlsm")

    print(f"\n{'='*60}")
    print(f"Excel 저장 중... ({len(results)}개 법령)")
    exporter.export_multi(output_path, results)

    print(f"\n완료!  ->  {output_path}")
    print(f"  Sheet 1: 법규준수평가")
    for idx, (name, df) in enumerate(results, start=2):
        print(f"  Sheet {idx}: {name}  ({len(df)}행)")


if __name__ == "__main__":
    main()
