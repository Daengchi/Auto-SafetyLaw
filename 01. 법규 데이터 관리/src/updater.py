"""
기존 법규준수평가표 .xlsm 파일 업데이트.

- 각 법령 시트(2번~)를 최신 API 조문으로 교체
- 기존 D열(해당여부) O/X 보존
- 신설·변경 조문: 노란색 배경 (최신 업데이트 분만)
- '변경사항' 시트: 전체 변경 이력 누적
"""
import re
from datetime import datetime
from typing import Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from src.exporter import (
    _style_sheet, _add_haedan_dropdown, _add_title_row, _set_print_layout,
    _HEADER_FILL, _HEADER_FONT, _HEADER_ALIGN, _BORDER,
)


def _header_row(ws) -> int:
    """헤더('No.')가 위치한 행 탐색. 제목 행 유무(구·신 레이아웃) 모두 대응."""
    for r in (1, 2, 3):
        if str(ws.cell(r, 1).value).strip() == "No.":
            return r
    return 1

_YELLOW_FILL   = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
_HISTORY_SHEET = "변경사항"
_HISTORY_HEADERS = ["업데이트 날짜", "법령명", "조문번호", "조문명", "변경유형"]

_ART_NUM_RE = re.compile(r'^제\d+조(?:의\d+)?')
_COLS       = ["A", "B", "C"]
_COL_NAMES  = ["법률 조문", "시행령 조문", "시행규칙 조문"]

_THIN   = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


# ─── 셀 파싱 헬퍼 ─────────────────────────────────────────────────────────────

def _parse_art_cell(value: str) -> tuple[str, str]:
    """"제118조\n규제의 재검토" → ("제118조", "규제의 재검토")"""
    if not value:
        return "", ""
    parts = str(value).split("\n", 1)
    num   = parts[0].strip()
    title = parts[1].strip() if len(parts) > 1 else ""
    return num, title


def _primary_key_col(a, b, c) -> tuple[str, str, str]:
    """
    A→B→C 순으로 첫 번째 유효 조문 반환.
    반환: (열 레터, 조문번호, 조문명)
    단독 부령처럼 번호가 "1, 2, 3" 형태여도 허용.
    """
    for col, val in zip(_COLS, [a, b, c]):
        num, title = _parse_art_cell(str(val) if val else "")
        if num:
            return col, num, title
    return "", "", ""


# ─── O/X 보존 ────────────────────────────────────────────────────────────────

def _read_ox_map(ws, hdr: int = 1) -> tuple[dict, dict]:
    """
    열+번호 조합 키로 O/X 추출.
    반환: (ox_by_num={(col,num):ox}, ox_by_title={(col,title):ox})
    """
    ox_by_num:   dict[tuple, str] = {}
    ox_by_title: dict[tuple, str] = {}

    for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
        if len(row) < 5:
            continue
        a, b, c, d = row[1], row[2], row[3], row[4]
        ox = str(d).strip() if d else ""
        if not ox:
            continue
        col, num, title = _primary_key_col(a, b, c)
        if col and num:
            ox_by_num.setdefault((col, num), ox)
        if col and title:
            ox_by_title.setdefault((col, title), ox)

    return ox_by_num, ox_by_title


def _row_triple(row, is_admrul: bool) -> tuple[str, str, str]:
    """새 DataFrame 행에서 키 산출용 (a, b, c) 추출.
    행정규칙은 조문제목(B)만 키로 사용."""
    if is_admrul:
        return str(row.get("조문제목", "") or ""), "", ""
    return row["법률 조문"], row["시행령 조문"], row["시행규칙 조문"]


def _apply_ox(new_df: pd.DataFrame,
              ox_by_num: dict, ox_by_title: dict,
              is_admrul: bool = False) -> pd.DataFrame:
    """새 DataFrame 각 행에 기존 O/X 매핑."""
    result = new_df.copy()
    for idx, row in result.iterrows():
        col, num, title = _primary_key_col(*_row_triple(row, is_admrul))
        ox = (ox_by_num.get((col, num)) or
              ox_by_title.get((col, title)) or "")
        result.at[idx, "해당여부"] = ox
    return result


# ─── 변경 감지 ────────────────────────────────────────────────────────────────

def _read_old_articles(ws, hdr: int = 1) -> dict:
    """기존 시트에서 {(col, num): set_of_titles} 반환.
    동일 조문번호가 여러 행에 걸쳐 있는 경우(의N 파싱 오류 호환)를 위해 set 사용."""
    result: dict[tuple, set] = {}
    for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
        if len(row) < 4:
            continue
        col, num, title = _primary_key_col(
            row[1] or "", row[2] or "", row[3] or ""
        )
        if col and num:
            result.setdefault((col, num), set()).add(title)
    return result


def _detect_changes(old_articles: dict,
                    new_df: pd.DataFrame,
                    is_admrul: bool = False) -> tuple[list, list]:
    """
    old_articles: {(col, num): set_of_titles}
    반환: (change_types, deleted_list)
      change_types: 각 행당 "동일" | "신설" | "변경"
      deleted_list: [{"조문번호", "조문명", "변경유형": "삭제"}, ...]
    """
    # (col, title) → num 역방향 맵 (각 col에서 첫 번째 제목 기준)
    old_by_title: dict[tuple, str] = {}
    for (col, num), titles in old_articles.items():
        for t in titles:
            old_by_title.setdefault((col, t), num)

    matched_old: set = set()
    change_types: list = []

    for _, row in new_df.iterrows():
        col, num, title = _primary_key_col(*_row_triple(row, is_admrul))
        if (col, num) in old_articles:
            matched_old.add((col, num))
            old_titles = old_articles[(col, num)]
            change_types.append("동일" if title in old_titles else "변경")
        elif (col, title) in old_by_title:
            old_num = old_by_title[(col, title)]
            matched_old.add((col, old_num))
            change_types.append("변경")          # 번호 이동
        else:
            change_types.append("신설")

    deleted = [
        {"조문번호": num, "조문명": title, "변경유형": "삭제"}
        for (col, num), titles in old_articles.items()
        if (col, num) not in matched_old
        for title in titles
    ]
    return change_types, deleted


# ─── 시트 재작성 ─────────────────────────────────────────────────────────────

def _rewrite_sheet(ws, new_df: pd.DataFrame, change_types: list,
                   law_name: str, is_admrul: bool = False) -> None:
    """기존 내용 전체 삭제 후 제목행+헤더+데이터 재작성. 신설·변경 행은 노란색.
    레이아웃: 1행=제목 / 2행=헤더 / 3행~=데이터 (구 레이아웃 파일도 신 레이아웃으로 정규화)."""
    for mc in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mc))
    if ws.max_row >= 1:
        ws.delete_rows(1, ws.max_row)

    ncols = len(new_df.columns)
    _add_title_row(ws, law_name, ncols)                    # 1행: 제목
    for col_idx, col_name in enumerate(new_df.columns, start=1):
        ws.cell(row=2, column=col_idx, value=col_name)     # 2행: 헤더

    for _, row in new_df.iterrows():                        # 3행~: 데이터
        if is_admrul:
            ws.append([
                row["No."], row.get("조문제목", ""), row.get("조문내용", ""),
                "", row["해당여부"],
            ])
        else:
            ws.append([
                row["No."], row["법률 조문"], row["시행령 조문"],
                row["시행규칙 조문"], row["해당여부"],
            ])

    _style_sheet(ws, new_df, header_row=2)

    for row_idx, ct in enumerate(change_types, start=3):
        if ct in ("신설", "변경"):
            for col_idx in range(1, 5):
                ws.cell(row=row_idx, column=col_idx).fill = _YELLOW_FILL

    _add_haedan_dropdown(ws, len(new_df), start_row=3)
    _set_print_layout(ws)
    if is_admrul:
        ws.column_dimensions["D"].hidden = True  # 미사용 placeholder 열


# ─── 변경사항 시트 ────────────────────────────────────────────────────────────

def _ensure_history_sheet(wb):
    """변경사항 시트 없으면 생성, 있으면 반환."""
    if _HISTORY_SHEET in wb.sheetnames:
        return wb[_HISTORY_SHEET]

    ws = wb.create_sheet(_HISTORY_SHEET)
    ws.append(_HISTORY_HEADERS)
    widths = [14, 36, 14, 32, 10]
    for col_idx, (h, w) in enumerate(zip(_HISTORY_HEADERS, widths), start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = _HEADER_FILL
        cell.font      = _HEADER_FONT
        cell.alignment = _HEADER_ALIGN
        cell.border    = _BORDER
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28
    return ws


def _append_history(ws_hist, today: str,
                    law_name: str, changes: list) -> None:
    """변경사항 시트에 이력 행 추가."""
    for change in changes:
        ws_hist.append([
            today,
            law_name,
            change["조문번호"],
            change["조문명"],
            change["변경유형"],
        ])
        row_idx = ws_hist.max_row
        for col_idx in range(1, 6):
            cell = ws_hist.cell(row=row_idx, column=col_idx)
            cell.font      = Font(size=9)
            cell.alignment = Alignment(vertical="center")
            cell.border    = _BORDER


# ─── 진입점 ───────────────────────────────────────────────────────────────────

def update_file(
    xlsx_path: str,
    process_fn: Callable[[str], tuple[str, pd.DataFrame] | None],
) -> None:
    """
    xlsx_path:  업데이트할 .xlsm / .xlsx 경로
    process_fn: law_name → (actual_name, combined_df) | None
    """
    print(f"\n파일 로드 중: {xlsx_path}")
    wb   = load_workbook(xlsx_path, keep_vba=True)
    today = datetime.now().strftime("%Y/%m/%d")

    updated, skipped = 0, 0
    all_changes: list[tuple[str, dict]] = []   # (law_name, change_dict)

    for sheet in wb.worksheets:
        title = sheet.title
        if ". " not in title:
            continue
        law_name = title.split(". ", 1)[1]
        if law_name in ("법규준수평가", _HISTORY_SHEET):
            continue

        print(f"\n  [{law_name}] 처리 중...")

        hdr = _header_row(sheet)
        ox_by_num, ox_by_title = _read_ox_map(sheet, hdr)
        old_articles           = _read_old_articles(sheet, hdr)
        preserved = sum(1 for v in ox_by_num.values() if v == "O")
        print(f"    기존 O 표시 조문: {preserved}개 보존 예정")

        result = process_fn(law_name)
        if result is None:
            print(f"    건너뜀 (API 조회 실패)")
            skipped += 1
            continue

        _, new_df = result
        is_admrul = "조문제목" in new_df.columns
        new_df = _apply_ox(new_df, ox_by_num, ox_by_title, is_admrul)

        change_types, deleted = _detect_changes(old_articles, new_df, is_admrul)

        cnt = {t: change_types.count(t) for t in ("신설", "변경", "동일")}
        print(f"    완료 ({len(new_df)}행) — 신설 {cnt['신설']}, 변경 {cnt['변경']}, 삭제 {len(deleted)}")

        _rewrite_sheet(sheet, new_df, change_types, law_name, is_admrul)
        updated += 1

        # 이력 수집
        for i, (_, row) in enumerate(new_df.iterrows()):
            if change_types[i] in ("신설", "변경"):
                col, num, title = _primary_key_col(*_row_triple(row, is_admrul))
                all_changes.append((law_name,
                    {"조문번호": num, "조문명": title, "변경유형": change_types[i]}))
        for d in deleted:
            all_changes.append((law_name, d))

    # 변경사항 시트 갱신
    if all_changes:
        ws_hist = _ensure_history_sheet(wb)
        # 법령명별로 묶어서 기록
        current_law, batch = "", []
        for law_name, change in all_changes + [("__END__", None)]:
            if law_name != current_law:
                if batch:
                    _append_history(ws_hist, today, current_law, batch)
                current_law, batch = law_name, []
            if change:
                batch.append(change)
        print(f"\n변경사항 시트: {len(all_changes)}건 기록")
    else:
        print("\n변경사항 없음 — 이력 미기록")

    wb.save(xlsx_path)
    print(f"\n{'─'*50}")
    print(f"업데이트 완료: {updated}개 시트 갱신, {skipped}개 건너뜀")
    print(f"저장: {xlsx_path}")
    print()
    print("※ '1. 법규준수평가' 시트를 갱신하려면 파일을 열고")
    print("  임의 조문의 해당여부(O/X)를 변경하면 자동 반영됩니다.")
