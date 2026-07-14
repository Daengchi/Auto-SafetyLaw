"""
amendments.json 데이터로 이메일 제목/본문(plain text + HTML) 생성.

입력 data 구조 (02. 제개정사항 모니터링이 출력):
{
  "생성일시": "2026-06-02T09:00",
  "개정": [
    {"법령명", "이전시행일", "현재시행일", "개정조문수",
     "조문": [{"조문번호", "조문명", "구법내용", "신법내용"}, ...]},
    ...
  ]
}
조문 내용(구법/신법)에는 변경 부분을 감싼 <P>...</P> 마커가 포함된다.
"""
import re
from datetime import datetime

MAX_CONTENT_LEN = 300   # 조문 내용이 길면 잘라서 표시
GREETING = "아래의 내용과 같이 시행된 개정법령 사항들을 안내드리오니 업무에 참조하시기 바랍니다"

BRAND = "#ED1C24"       # 회사 공식 레드 (RGB 237,28,36)

_P_RE = re.compile(r'<P>(.*?)</P>', re.IGNORECASE | re.DOTALL)


def _strip_p(text: str) -> str:
    """<P> 마커 제거 (plain text용)."""
    return _P_RE.sub(r'\1', text or "")


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _article_title(art: dict) -> str:
    """조문 제목: '제33조(안전보건교육기관)'.
    조문번호는 '제33조'·'별표 1'처럼 이미 완성된 표기이므로 그대로 쓴다."""
    제목 = art.get("조문번호", "")
    if art.get("조문명"):
        제목 += f"({art['조문명']})"
    return 제목


def _trim(text: str) -> str:
    text = (text or "").strip()
    if len(text) <= MAX_CONTENT_LEN:
        return text
    return text[:MAX_CONTENT_LEN] + " …(생략)"


# ─── plain text ───────────────────────────────────────────────────────────────

def _build_text(amendments: list, today: str) -> str:
    lines = [GREETING, "", f"(기준일: {today} · 총 {len(amendments)}개 법령)", ""]
    for law in amendments:
        lines.append("=" * 60)
        lines.append(f"■ {law['법령명']}")
        lines.append(f"   시행일: {law['이전시행일']} → {law['현재시행일']}")
        lines.append(f"   개정 조문 수: {law['개정조문수']}개")
        lines.append("")
        for art in law.get("조문", []):
            제목 = _article_title(art)
            lines.append(f"  · {제목}")
            lines.append(f"    [구법] {_trim(_strip_p(art.get('구법내용', '')))}")
            lines.append(f"    [신법] {_trim(_strip_p(art.get('신법내용', '')))}")
            lines.append("")
    return "\n".join(lines)


# ─── HTML ─────────────────────────────────────────────────────────────────────

def _render_content_html(text: str) -> str:
    """<P>...</P> 구간을 브랜드 레드 볼드로 변환, 나머지는 이스케이프 + 줄바꿈 처리."""
    raw = (text or "").strip()
    # 가시 길이(태그 제외)가 길면 마커 제거 후 잘라 강조 없이 표시
    if len(_strip_p(raw)) > MAX_CONTENT_LEN:
        return _esc(_trim(_strip_p(raw))).replace("\n", "<br>")
    out, last = [], 0
    for m in _P_RE.finditer(raw):
        if m.start() > last:
            out.append(_esc(raw[last:m.start()]))
        inner = _esc(m.group(1))
        if inner:
            out.append(f'<b style="color:{BRAND}">{inner}</b>')
        last = m.end()
    if last < len(raw):
        out.append(_esc(raw[last:]))
    return "".join(out).replace("\n", "<br>")


def _article_html(art: dict) -> str:
    제목 = _article_title(art)
    old = (art.get("구법내용") or "").strip()
    new = (art.get("신법내용") or "").strip()

    title_html = (
        f'<div style="font-size:14px;font-weight:bold;color:#222222;'
        f'margin:18px 0 8px;">{_esc(제목)}</div>'
    )

    def box(label, label_color, bg, border, text_color, body_html, strike=False):
        deco = "text-decoration:line-through;" if strike else ""
        return (
            f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
            f'style="margin-bottom:6px;"><tr><td style="background-color:{bg};'
            f'{border}border-radius:4px;padding:10px 13px;font-size:13px;'
            f'color:{text_color};line-height:1.6;{deco}">'
            f'<span style="font-size:11px;font-weight:bold;color:{label_color};">'
            f'{label}</span><br>{body_html}</td></tr></table>'
        )

    parts = [title_html]
    if not old:          # 신설
        parts.append(box("신설", BRAND, "#fff0f1",
                         f"border-left:3px solid {BRAND};", "#222222",
                         _render_content_html(new)))
    elif not new:        # 삭제
        parts.append(box("삭제", "#999999", "#f4f4f4", "", "#999999",
                         _render_content_html(old), strike=True))
    else:                # 개정
        parts.append(box("구법", "#aaaaaa", "#f4f4f4", "", "#888888",
                         _render_content_html(old)))
        parts.append(box("신법", BRAND, "#fff0f1",
                         f"border-left:3px solid {BRAND};", "#222222",
                         _render_content_html(new)))
    return "".join(parts)


def _law_card_html(law: dict) -> str:
    arts = "".join(_article_html(a) for a in law.get("조문", []))
    return (
        '<tr><td style="padding:14px 28px;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="border:1px solid #eeeeee;border-radius:6px;">'
        # 법령 헤더
        f'<tr><td style="padding:16px 20px;border-left:4px solid {BRAND};'
        'background-color:#fff8f8;">'
        f'<span style="font-size:16px;font-weight:bold;color:{BRAND};'
        f'background-color:#FFF0F1;padding:2px 8px;border-radius:3px;">'
        f'{_esc(law["법령명"])}</span>'
        '<div style="font-size:13px;color:#666666;margin-top:10px;">'
        f'시행일 {_esc(law["이전시행일"])} '
        f'<span style="color:#999999;">&rarr;</span> '
        f'<b style="color:{BRAND}">{_esc(law["현재시행일"])}</b>'
        '&nbsp;&nbsp;·&nbsp;&nbsp;'
        f'개정 조문 <b>{law["개정조문수"]}</b>개</div></td></tr>'
        # 조문들
        f'<tr><td style="padding:4px 20px 16px;">{arts}</td></tr>'
        '</table></td></tr>'
    )


def _build_html(amendments: list, today: str) -> str:
    cards = "".join(_law_card_html(law) for law in amendments)
    return (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0"></head>'
        '<body style="margin:0;padding:0;background-color:#f0f0f2;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="background-color:#f0f0f2;padding:24px 0;'
        "font-family:'Malgun Gothic','맑은 고딕',Arial,sans-serif;\">"
        '<tr><td align="center">'
        '<table role="presentation" width="640" cellpadding="0" cellspacing="0" '
        'style="width:640px;max-width:640px;background-color:#ffffff;'
        'border-radius:8px;overflow:hidden;border:1px solid #e5e5e5;">'
        # 헤더 배너
        f'<tr><td style="background-color:{BRAND};padding:26px 32px;">'
        '<div style="color:#ffffff;font-size:21px;font-weight:bold;">'
        '법규 개정 알림</div>'
        '<div style="color:#ffd9da;font-size:13px;margin-top:7px;">'
        f'기준일 {today} &nbsp;·&nbsp; 총 {len(amendments)}개 법령 개정</div></td></tr>'
        # 인사 문구
        '<tr><td style="padding:24px 32px 4px;color:#333333;font-size:14px;'
        f'line-height:1.7;">{_esc(GREETING)}</td></tr>'
        # 법령 카드들
        f'{cards}'
        # 푸터
        '<tr><td style="background-color:#fafafa;padding:18px 32px;color:#999999;'
        'font-size:12px;border-top:1px solid #eeeeee;line-height:1.6;">'
        '본 메일은 안전보건 법규 모니터링 시스템에서 자동 발송되었습니다.</td></tr>'
        '</table></td></tr></table></body></html>'
    )


# ─── 진입점 ───────────────────────────────────────────────────────────────────

def build(data: dict) -> tuple[str, str, str]:
    """반환: (제목, 텍스트 본문, HTML 본문). 개정 항목이 없으면 호출하지 않는다."""
    amendments = data.get("개정", [])
    today = datetime.now().strftime("%Y-%m-%d")
    subject = f"[법규 개정 알림] {len(amendments)}개 법령 개정 ({today})"
    text_body = _build_text(amendments, today)
    html_body = _build_html(amendments, today)
    return subject, text_body, html_body


def build_heartbeat(data: dict) -> tuple[str, str, str]:
    """
    개정이 없는 주에 발송하는 생존 확인 메일.
    이 메일이 오지 않으면 자동 실행이 멈춘 것이므로, 침묵과 정상을 구분해준다.
    """
    today   = datetime.now().strftime("%Y-%m-%d")
    점검건수 = data.get("점검건수", "-")
    오류건수 = data.get("오류건수", 0)
    최근점검 = data.get("생성일시", "-").replace("T", " ")

    상태  = "정상" if not 오류건수 else f"주의 - 조회 오류 {오류건수}건"
    색상  = BRAND if not 오류건수 else "#C77700"
    subject = f"[법규 모니터링] 주간 정상 동작 확인 ({today})"

    text_body = "\n".join([
        "안전보건 법규 제개정 모니터링 시스템이 정상 동작 중입니다.",
        "",
        f"  상태      : {상태}",
        f"  최근 점검  : {최근점검}",
        f"  점검 법령  : {점검건수}건",
        f"  개정 사항  : 없음",
        f"  조회 오류  : {오류건수}건",
        "",
        "이 메일은 개정 사항이 없는 주에 1회(월요일) 발송됩니다.",
        "메일이 도착하지 않으면 자동 실행이 중단된 것이므로 확인이 필요합니다.",
    ])

    def row(label, value):
        return (
            '<tr>'
            '<td style="padding:7px 0;color:#888888;font-size:13px;width:110px;">'
            f'{_esc(label)}</td>'
            '<td style="padding:7px 0;color:#333333;font-size:13px;font-weight:bold;">'
            f'{_esc(str(value))}</td></tr>'
        )

    html_body = (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width, initial-scale=1.0"></head>'
        '<body style="margin:0;padding:0;background-color:#f0f0f2;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="background-color:#f0f0f2;padding:24px 0;'
        "font-family:'Malgun Gothic','맑은 고딕',Arial,sans-serif;\">"
        '<tr><td align="center">'
        '<table role="presentation" width="640" cellpadding="0" cellspacing="0" '
        'style="width:640px;max-width:640px;background-color:#ffffff;'
        'border-radius:8px;overflow:hidden;border:1px solid #e5e5e5;">'
        f'<tr><td style="background-color:{색상};padding:26px 32px;">'
        '<div style="color:#ffffff;font-size:21px;font-weight:bold;">'
        '모니터링 정상 동작 확인</div>'
        '<div style="color:#ffffff;opacity:0.85;font-size:13px;margin-top:7px;">'
        f'기준일 {today} &nbsp;·&nbsp; 상태 {_esc(상태)}</div></td></tr>'
        '<tr><td style="padding:24px 32px 8px;color:#333333;font-size:14px;'
        'line-height:1.7;">이번 주 점검 결과 개정된 법령은 없습니다. '
        '시스템은 정상 동작 중입니다.</td></tr>'
        '<tr><td style="padding:8px 32px 24px;">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="border-top:1px solid #eeeeee;">'
        + row("최근 점검", 최근점검)
        + row("점검 법령", f"{점검건수}건")
        + row("개정 사항", "없음")
        + row("조회 오류", f"{오류건수}건")
        + '</table></td></tr>'
        '<tr><td style="background-color:#fafafa;padding:18px 32px;color:#999999;'
        'font-size:12px;border-top:1px solid #eeeeee;line-height:1.6;">'
        '이 메일은 개정 사항이 없는 주에 1회(월요일) 발송됩니다. '
        '메일이 도착하지 않으면 자동 실행이 중단된 것이므로 확인이 필요합니다.'
        '</td></tr>'
        '</table></td></tr></table></body></html>'
    )
    return subject, text_body, html_body
