"""
법규 제개정 이메일 알림 - CLI

02. 제개정사항 모니터링이 출력한 amendments.json을 읽어,
개정이 있으면 recipients.json의 수신자에게 사내 SMTP로 메일을 발송한다.

개정이 없어도 매주 월요일에는 생존 확인 메일을 관리자에게 발송한다.
(메일이 오지 않는 것이 "개정 없음"인지 "실행 중단"인지 구분하기 위함)

사용 예시:
  python main.py              # amendments.json 읽고 개정 있으면 메일 발송
  python main.py --dry-run    # 발송하지 않고 제목/본문만 콘솔 출력 (테스트)
  python main.py --heartbeat --dry-run   # 생존 확인 메일 미리보기 (요일 무관)
"""
import argparse
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dotenv import load_dotenv

from src import composer, emailer


AMENDMENTS_FILE = os.path.join(os.path.dirname(__file__), "..", "프로그램", "amendments.json")
RECIPIENTS_FILE = os.path.join(os.path.dirname(__file__), "recipients.json")


def main() -> None:
    load_dotenv()

    arg_parser = argparse.ArgumentParser(
        description="법규 제개정 이메일 알림",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    arg_parser.add_argument("--dry-run", action="store_true",
                            help="발송하지 않고 제목/본문만 콘솔 출력")
    arg_parser.add_argument("--heartbeat", action="store_true",
                            help="요일과 무관하게 생존 확인 메일을 생성 (테스트용)")
    args = arg_parser.parse_args()

    if not os.path.exists(AMENDMENTS_FILE):
        print(f"amendments.json이 없습니다 ({AMENDMENTS_FILE}).")
        print("  먼저 02. 제개정사항 모니터링을 실행하세요.")
        sys.exit(0)

    with open(AMENDMENTS_FILE, encoding="utf-8") as f:
        data = json.load(f)

    if data.get("개정") and not args.heartbeat:
        group = "개정알림"
        subject, text_body, html_body = composer.build(data)
    elif args.heartbeat or datetime.now().weekday() == 0:   # 월요일 = 생존 확인
        group = "생존확인"
        subject, text_body, html_body = composer.build_heartbeat(data)
    else:
        print("개정된 법령이 없습니다. 발송 대상 없음.")
        return

    if args.dry_run:
        preview = os.path.join(os.path.dirname(__file__), "email_preview.html")
        with open(preview, "w", encoding="utf-8") as f:
            f.write(html_body)
        print(f"제목: {subject}\n{'─' * 60}\n{text_body}")
        print(f"{'─' * 60}\nHTML 미리보기 저장: {preview}")
        return

    with open(RECIPIENTS_FILE, encoding="utf-8") as f:
        recipients = json.load(f)

    emailer.send(subject, text_body, html_body, recipients[group])


if __name__ == "__main__":
    main()
