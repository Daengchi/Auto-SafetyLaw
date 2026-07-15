"""
사내 SMTP 서버로 개정 알림 메일 발송.
SMTP 설정은 .env 환경변수에서 읽는다.

헤더 배너는 메일에 함께 담아(CID) 보낸다. 외부 URL로 걸면 Outlook이
기본적으로 차단해 빈 네모로 보이기 때문이다.
"""
import os
import smtplib
import sys
from email.message import EmailMessage

from src.composer import IMAGES

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")


def send(subject: str, text_body: str, html_body: str, recipients: list[str]) -> None:
    host     = os.getenv("SMTP_HOST")
    port     = int(os.getenv("SMTP_PORT", "25"))
    user     = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    sender   = os.getenv("SMTP_FROM")
    use_tls  = os.getenv("SMTP_USE_TLS", "false").lower() == "true"

    if not host or not sender:
        print("오류: SMTP 설정이 없습니다.")
        print("  .env 파일에 SMTP_HOST, SMTP_FROM 등을 설정하세요 (.env.example 참고).")
        sys.exit(1)
    if not recipients:
        print("오류: 수신자가 없습니다. recipients.json을 확인하세요.")
        sys.exit(1)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"]    = sender
    msg["To"]      = ", ".join(recipients)
    msg.set_content(text_body)                       # plain text fallback
    msg.add_alternative(html_body, subtype="html")   # HTML (우선 표시)

    # HTML이 참조하는 이미지만 CID로 첨부
    html_part = msg.get_payload()[1]
    for cid, (filename, _) in IMAGES.items():
        if f"cid:{cid}" not in html_body:
            continue
        subtype = "jpeg" if filename.lower().endswith((".jpg", ".jpeg")) else "png"
        with open(os.path.join(ASSETS_DIR, filename), "rb") as f:
            html_part.add_related(f.read(), maintype="image",
                                  subtype=subtype, cid=f"<{cid}>")

    with smtplib.SMTP(host, port, timeout=30) as smtp:
        if use_tls:
            smtp.starttls()
        if user:
            smtp.login(user, password)
        smtp.send_message(msg)

    print(f"메일 발송 완료: {len(recipients)}명  ->  {', '.join(recipients)}")
