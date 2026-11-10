"""Send the latest Opinet Excel report by Gmail SMTP.

Default behavior is a dry run. Add --send to actually send the email.
"""
from __future__ import annotations

import argparse
import mimetypes
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def load_env() -> dict[str, str]:
    env_path = ROOT / ".env"
    values: dict[str, str] = {}
    if not env_path.exists():
        raise FileNotFoundError(f"환경 파일이 없습니다: {env_path}")
    for raw in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def latest_report() -> Path:
    reports = list((ROOT / "data" / "opinet").glob("* /유가_기간비교.xlsx".replace("* ", "*")))
    reports = [p for p in reports if p.is_file()]
    if not reports:
        raise FileNotFoundError("첨부할 유가_기간비교.xlsx 파일을 찾지 못했습니다.")
    return max(reports, key=lambda p: p.stat().st_mtime)


def build_message(sender: str, recipient: str, attachment: Path, subject: str, body: str) -> EmailMessage:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    mime_type, _ = mimetypes.guess_type(attachment.name)
    maintype, subtype = (mime_type or "application/octet-stream").split("/", 1)
    message.add_attachment(attachment.read_bytes(), maintype=maintype, subtype=subtype, filename=attachment.name)
    return message


def main() -> int:
    parser = argparse.ArgumentParser(description="Gmail SMTP로 오피넷 엑셀 보고서를 발송합니다.")
    parser.add_argument("--send", action="store_true", help="실제 메일을 발송합니다. 생략하면 검증만 수행합니다.")
    parser.add_argument("--attachment", type=Path, help="첨부할 파일 경로")
    parser.add_argument("--to", help="받는 사람 주소. 기본값은 .env의 email_address")
    parser.add_argument("--subject", default="오피넷 유가 기간 비교 보고서", help="메일 제목")
    parser.add_argument("--body", default="오피넷 유가 기간 비교 엑셀 보고서를 첨부합니다.", help="메일 본문")
    args = parser.parse_args()

    env = load_env()
    sender = env.get("email_address", "")
    app_password = env.get("gmail_app_password", "")
    recipient = args.to or sender
    attachment = (args.attachment or latest_report()).resolve()

    if not sender or not app_password:
        raise ValueError(".env에 email_address와 gmail_app_password가 모두 필요합니다.")
    if not attachment.is_file():
        raise FileNotFoundError(f"첨부파일을 찾지 못했습니다: {attachment}")
    message = build_message(sender, recipient, attachment, args.subject, args.body)

    if not args.send:
        print("검증 완료: 실제 메일은 발송하지 않았습니다.")
        print(f"첨부파일: {attachment}")
        print(f"받는 사람: {recipient}")
        print(f"제목: {args.subject}")
        print(f"첨부 크기: {attachment.stat().st_size:,} bytes")
        return 0

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(sender, app_password)
        smtp.send_message(message)
    print("메일을 발송했습니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
