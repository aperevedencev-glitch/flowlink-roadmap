"""Оповещения оператора: Telegram и/или почта."""
from __future__ import annotations

import logging
import smtplib
from email.message import EmailMessage

import requests

log = logging.getLogger(__name__)


def notify(settings: dict, subject: str, text: str) -> None:
    sent = False
    token, chat = settings.get("telegram_bot_token"), settings.get("telegram_chat_id")
    if token and chat:
        try:
            requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
                          data={"chat_id": chat, "text": f"{subject}\n\n{text}"[:4000]}, timeout=20)
            sent = True
        except requests.RequestException as exc:
            log.error("Не отправлено в Telegram: %s", exc)
    if settings.get("smtp_host") and settings.get("email_to"):
        try:
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = settings.get("smtp_user") or "eis2nostroy@localhost"
            msg["To"] = settings["email_to"]
            msg.set_content(text)
            with smtplib.SMTP(settings["smtp_host"], int(settings.get("smtp_port", 587)), timeout=30) as s:
                s.starttls()
                if settings.get("smtp_user"):
                    s.login(settings["smtp_user"], settings.get("smtp_password", ""))
                s.send_message(msg)
            sent = True
        except (OSError, smtplib.SMTPException) as exc:
            log.error("Письмо не отправлено: %s", exc)
    if not sent:
        log.info("%s\n%s", subject, text)
