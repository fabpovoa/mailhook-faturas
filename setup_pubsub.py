"""
Script auxiliar — configura o Gmail Watch (Pub/Sub) para receber notificações.
Execute UMA VEZ para ativar o push do Gmail.

Pré-requisito: token.json já criado (rode o servidor uma vez para autenticar).
"""
import os
from dotenv import load_dotenv
load_dotenv()

from gmail import get_gmail_service

GMAIL_USER = os.getenv("GMAIL_USER", "me")
PROJECT_ID = os.getenv("PUBSUB_PROJECT_ID")
TOPIC = os.getenv("PUBSUB_TOPIC", "gmail-faturas")


def watch():
    service = get_gmail_service()
    topic_name = f"projects/{PROJECT_ID}/topics/{TOPIC}"

    response = service.users().watch(
        userId=GMAIL_USER,
        body={
            "topicName": topic_name,
            "labelIds": ["INBOX"],
            "labelFilterAction": "include",
        },
    ).execute()

    print(f"✅ Gmail Watch ativado!")
    print(f"   historyId: {response.get('historyId')}")
    print(f"   expiration: {response.get('expiration')} (milissegundos Unix — expira em ~7 dias)")
    print()
    print("⚠️  Renove semanalmente rodando este script novamente.")


def stop_watch():
    """Para de receber notificações (use se quiser cancelar)."""
    service = get_gmail_service()
    service.users().stop(userId=GMAIL_USER).execute()
    print("🛑 Gmail Watch desativado.")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "stop":
        stop_watch()
    else:
        watch()
