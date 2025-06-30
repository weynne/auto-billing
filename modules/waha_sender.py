# modules/waha_sender.py (Versão Final Corrigida com Sessão no Payload)
import requests
import logging
import time
import os

logger = logging.getLogger(__name__)

WAHA_BASE_URL = os.getenv("WAHA_BASE_URL", "http://localhost:3000")
WAHA_API_KEY = os.getenv("WAHA_API_KEY")
WAHA_SESSION_NAME = os.getenv("WAHA_SESSION_NAME")
WAHA_SEND_TEXT_URL = f"{WAHA_BASE_URL}/api/sendText"

def enviar_mensagem_waha(telefone, mensagem):
    if not all([WAHA_API_KEY, WAHA_SESSION_NAME]):
        logger.error("Configurações do WAHA (KEY ou SESSION_NAME) não encontradas no .env.")
        return False

    if not telefone or not telefone.startswith('+'):
        logger.error(f"Telefone '{telefone}' é inválido. Pulando.")
        return False

    chat_id = f"{telefone.replace('+', '')}@c.us"

    # MUDANÇA: Adicionamos a chave "session" ao payload
    payload = {
        "chatId": chat_id,
        "text": mensagem,
        "session": WAHA_SESSION_NAME
    }
    
    headers = {
        "Content-Type": "application/json",
        "X-Api-Key": WAHA_API_KEY
    }

    try:
        logger.info(f"Enviando mensagem para {chat_id} na sessão '{WAHA_SESSION_NAME}'...")
        response = requests.post(WAHA_SEND_TEXT_URL, json=payload, headers=headers, timeout=20)

        if response.status_code == 201:
            logger.info(f"Mensagem para {chat_id} enfileirada com sucesso.")
            time.sleep(5)
            return True
        else:
            logger.error(f"Falha ao enviar via WAHA. Status: {response.status_code}, Resposta: {response.text}")
            return False

    except requests.exceptions.RequestException as e:
        logger.error(f"Erro de conexão com o servidor WAHA ({WAHA_SEND_TEXT_URL}): {e}")
        return False