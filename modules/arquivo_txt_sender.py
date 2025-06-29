# modules/arquivo_txt_sender.py

import os
import re
import logging 
from datetime import datetime

logger = logging.getLogger(__name__)

DIRETORIO_SAIDA = "mensagens_txt_geradas"

def limpar_nome_arquivo(nome): 
    nome = re.sub(r'[^\w\s-]', '', nome)
    nome = re.sub(r'\s+', '_', nome)
    return nome[:50]

def salvar_mensagem_em_txt(telefone_destino, nome_cliente, mensagem): 
    if not telefone_destino:
        logger.warning("Telefone de destino ausente ou inválido, usando 'telefone_na' no nome do arquivo.")
        telefone_destino = "telefone_na"
    if not nome_cliente:
        logger.warning("Nome do cliente ausente, usando 'cliente_desconhecido' no nome do arquivo.")
        nome_cliente = "cliente_desconhecido"

    try:
        os.makedirs(DIRETORIO_SAIDA, exist_ok=True)
    except OSError as e:
        logger.error(f"Erro ao criar o diretório de saída '{DIRETORIO_SAIDA}': {e}")
        return False

    nome_cliente_limpo = limpar_nome_arquivo(nome_cliente)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"cobranca_{telefone_destino}_{nome_cliente_limpo}_{timestamp}.txt"
    caminho_completo = os.path.join(DIRETORIO_SAIDA, nome_arquivo)

    try:
        with open(caminho_completo, 'w', encoding='utf-8') as f:
            f.write(f"Destinatário (Telefone): {telefone_destino}\n")
            f.write(f"Nome: {nome_cliente}\n")
            f.write("="*30 + "\n")
            f.write(mensagem) 
        logger.info(f"Mensagem para {nome_cliente} ({telefone_destino}) salva em: {caminho_completo}") # Mantido INFO pois é uma confirmação importante
        return True
    except IOError as e:
        logger.error(f"Erro de I/O ao salvar arquivo '{caminho_completo}': {e}")
        return False
    except Exception as e: 
        logger.exception(f"Erro inesperado ao salvar arquivo para {nome_cliente}")
        return False