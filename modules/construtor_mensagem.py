# -*- coding: utf-8 -*-
# modules/construtor_mensagem.py

import os
import pandas as pd
from dotenv import load_dotenv
import locale
import logging # <--- Importado logging

# Carrega variáveis de ambiente
load_dotenv()

# Pega nomes das colunas
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_VALOR = os.getenv('COLUNA_VALOR')
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO')

# Obtém logger
logger = logging.getLogger(__name__)

# Configura locale
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    logger.info("Locale 'pt_BR.UTF-8' configurado para formatação de moeda.")
    use_locale = True
except locale.Error:
    logger.warning("Locale 'pt_BR.UTF-8' não encontrado. Usando formatação de moeda manual.")
    use_locale = False

def criar_mensagem_cobranca(dados_cliente):
    """
    Cria uma mensagem de cobrança personalizada.
    Args: dados_cliente (pandas.Series)
    Returns: str
    """
    logger.debug(f"Iniciando construção de mensagem para: {dados_cliente.get(COLUNA_NOME, 'Nome Desconhecido')}") # Nível DEBUG

    nome = dados_cliente.get(COLUNA_NOME, "Cliente")
    valor_raw = dados_cliente.get(COLUNA_VALOR)
    vencimento_obj = dados_cliente.get(COLUNA_VENCIMENTO)

    # Formata Valor
    valor_formatado = "[Valor Indisponível]"
    if pd.notna(valor_raw):
        try:
            valor_float = float(valor_raw)
            if use_locale:
                valor_formatado = locale.currency(valor_float, grouping=True, symbol='R$')
            else:
                valor_formatado = f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            logger.debug(f"Valor formatado: {valor_formatado}") # Nível DEBUG
        except (ValueError, TypeError):
            logger.warning(f"Não foi possível formatar o valor '{valor_raw}' como moeda para {nome}.")
            pass # Mantém "[Valor Indisponível]"

    # Formata Data de Vencimento
    vencimento_formatado = "[Data Indisponível]"
    if pd.notna(vencimento_obj) and hasattr(vencimento_obj, 'strftime'):
        try:
            vencimento_formatado = vencimento_obj.strftime('%d/%m/%Y')
            logger.debug(f"Vencimento formatado: {vencimento_formatado}") # Nível DEBUG
        except ValueError:
            logger.warning(f"Não foi possível formatar a data '{vencimento_obj}' como DD/MM/YYYY para {nome}.")
            pass # Mantém "[Data Indisponível]"

    # --- Monte sua Mensagem Aqui ---
    # **Adapte conforme sua necessidade.**
    mensagem = (
        f"Olá {nome},\n\n"
        f"Esperamos que esteja tudo bem.\n\n"
        f"Verificamos em nosso sistema um valor em aberto de {valor_formatado}, "
        f"referente ao vencimento em {vencimento_formatado}.\n\n"
        "Para regularizar sua situação ou tirar dúvidas, por favor, entre em contato conosco respondendo esta mensagem ou através do [Seu Número de Telefone/Link de Contato].\n\n"
        "Se o pagamento já foi efetuado, por favor, desconsidere esta mensagem.\n\n"
        "Agradecemos sua atenção,\n"
        "[Nome da Sua Empresa]"
    )
    logger.debug(f"Mensagem final construída para {nome}.") # Nível DEBUG
    return mensagem