# -*- coding: utf-8 -*-
# modules/construtor_mensagem.py

import os
import pandas as pd
from dotenv import load_dotenv
import locale
import logging

# Carrega variáveis de ambiente
load_dotenv()

# Pega os nomes das colunas do .env
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_VALOR = os.getenv('COLUNA_VALOR')
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO')
COLUNA_LOTE = os.getenv('LOTE')

# Obtém logger
logger = logging.getLogger(__name__)

# Configura locale
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    use_locale = True
except locale.Error:
    logger.warning("Locale 'pt_BR.UTF-8' não encontrado. Usando formatação de moeda manual.")
    use_locale = False

# --- Função criar_mensagem_cobranca ---
def criar_mensagem_cobranca(dados_cliente):
    """ Cria uma mensagem de cobrança personalizada para um cliente. """
    # Pega o nome já limpo pelo leitor_planilha.py
    nome = dados_cliente.get(COLUNA_NOME, "Cliente")
    logger.debug(f"Iniciando construção de mensagem para: {nome}") # Usa nome limpo no log

    # --- Extração e Formatação (Valor, Vencimento) ---
    valor_formatado = "[Valor Indisponível]"
    # Só tenta formatar se a coluna VALOR foi configurada E existe nos dados E não é NaN
    if COLUNA_VALOR and COLUNA_VALOR in dados_cliente and pd.notna(dados_cliente.get(COLUNA_VALOR)):
        try:
            valor_float = float(dados_cliente.get(COLUNA_VALOR))
            if use_locale: valor_formatado = locale.currency(valor_float, grouping=True, symbol='R$')
            else: valor_formatado = f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            logger.debug(f"Valor formatado: {valor_formatado}")
        except (ValueError, TypeError):
            logger.warning(f"Não formatou valor '{dados_cliente.get(COLUNA_VALOR)}' para {nome}.")

    vencimento_formatado = "[Data Indisponível]"
    # Só tenta formatar se a coluna VENCIMENTO foi configurada E existe E não é NaT E tem strftime
    vencimento_obj = dados_cliente.get(COLUNA_VENCIMENTO) if COLUNA_VENCIMENTO in dados_cliente else None
    if COLUNA_VENCIMENTO and pd.notna(vencimento_obj) and hasattr(vencimento_obj, 'strftime'):
       try:
           vencimento_formatado = vencimento_obj.strftime('%d/%m/%Y')
           logger.debug(f"Vencimento formatado: {vencimento_formatado}")
       except ValueError:
           logger.warning(f"Não formatou data '{vencimento_obj}' para {nome}.")


    # --- EXTRAÇÃO E PROCESSAMENTO DO LOTE/NÚMERO ---
    lote_info_processada = "[Ref não informada]" # Texto padrão
    # Verifica se LOTE foi definido no .env (COLUNA_LOTE não é None) E se a coluna existe nos dados
    if COLUNA_LOTE and COLUNA_LOTE in dados_cliente:
        # Pega valor usando o nome da coluna que veio do .env (guardado em COLUNA_LOTE)
        lote_raw = dados_cliente.get(COLUNA_LOTE)

        if lote_raw and isinstance(lote_raw, str):
            lote_texto = lote_raw.strip()
            if '/' in lote_texto:
                lote_info_processada = lote_texto.split('/')[0].strip() # Pega antes da barra
                logger.debug(f"Extraído Lote (antes de '/'): '{lote_info_processada}' de '{lote_texto}'")
            else:
                lote_info_processada = lote_texto # Usa tudo se não tem barra
                logger.debug(f"Usado Lote completo (sem '/'): '{lote_info_processada}'")
        elif lote_raw: # Lida com caso de ser número ou outro tipo
             lote_info_processada = str(lote_raw).strip()
             logger.debug(f"Usado Lote (convertido p/ string): '{lote_info_processada}'")
        # Se lote_raw for vazio/None, mantém o default.
    else:
         # Loga apenas se a variável foi definida mas a coluna não veio (improvável)
         if COLUNA_LOTE and COLUNA_LOTE not in dados_cliente:
              logger.warning(f"Coluna '{COLUNA_LOTE}' (var LOTE no .env) não encontrada nos dados de '{nome}'.")
         # Se COLUNA_LOTE é None (LOTE não no .env), não loga nada aqui.

    # --- Montagem da Mensagem ---
    # Usando uma das opções de texto discutidas (Opção 1 como exemplo)
    mensagem = (
        f"Olá {nome}! 👋\n\n"  # Saudação amigável
        f"Identificamos uma pendência em seu nome referente ao lote _{lote_info_processada}_.\n\n" # Italico para ref
        f"🗓️ Vencimento Original: {vencimento_formatado}\n"
        f"💰 Valor: *{valor_formatado}*\n\n" # Negrito para valor
        f"👇 Como regularizar ou tirar dúvidas:\n"
        f"1️⃣ Responda esta mensagem.\n"
        f"2️⃣ Ligue para: [Seu Número de Telefone]\n\n"
        f"Caso o pagamento já tenha sido efetuado, por favor, desconsidere esta mensagem.\n\n"
        f"Atenciosamente,\n"
        f"_[Nome da Sua Empresa]_" # Italico na assinatura
    )
    return mensagem