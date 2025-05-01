# -*- coding: utf-8 -*-
# modules/construtor_mensagem.py

import os
import pandas as pd # Importado para pd.Timestamp.min na ordenação
from dotenv import load_dotenv
import locale
import logging

# Carrega variáveis de ambiente (pode ser removido se não houver mais config aqui)
# load_dotenv()

# Obtém logger
logger = logging.getLogger(__name__)

# Configura locale (igual)
# try:
#     locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
#     use_locale = True # Flag não necessária aqui se formatação vem de fora
# except locale.Error:
#     logger.warning("Locale 'pt_BR.UTF-8' não encontrado.")
#     use_locale = False
# ------------------------------


# --- FUNÇÃO PARA MENSAGEM CONSOLIDADA ---
def criar_mensagem_consolidada(nome_cliente, lista_parcelas_info, valor_total_fmt):
    """
    Cria uma mensagem de cobrança consolidada para WhatsApp.
    Recebe dados já formatados.

    Args:
        nome_cliente (str): Nome do cliente (já limpo e formatado).
        lista_parcelas_info (list): Lista de dicts [{'lote': str, 'vencimento': str, 'valor': str}].
                                     O valor e vencimento já devem vir formatados.
        valor_total_fmt (str): O valor total já formatado como string R$.

    Returns:
        str: A mensagem formatada.
    """
    logger.debug(f"Criando msg consolidada para {nome_cliente} com {len(lista_parcelas_info)} parcela(s).")

# Monta a string com os detalhes das parcelas (FORMATO NOVO - MÚLTIPLAS LINHAS)
    detalhes_parcelas_str = ""
    if not lista_parcelas_info:
        detalhes_parcelas_str = "Não foi possível listar os detalhes das parcelas.\n"
    else:
        # Opcional: Ordenar a lista aqui se não foi feito no gerador
        # def sort_key(item): ...
        # lista_parcelas_info.sort(key=sort_key)

        for i, parcela in enumerate(lista_parcelas_info):
            lote = parcela.get('lote', '[N/D]')
            venc = parcela.get('vencimento', '[N/D]')
            valor = parcela.get('valor', '[N/D]') # Valor já vem formatado

            # Adiciona as múltiplas linhas formatadas
            detalhes_parcelas_str += f"{i+1}. Lote: _{lote}_\n"      # <-- ALTERADO AQUI
            detalhes_parcelas_str += f"     Venc: {venc}\n"      # Indenta com espaços
            detalhes_parcelas_str += f"     Valor: *{valor}*\n\n" # Indenta, Negrito no Valor, Linha extra

    # --- Template da Mensagem (Exemplo 1 - ajuste se necessário) ---
    mensagem = (
        f"Olá {nome_cliente}! 👋\n\n"
        f"Identificamos as seguintes pendências:\n\n"
        f"{detalhes_parcelas_str}\n" # Insere a lista de parcelas
        f"*Valor Total: {valor_total_fmt}*\n\n" # Mostra o total geral
        f"👇 Como regularizar ou tirar dúvidas:\n"
        f"1️⃣ Responda esta mensagem.\n"
        f"2️⃣ Ligue para: [Seu Número de Telefone]\n\n"
        f"Caso o pagamento já tenha sido efetuado, por favor, desconsidere esta mensagem.\n\n"
        f"Atenciosamente,\n"
        f"_[Nome da Sua Empresa]_"
    )

    return mensagem