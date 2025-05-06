# -*- coding: utf-8 -*-
# modules/construtor_mensagem.py

import logging

logger = logging.getLogger(__name__)

def criar_mensagem_consolidada(nome_cliente, lista_parcelas_info, valor_total_fmt, nome_empresa):
 
    logger.debug(f"Criando msg consolidada para {nome_cliente} com {len(lista_parcelas_info)} parcela(s) pela empresa {nome_empresa}.")

    num_parcelas = len(lista_parcelas_info)

    # --- Adapta texto introdutório ---
    if num_parcelas == 0:
        texto_intro = "Não identificamos pendências nos detalhes fornecidos."
        detalhes_parcelas_str = ""
    elif num_parcelas == 1:
        texto_intro = "Identificamos a seguinte pendência em aberto conosco:"
    else:
        texto_intro = "Identificamos as seguintes pendências em aberto conosco:"

    # --- Monta detalhes das parcelas ---
    detalhes_parcelas_str = ""
    if num_parcelas > 0:
        for i, parcela in enumerate(lista_parcelas_info):
            loteamento = parcela.get('loteamento_nome', '[Loteamento N/D]')
            lote = parcela.get('lote', '[Ref N/D]') # Ref. original Lote/Número
            venc = parcela.get('vencimento', '[N/D]')
            valor = parcela.get('valor', '[N/D]') 
            detalhes_parcelas_str += f"{i+1}. *Loteamento {loteamento}*\n" 
            detalhes_parcelas_str += f"    Lote: {lote}\n" 
            detalhes_parcelas_str += f"    Venc: {venc}\n" 
            detalhes_parcelas_str += f"    Valor: *{valor}*\n\n" 

        # Remove o último \n\n extra
        if detalhes_parcelas_str.endswith('\n\n'):
             detalhes_parcelas_str = detalhes_parcelas_str[:-2]

    # --- Template da Mensagem ---
    mensagem = (
        f"Olá {nome_cliente}! 👋\n\n"
        f"{texto_intro}\n\n"
        f"{detalhes_parcelas_str}\n\n"
        f"*Valor Total: {valor_total_fmt}*\n\n"
        f"👇 Como regularizar ou tirar dúvidas:\n"
        f"1️⃣ Responda esta mensagem.\n"
        f"2️⃣ Ligue para: [Seu Número de Telefone]\n\n" 
        f"Caso o pagamento já tenha sido efetuado, por favor, desconsidere esta mensagem.\n\n"
        f"Atenciosamente,\n"
        f"{nome_empresa}"
    )

    # --- Limpeza Final ---
    linhas = mensagem.splitlines()
    linhas_filtradas = []
    if linhas:
        linhas_filtradas.append(linhas[0])
        for i in range(1, len(linhas)):
            if linhas[i].strip() or linhas[i-1].strip():
                linhas_filtradas.append(linhas[i])
    mensagem = '\n'.join(linhas_filtradas)

    return mensagem