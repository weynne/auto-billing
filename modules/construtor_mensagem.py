# modules/construtor_mensagem.py
import logging

logger = logging.getLogger(__name__)

def criar_mensagem_consolidada(nome_cliente, lista_parcelas_info, valor_total_fmt, nome_empresa, telefone_contato):
    logger.debug(f"Criando msg para {nome_cliente} ({nome_empresa}) com {len(lista_parcelas_info)} parcela(s).")

    num_parcelas = len(lista_parcelas_info)
    
    if num_parcelas == 1:
        texto_intro = "Identificamos a seguinte pendência em aberto conosco:"
    else:
        texto_intro = "Identificamos as seguintes pendências em aberto conosco:"

    detalhes_parcelas_str = ""
    for i, p in enumerate(lista_parcelas_info):
        detalhes_parcelas_str += (
            f"*{i+1}. Loteamento {p.get('loteamento_nome', 'N/D')}*\n"
            f"   • Lote/Ref.: {p.get('lote', 'N/D')}\n"
            f"   • Vencimento: {p.get('vencimento', 'N/D')}\n"
            f"   • Valor: {p.get('valor', 'N/D')}\n\n"
        )

    mensagem = (
        f"Olá, {nome_cliente}! 👋\n\n"
        f"{texto_intro}\n\n"
        f"{detalhes_parcelas_str.strip()}\n\n"
        f"Totalizando: *{valor_total_fmt}*\n\n"
        f"Para regularizar sua situação ou tirar dúvidas, por favor, escolha uma das opções abaixo:\n"
        f"1️⃣ *Responder esta mensagem*\n"
        f"2️⃣ *Ligar para:* {telefone_contato}\n\n"
        f"Caso o pagamento já tenha sido realizado, por favor, desconsidere este aviso.\n\n"
        f"Agradecemos a sua atenção.\n"
        f"Atenciosamente."
    )

    return mensagem