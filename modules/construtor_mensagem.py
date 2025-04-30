# -*- coding: utf-8 -*-
# modules/construtor_mensagem.py

import os
import pandas as pd
from dotenv import load_dotenv
import locale # Para formatação de moeda, opcional mas recomendado

# Carrega variáveis de ambiente
load_dotenv()

# Pega os nomes das colunas do .env
COLUNA_NOME = os.getenv('COLUNA_NOME') # Ex: CLIENTE
COLUNA_VALOR = os.getenv('COLUNA_VALOR') # Ex: SALDO
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO') # Ex: VENCTO.

# Configura a localização para formatação de moeda brasileira (opcional)
# Pode ser necessário instalar o locale no sistema: sudo locale-gen pt_BR.UTF-8
# Ou pode falhar em alguns ambientes. Se falhar, a formatação manual abaixo funciona.
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    print("Locale 'pt_BR.UTF-8' configurado para formatação de moeda.")
    use_locale = True
except locale.Error:
    print("Aviso: Locale 'pt_BR.UTF-8' não encontrado. Usando formatação de moeda manual.")
    use_locale = False


def criar_mensagem_cobranca(dados_cliente):
    """
    Cria uma mensagem de cobrança personalizada para um cliente.

    Args:
        dados_cliente (pandas.Series): Uma linha do DataFrame processado
                                       pelo leitor_planilha.py.

    Returns:
        str: A mensagem formatada, pronta para ser enviada.
    """

    # Extrai os dados usando os nomes das colunas do .env
    # Usamos .get() com valor padrão para evitar erro se a coluna não existir por algum motivo
    nome = dados_cliente.get(COLUNA_NOME, "Cliente")
    valor_raw = dados_cliente.get(COLUNA_VALOR)
    vencimento_obj = dados_cliente.get(COLUNA_VENCIMENTO)

    # Formata o Valor (Saldo)
    valor_formatado = "[Valor Indisponível]" # Valor padrão em caso de erro
    if pd.notna(valor_raw): # Verifica se não é NaN (erro de conversão anterior)
        try:
            valor_float = float(valor_raw)
            if use_locale:
                 # Usa formatação de moeda local (ex: R$ 1.234,56)
                 valor_formatado = locale.currency(valor_float, grouping=True, symbol='R$')
            else:
                 # Formatação manual (ex: R$ 1.234,56)
                 valor_formatado = f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
             # Se ainda assim não for um número válido, mantém o padrão
             pass

    # Formata a Data de Vencimento
    vencimento_formatado = "[Data Indisponível]" # Valor padrão
    # Verifica se não é NaT (Not a Time - erro de conversão de data)
    if pd.notna(vencimento_obj) and hasattr(vencimento_obj, 'strftime'):
        try:
            vencimento_formatado = vencimento_obj.strftime('%d/%m/%Y') # Formato DD/MM/YYYY
        except ValueError:
             # Se for uma data inválida que passou pela conversão inicial
             pass

    # --- Monte sua Mensagem Aqui ---
    # Use f-strings para inserir as variáveis formatadas.
    # **Adapte o texto conforme sua necessidade.**
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

    return mensagem

# Bloco para teste rápido do módulo
if __name__ == '__main__':
    print("\n--- Testando construtor_mensagem.py ---")

    # Garante que as variáveis do .env foram carregadas para o teste
    if not all([COLUNA_NOME, COLUNA_VALOR, COLUNA_VENCIMENTO]):
         print("Erro no teste: Defina COLUNA_NOME, COLUNA_VALOR, COLUNA_VENCIMENTO no seu .env")
    else:
        # Dados de exemplo simulando uma linha do DataFrame
        dados_teste_ok = pd.Series({
            COLUNA_NOME: "Fulano de Tal Teste",
            COLUNA_VALOR: 1234.56,
            COLUNA_VENCIMENTO: pd.to_datetime("2025-03-15")
        })

        dados_teste_erro = pd.Series({
            COLUNA_NOME: "Ciclano Sem Dados",
            COLUNA_VALOR: None, # Simula erro na conversão de valor (NaN)
            COLUNA_VENCIMENTO: pd.NaT # Simula erro na conversão de data (NaT)
        })

        print("\n--- Mensagem Teste (Dados OK) ---")
        print(criar_mensagem_cobranca(dados_teste_ok))

        print("\n--- Mensagem Teste (Dados com Erro/Ausentes) ---")
        print(criar_mensagem_cobranca(dados_teste_erro))