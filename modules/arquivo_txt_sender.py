# -*- coding: utf-8 -*-
# modules/arquivo_txt_sender.py

import os
import re # Para limpar nomes de arquivos
import logging
from datetime import datetime

# Configura logging básico
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Diretório onde os arquivos .txt serão salvos
# Será criado na pasta raiz do projeto se não existir
DIRETORIO_SAIDA = "mensagens_txt_geradas"

def limpar_nome_arquivo(nome):
    """Remove caracteres inválidos para nomes de arquivo."""
    # Remove caracteres não alfanuméricos, exceto espaço, hífen e underscore
    nome = re.sub(r'[^\w\s-]', '', nome)
    # Substitui espaços por underscore
    nome = re.sub(r'\s+', '_', nome)
    # Limita o comprimento para evitar nomes de arquivo muito longos
    return nome[:50] # Limita a 50 caracteres

def salvar_mensagem_em_txt(telefone_destino, nome_cliente, mensagem):
    """
    Salva a mensagem de cobrança em um arquivo .txt individual.

    Args:
        telefone_destino (str): Número de telefone do destinatário (usado no nome do arquivo).
        nome_cliente (str): Nome do cliente (usado no nome do arquivo).
        mensagem (str): O texto completo da mensagem a ser salva.

    Returns:
        bool: True se o arquivo foi salvo com sucesso, False caso contrário.
    """
    if not telefone_destino:
        telefone_destino = "telefone_na" # Fallback se telefone for inválido/ausente
    if not nome_cliente:
         nome_cliente = "cliente_desconhecido" # Fallback

    # Cria o diretório de saída se não existir
    try:
        os.makedirs(DIRETORIO_SAIDA, exist_ok=True)
    except OSError as e:
        logging.error(f"Erro ao criar o diretório de saída '{DIRETORIO_SAIDA}': {e}")
        return False

    # Monta o nome do arquivo de forma segura
    nome_cliente_limpo = limpar_nome_arquivo(nome_cliente)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Formato: cobranca_5581999998888_Fulano_de_Tal_Teste_20250428_213000.txt
    nome_arquivo = f"cobranca_{telefone_destino}_{nome_cliente_limpo}_{timestamp}.txt"
    caminho_completo = os.path.join(DIRETORIO_SAIDA, nome_arquivo)

    # Salva a mensagem no arquivo
    try:
        with open(caminho_completo, 'w', encoding='utf-8') as f:
            f.write(f"Destinatário (Telefone): {telefone_destino}\n")
            f.write(f"Nome: {nome_cliente}\n")
            f.write("="*30 + "\n")
            f.write(mensagem)
        logging.info(f"Mensagem para {nome_cliente} ({telefone_destino}) salva em: {caminho_completo}")
        return True
    except IOError as e:
        logging.error(f"Erro ao salvar arquivo '{caminho_completo}': {e}")
        return False
    except Exception as e:
         logging.error(f"Erro inesperado ao salvar arquivo para {nome_cliente}: {e}")
         return False

# Bloco para teste rápido do módulo
if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO) # Garante logging no teste
    print("\n--- Testando arquivo_txt_sender.py ---")

    teste_tel = "5511987654321"
    teste_nome = "João da Silva Teste"
    teste_msg = "Esta é uma mensagem de teste para ser salva em arquivo.\nSegunda linha."

    print(f"\nTentando salvar mensagem para {teste_nome}...")
    sucesso = salvar_mensagem_em_txt(teste_tel, teste_nome, teste_msg)
    print(f"Resultado: {'Sucesso' if sucesso else 'Falha'}")
    if sucesso:
         print(f"Verifique a pasta '{DIRETORIO_SAIDA}' na raiz do seu projeto.")

    print("\n--- Teste com nome inválido para arquivo ---")
    teste_nome_invalido = "Maria D'Ávila & Souza / Teste"
    sucesso_invalido = salvar_mensagem_em_txt(teste_tel, teste_nome_invalido, teste_msg)
    print(f"Resultado (nome inválido): {'Sucesso' if sucesso_invalido else 'Falha'}")


    print("\n--- Teste concluído ---")