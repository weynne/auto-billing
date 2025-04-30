#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# gerador_mensagens_cobranca.py

import time
import os
from dotenv import load_dotenv
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import logging

# --- Configuração do Logging (mantém a configuração anterior com dois handlers) ---
log_formatter_detalhado = logging.Formatter(
    '%(asctime)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
log_formatter_console = logging.Formatter('%(message)s')
logger_raiz = logging.getLogger()
logger_raiz.setLevel(logging.INFO)
if logger_raiz.hasHandlers():
    logger_raiz.handlers.clear()
console_handler = logging.StreamHandler()
console_handler.setFormatter(log_formatter_console)
nome_arquivo_log = "processamento_cobrancas.log"
try:
    file_handler = logging.FileHandler(nome_arquivo_log, mode='a', encoding='utf-8')
    file_handler.setFormatter(log_formatter_detalhado)
    logger_raiz.addHandler(console_handler)
    logger_raiz.addHandler(file_handler)
    logging.info(f"--- Logging iniciado. Console: limpo | Arquivo: '{nome_arquivo_log}' (detalhado) ---")
except IOError as e:
    logger_raiz.addHandler(console_handler)
    logging.error(f"ERRO CRÍTICO: Não foi possível criar/abrir o arquivo de log '{nome_arquivo_log}'. {e}")
    logging.error("Logs detalhados NÃO serão salvos no arquivo.")
# ------------------------------

# --- Carregamento de Módulos e Env (igual a antes) ---
try:
    from modules import leitor_planilha
    from modules import construtor_mensagem
    from modules import arquivo_txt_sender
    logging.info("Módulos locais (leitor, construtor, sender) importados.")
except ImportError as e:
    logging.exception("Erro Crítico: Falha ao importar módulos da pasta 'modules'.")
    exit()

if load_dotenv():
    logging.info("Arquivo .env carregado.")
else:
    logging.warning("AVISO: Arquivo .env não encontrado.")

# --- Configurações (igual a antes) ---
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')
try:
    DELAY_ENTRE_MENSAGENS_SEGUNDOS = int(os.getenv('DELAY_SEGUNDOS', 1))
    logging.info(f"Delay entre mensagens: {DELAY_ENTRE_MENSAGENS_SEGUNDOS}s.")
except ValueError:
    logging.warning("AVISO: Valor inválido para DELAY_SEGUNDOS. Usando padrão: 1s.")
    DELAY_ENTRE_MENSAGENS_SEGUNDOS = 1

# --- Função selecionar_arquivo_planilha (igual a antes) ---
def selecionar_arquivo_planilha():
    # ... (código igual ao da versão anterior) ...
    logging.info("\nAbrindo janela para seleção da planilha...")
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha Excel TRATADA",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Arquivos CSV", "*.csv"), ("Todos", "*.*")]
    )
    root.destroy()
    if caminho_arquivo:
        logging.info(f"Arquivo selecionado: {os.path.basename(caminho_arquivo)}")
        # Log detalhado opcional
        # logger_raiz.handlers[1].handle(...) # Log caminho completo se necessário
        return caminho_arquivo
    else:
        logging.warning("Nenhum arquivo foi selecionado.")
        return None

# --- Função processar_cobrancas (com validação de telefone AJUSTADA) ---
def processar_cobrancas():
    """ Orquestra leitura, construção e salvamento das mensagens. """
    logging.info("\n--- Iniciando Processo de Geração de Arquivos de Cobrança (.txt) ---")

    if not COLUNA_NOME or not COLUNA_TELEFONE:
        logging.error("ERRO CRÍTICO: COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env.")
        return

    caminho_planilha_selecionada = selecionar_arquivo_planilha()
    if not caminho_planilha_selecionada:
        logging.info("Processo cancelado. Encerrando.")
        return

    logging.info(f"\n--- Carregando e Processando Planilha: {os.path.basename(caminho_planilha_selecionada)} ---")
    dados_inadimplentes = leitor_planilha.carregar_inadimplentes(caminho_planilha_selecionada)

    if dados_inadimplentes is None or dados_inadimplentes.empty:
        logging.error("Falha ao carregar dados ou planilha vazia. Encerrando.")
        return

    total_registros = len(dados_inadimplentes)
    salvos_sucesso = 0
    falhas = 0
    logging.info(f"\n--- Iniciando processamento de {total_registros} registros ---")

    for indice, cliente in dados_inadimplentes.iterrows():
        logging.info("-" * 20)
        logging.debug(f"Processando linha DataFrame índice {indice}...")
        logging.info(f"Processando Registro {indice + 1}/{total_registros}...")

        nome_cliente = cliente.get(COLUNA_NOME, "Nome Ausente")
        # Pega o telefone já formatado (ou não) pelo leitor
        telefone = cliente.get(COLUNA_TELEFONE, "")

        # --- VALIDAÇÃO AJUSTADA ---
        # Espera formato E.164: +55 seguido de 10 ou 11 dígitos
        # Comprimento total deve ser 13 (+55 + 10) ou 14 (+55 + 11)
        if not isinstance(telefone, str) or not telefone.startswith("+55") or not (len(telefone) == 13 or len(telefone) == 14):
            logging.warning(f"AVISO: Telefone inválido ou fora do padrão E.164 (+55...) para '{nome_cliente}' (Valor lido: '{telefone}'). Pulando.")
            falhas += 1
            continue
        # --------------------------

        logging.info(f"  - Cliente: {nome_cliente}")
        logging.info(f"  - Telefone (E.164): {telefone}") # Mostra o telefone formatado

        # Construir mensagem (igual a antes)
        try:
            mensagem = construtor_mensagem.criar_mensagem_cobranca(cliente)
            logging.debug(f"Mensagem construída para {nome_cliente}")
        except Exception as e_msg:
            logging.error(f"ERRO: Falha ao construir mensagem para '{nome_cliente}': {e_msg}. Pulando.")
            logging.exception(f"Detalhe da exceção:") # Log traceback no arquivo
            falhas += 1
            continue

        # Salvar mensagem em .txt (igual a antes)
        logging.info(f"  - Tentando salvar mensagem em arquivo .txt...")
        try:
            # Passa o telefone formatado (+55...) para o sender
            sucesso_salvar = arquivo_txt_sender.salvar_mensagem_em_txt(telefone, nome_cliente, mensagem)
        except Exception as e_save:
            logging.error(f"ERRO: Falha inesperada ao tentar salvar arquivo para '{nome_cliente}': {e_save}")
            logging.exception(f"Detalhe da exceção:") # Log traceback no arquivo
            sucesso_salvar = False

        if sucesso_salvar:
            salvos_sucesso += 1
            logging.debug(f"Marcação de sucesso para {nome_cliente}")
        else:
            falhas += 1
            logging.info(f"  - Falha ao salvar a mensagem.")

        # Pausa (igual a antes)
        if total_registros > 1 and indice < total_registros - 1:
            logging.info(f"  - Aguardando {DELAY_ENTRE_MENSAGENS_SEGUNDOS} segundo(s)...")
            time.sleep(DELAY_ENTRE_MENSAGENS_SEGUNDOS)

    # --- Resumo final (igual a antes) ---
    logging.info("\n" + "="*40)
    logging.info("--- Processo Concluído ---")
    # ... (resto do resumo igual) ...
    logging.info(f"Planilha processada: {os.path.basename(caminho_planilha_selecionada)}")
    logging.info(f"Total de registros na planilha: {total_registros}")
    logging.info(f"Mensagens salvas com sucesso em arquivos .txt: {salvos_sucesso}")
    logging.info(f"Registros com falha ou pulados: {falhas}")
    if salvos_sucesso > 0:
        try:
            dir_saida = arquivo_txt_sender.DIRETORIO_SAIDA
            logging.info(f"Verifique os arquivos na pasta: '{dir_saida}'")
        except AttributeError:
             logging.warning("AVISO: Não foi possível determinar o diretório de saída.")
    logging.info("="*40)


# --- Ponto de entrada (igual a antes) ---
if __name__ == "__main__":
    processar_cobrancas()