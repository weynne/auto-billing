#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# gerador_mensagens_cobranca.py

import time
import os
from dotenv import load_dotenv
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import logging # Importado logging

# --- Configuração do Logging ---
# Formato detalhado para o arquivo de log
log_formatter_detalhado = logging.Formatter(
    '%(asctime)s - %(levelname)s - [%(module)s:%(lineno)d] - %(message)s', # Adiciona módulo e linha
    datefmt='%Y-%m-%d %H:%M:%S'
)
# Formato simples para a saída do console (apenas a mensagem)
log_formatter_console = logging.Formatter('%(message)s')

# Configura o logger raiz
logger_raiz = logging.getLogger()
logger_raiz.setLevel(logging.INFO) # Define o nível geral (ex: INFO, DEBUG)

# Remove handlers existentes para evitar duplicação se o script for re-executado no mesmo processo
# (Útil em alguns cenários de teste ou execução repetida)
if logger_raiz.hasHandlers():
    logger_raiz.handlers.clear()

# Handler para CONSOLE (saída padrão - como o print)
console_handler = logging.StreamHandler() # Envia para stderr por padrão (pode usar sys.stdout)
console_handler.setFormatter(log_formatter_console)
# Opcional: Definir nível específico para o console, se diferente do logger raiz
# console_handler.setLevel(logging.INFO)

# Handler para ARQUIVO (guarda log detalhado)
nome_arquivo_log = "processamento_cobrancas.log"
try:
    file_handler = logging.FileHandler(nome_arquivo_log, mode='a', encoding='utf-8') # 'a' para append
    file_handler.setFormatter(log_formatter_detalhado)
    # Opcional: Definir nível específico para o arquivo, ex: DEBUG
    # file_handler.setLevel(logging.DEBUG)

    # Adiciona os handlers ao logger raiz
    logger_raiz.addHandler(console_handler)
    logger_raiz.addHandler(file_handler)

    # Mensagem inicial para confirmar configuração (será logada nos dois destinos)
    logging.info(f"--- Logging iniciado. Console: limpo | Arquivo: '{nome_arquivo_log}' (detalhado) ---")

except IOError as e:
    # Se não conseguir criar o arquivo de log, avisa no console e continua sem log em arquivo
    logger_raiz.addHandler(console_handler) # Garante que o console funcione
    logging.error(f"ERRO CRÍTICO: Não foi possível criar/abrir o arquivo de log '{nome_arquivo_log}'. {e}")
    logging.error("O script continuará, mas logs detalhados NÃO serão salvos no arquivo.")

# ------------------------------

# Carrega os módulos locais da aplicação
try:
    from modules import leitor_planilha
    from modules import construtor_mensagem
    from modules import arquivo_txt_sender
    logging.info("Módulos locais (leitor, construtor, sender) importados.") # Mais conciso
except ImportError as e:
    logging.exception("Erro Crítico: Falha ao importar módulos da pasta 'modules'.")
    logging.error("Verifique a estrutura da pasta 'modules' e seus arquivos.")
    exit()

# Carrega variáveis de ambiente do arquivo .env
if load_dotenv():
    logging.info("Arquivo .env carregado.")
else:
    logging.warning("AVISO: Arquivo .env não encontrado.")

# --- Configurações ---
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')

try:
    DELAY_ENTRE_MENSAGENS_SEGUNDOS = int(os.getenv('DELAY_SEGUNDOS', 1))
    logging.info(f"Delay entre mensagens: {DELAY_ENTRE_MENSAGENS_SEGUNDOS}s.")
except ValueError:
    logging.warning("AVISO: Valor inválido para DELAY_SEGUNDOS no .env. Usando padrão: 1s.")
    DELAY_ENTRE_MENSAGENS_SEGUNDOS = 1

def selecionar_arquivo_planilha():
    """
    Abre janela gráfica para selecionar o arquivo da planilha Excel tratada.
    Returns: str or None
    """
    logging.info("\nAbrindo janela para seleção da planilha...") # Adiciona linha extra para espaçar
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha Excel TRATADA",
        filetypes=[
            ("Arquivos Excel", "*.xlsx *.xls"),
            ("Arquivos CSV (Fallback)", "*.csv"),
            ("Todos os arquivos", "*.*")
        ]
    )
    root.destroy()

    if caminho_arquivo:
        logging.info(f"Arquivo selecionado: {os.path.basename(caminho_arquivo)}") # Loga só o nome base
        # Log detalhado no arquivo
        logger_raiz.handlers[1].handle(logging.LogRecord(
            name=logger_raiz.name, level=logging.DEBUG, pathname=None, lineno=0,
            msg=f"Caminho completo selecionado: {caminho_arquivo}", args=[], exc_info=None, func=''))
        return caminho_arquivo
    else:
        logging.warning("Nenhum arquivo foi selecionado.")
        return None

def processar_cobrancas():
    """
    Função principal para orquestrar a leitura, construção e salvamento das mensagens.
    """
    logging.info("\n--- Iniciando Processo de Geração de Arquivos de Cobrança (.txt) ---")

    # Verifica colunas essenciais
    if not COLUNA_NOME or not COLUNA_TELEFONE:
        logging.error("ERRO CRÍTICO: COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env.")
        return

    # 1. Selecionar planilha
    caminho_planilha_selecionada = selecionar_arquivo_planilha()
    if not caminho_planilha_selecionada:
        logging.info("Processo cancelado. Encerrando.")
        return

    # 2. Carregar dados
    logging.info(f"\n--- Carregando e Processando Planilha: {os.path.basename(caminho_planilha_selecionada)} ---")
    # Os logs internos do leitor aparecerão aqui (limpos no console, detalhados no arquivo)
    dados_inadimplentes = leitor_planilha.carregar_inadimplentes(caminho_planilha_selecionada)

    if dados_inadimplentes is None:
        logging.error("Falha crítica ao carregar dados. Verifique logs anteriores. Encerrando.")
        return
    if dados_inadimplentes.empty:
        logging.info("Planilha vazia ou sem dados válidos. Nenhum registro para processar.")
        return

    # 3. Processamento dos registros
    total_registros = len(dados_inadimplentes)
    salvos_sucesso = 0
    falhas = 0
    logging.info(f"\n--- Iniciando processamento de {total_registros} registros ---")

    for indice, cliente in dados_inadimplentes.iterrows():
        # Usar um log INFO para o separador no console
        logging.info("-" * 20)
        # E um log DEBUG (que só vai pro arquivo por padrão) para indicar o processamento
        logging.debug(f"Processando linha DataFrame índice {indice}...")

        # Log INFO para o console
        logging.info(f"Processando Registro {indice + 1}/{total_registros}...")

        nome_cliente = cliente.get(COLUNA_NOME, "Nome Ausente")
        telefone = cliente.get(COLUNA_TELEFONE, "")

        # Validação do telefone
        if not isinstance(telefone, str) or not (10 <= len(telefone) <= 11):
            # Log WARNING para console e arquivo
            logging.warning(f"AVISO: Telefone inválido/ausente para '{nome_cliente}' (Valor: '{telefone}'). Pulando.")
            falhas += 1
            continue

        # Logs INFO para console
        logging.info(f"  - Cliente: {nome_cliente}")
        logging.info(f"  - Telefone (limpo): {telefone}")

        # 4. Construir mensagem
        try:
            # Logs internos do construtor (se houver) seguirão a mesma formatação
            mensagem = construtor_mensagem.criar_mensagem_cobranca(cliente)
            logging.debug(f"Mensagem construída para {nome_cliente}") # Só no arquivo
        except Exception as e_msg:
            # Log ERROR para console e arquivo
            logging.error(f"ERRO: Falha ao construir mensagem para '{nome_cliente}': {e_msg}. Pulando.")
            # Log mais detalhado com traceback para o arquivo
            logging.exception(f"Detalhe da exceção ao construir msg para {nome_cliente}:")
            falhas += 1
            continue

        # 5. Salvar mensagem em .txt
        # Log INFO para console
        logging.info(f"  - Tentando salvar mensagem em arquivo .txt...")
        try:
            # Logs internos do sender seguirão a formatação (INFO dele vai pro console limpo)
            sucesso_salvar = arquivo_txt_sender.salvar_mensagem_em_txt(telefone, nome_cliente, mensagem)
        except Exception as e_save:
            # Log ERROR para console e arquivo
            logging.error(f"ERRO: Falha inesperada ao tentar salvar arquivo para '{nome_cliente}': {e_save}")
            # Log mais detalhado com traceback para o arquivo
            logging.exception(f"Detalhe da exceção ao salvar arquivo para {nome_cliente}:")
            sucesso_salvar = False

        if sucesso_salvar:
            salvos_sucesso += 1
            # A confirmação de sucesso já é logada pelo sender, não precisa logar de novo aqui.
            # Poderia logar um DEBUG aqui se quisesse algo extra só no arquivo.
            logging.debug(f"Marcação de sucesso para {nome_cliente}")
        else:
            falhas += 1
            # O erro já foi logado acima ou dentro do sender
            logging.info(f"  - Falha ao salvar a mensagem.") # Informa no console

        # 6. Pausa
        if total_registros > 1 and indice < total_registros - 1:
             # Log INFO para console
            logging.info(f"  - Aguardando {DELAY_ENTRE_MENSAGENS_SEGUNDOS} segundo(s)...")
            time.sleep(DELAY_ENTRE_MENSAGENS_SEGUNDOS)

    # 7. Resumo final
    logging.info("\n" + "="*40)
    logging.info("--- Processo Concluído ---")
    logging.info(f"Planilha processada: {os.path.basename(caminho_planilha_selecionada)}")
    logging.info(f"Total de registros na planilha: {total_registros}")
    logging.info(f"Mensagens salvas com sucesso em arquivos .txt: {salvos_sucesso}")
    logging.info(f"Registros com falha ou pulados: {falhas}")
    if salvos_sucesso > 0:
        try:
            dir_saida = arquivo_txt_sender.DIRETORIO_SAIDA
            logging.info(f"Verifique os arquivos na pasta: '{dir_saida}'")
        except AttributeError:
             logging.warning("AVISO: Não foi possível determinar o diretório de saída do 'arquivo_txt_sender'.")
    logging.info("="*40)

# Ponto de entrada principal
if __name__ == "__main__":
    processar_cobrancas()