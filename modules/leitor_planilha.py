# -*- coding: utf-8 -*-
# modules/leitor_planilha.py

import pandas as pd
import os
from dotenv import load_dotenv
import logging # <--- Importado logging

# Carrega variáveis de ambiente
load_dotenv()

# Pega os nomes das colunas esperadas do .env
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')
COLUNA_VALOR = os.getenv('COLUNA_VALOR')
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO')

# Obtém um logger específico para este módulo
# Isso permite controle mais granular do logging se necessário no futuro
logger = logging.getLogger(__name__)

def carregar_inadimplentes(caminho_arquivo_planilha):
    """
    Carrega e pré-processa dados da planilha Excel TRATADA.
    Args: caminho_arquivo_planilha (str)
    Returns: pandas.DataFrame or None
    """
    logger.info(f"--- Iniciando leitura do arquivo Excel: {os.path.basename(caminho_arquivo_planilha)} ---")
    if not caminho_arquivo_planilha or not os.path.exists(caminho_arquivo_planilha):
        logger.error(f"Arquivo da planilha não encontrado em '{caminho_arquivo_planilha}'")
        return None

    try:
        df = pd.read_excel(caminho_arquivo_planilha, header=0)
        logger.info(f"Arquivo Excel lido. {len(df)} linhas encontradas antes da validação.")
    except FileNotFoundError:
         logger.error(f"Arquivo não encontrado (verificado novamente) em '{caminho_arquivo_planilha}'")
         return None
    except Exception as e:
        logger.exception(f"Erro Crítico ao ler o arquivo Excel '{os.path.basename(caminho_arquivo_planilha)}'") # Captura traceback
        logger.error("Verifique se é um Excel válido e se 'openpyxl'/'xlrd' está instalado.")
        return None

    # --- Validação de Colunas Essenciais ---
    colunas_esperadas = {
        'Nome': COLUNA_NOME, 'Telefone': COLUNA_TELEFONE,
        'Valor': COLUNA_VALOR, 'Vencimento': COLUNA_VENCIMENTO
    }
    colunas_faltantes_nomes = []
    colunas_faltantes_keys = []
    colunas_presentes = []

    logger.info("Verificando colunas essenciais (definidas no .env):")
    for key, col_name in colunas_esperadas.items():
        if col_name:
            logger.debug(f"  - Esperando '{key}' na coluna: '{col_name}'") # Nível DEBUG
            if col_name in df.columns:
                colunas_presentes.append(col_name)
            else:
                colunas_faltantes_nomes.append(col_name)
                colunas_faltantes_keys.append(key)
        else:
            logger.warning(f"  - Variável de ambiente para '{key}' não definida no .env (Ex: COLUNA_{key.upper()})")

    if colunas_faltantes_nomes:
        logger.error("Erro Crítico: Colunas essenciais NÃO encontradas:")
        for i in range(len(colunas_faltantes_nomes)):
            logger.error(f"  - {colunas_faltantes_keys[i]}: Esperava '{colunas_faltantes_nomes[i]}'")
        logger.error(f"Colunas encontradas na planilha: {df.columns.to_list()}")
        logger.error("Verifique o .env e os cabeçalhos da planilha.")
        return None
    if not colunas_presentes:
        logger.error("Erro Crítico: Nenhuma coluna essencial encontrada/definida. Não é possível continuar.")
        return None

    logger.info("Todas as colunas essenciais definidas no .env foram encontradas.")
    logger.info(f"Iniciando processamento e limpeza de {len(df)} registros...")

    # --- Limpeza e Conversões ---
    logger.debug("Iniciando limpeza/conversão...") # Nível DEBUG

    # Telefone
    if COLUNA_TELEFONE and COLUNA_TELEFONE in df.columns:
        logger.debug(f"Tentando limpar coluna '{COLUNA_TELEFONE}'...") # Nível DEBUG
        try:
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].apply(lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if pd.notna(x) else '').astype(str)
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].str.replace(r'[^\d]', '', regex=True)
            logger.debug(f"Coluna '{COLUNA_TELEFONE}' limpa.") # Nível DEBUG
        except Exception as e:
            logger.exception(f"Erro ao limpar coluna de telefone '{COLUNA_TELEFONE}'") # Captura traceback
            return None

    # Data de Vencimento
    if COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in df.columns:
        logger.debug(f"Tentando converter coluna '{COLUNA_VENCIMENTO}' para data...") # Nível DEBUG
        try:
            df[COLUNA_VENCIMENTO] = pd.to_datetime(df[COLUNA_VENCIMENTO], errors='coerce')
            nat_count = df[COLUNA_VENCIMENTO].isnull().sum()
            if nat_count > 0:
                logger.warning(f"{nat_count} valores na coluna '{COLUNA_VENCIMENTO}' não puderam ser convertidos para data (viraram NaT).")
            logger.debug(f"Coluna '{COLUNA_VENCIMENTO}' convertida para data (erros viraram NaT).") # Nível DEBUG
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de vencimento '{COLUNA_VENCIMENTO}'") # Captura traceback
            return None

    # Valor
    if COLUNA_VALOR and COLUNA_VALOR in df.columns:
        logger.debug(f"Tentando converter coluna '{COLUNA_VALOR}' para numérico...") # Nível DEBUG
        try:
            original_dtype_is_numeric = pd.api.types.is_numeric_dtype(df[COLUNA_VALOR])
            if not original_dtype_is_numeric:
                 logger.debug(f"Coluna '{COLUNA_VALOR}' não é numérica, tentando limpar string...") # Nível DEBUG
                 df[COLUNA_VALOR] = df[COLUNA_VALOR].astype(str).str.replace(r'[R$\s.]', '', regex=True).str.replace(',', '.', regex=False)

            df[COLUNA_VALOR] = pd.to_numeric(df[COLUNA_VALOR], errors='coerce')
            nan_count = df[COLUNA_VALOR].isnull().sum()
            if nan_count > 0:
                logger.warning(f"{nan_count} valores na coluna '{COLUNA_VALOR}' não puderam ser convertidos para número (viraram NaN).")
            logger.debug(f"Coluna '{COLUNA_VALOR}' convertida para numérico (erros viraram NaN).") # Nível DEBUG
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de valor '{COLUNA_VALOR}'") # Captura traceback
            return None

    logger.info("--- Leitura e pré-processamento do arquivo Excel concluídos ---")
    return df