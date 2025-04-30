# -*- coding: utf-8 -*-
# modules/leitor_planilha.py

import pandas as pd
import os
from dotenv import load_dotenv
import logging

# Carrega variáveis de ambiente
load_dotenv()

# Pega os nomes das colunas esperadas do .env
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')
COLUNA_VALOR = os.getenv('COLUNA_VALOR')
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO')

# Obtém logger
logger = logging.getLogger(__name__)

# --- Função formatar_telefone_br (sem alterações) ---
def formatar_telefone_br(tel_original):
    # ... (código da função igual ao anterior) ...
    if not isinstance(tel_original, str) or not tel_original.strip():
         return tel_original
    tel_digits = ''.join(filter(str.isdigit, tel_original))
    if len(tel_digits) == 10 or len(tel_digits) == 11:
        logger.debug(f"Formatando tel {tel_digits} (len {len(tel_digits)}) para +55...")
        return "+55" + tel_digits
    elif (len(tel_digits) == 12 or len(tel_digits) == 13) and tel_digits.startswith('55'):
        logger.debug(f"Formatando tel {tel_digits} (len {len(tel_digits)}, starts '55') para +...")
        return "+" + tel_digits
    elif isinstance(tel_original, str) and tel_original.startswith('+55') and (len(tel_digits) == 12 or len(tel_digits) == 13):
         logger.debug(f"Telefone '{tel_original}' já parece estar no formato +55. Mantendo.")
         return tel_original
    else:
        logger.warning(f"Telefone '{tel_original}' (dígitos: '{tel_digits}') com formato/comprimento inesperado. Não foi possível formatar para +55.")
        return tel_digits
# --------------------------------------------------

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
        logger.exception(f"Erro Crítico ao ler o arquivo Excel '{os.path.basename(caminho_arquivo_planilha)}'")
        logger.error("Verifique se é um Excel válido e se 'openpyxl'/'xlrd' está instalado.")
        return None

    # --- Validação de Colunas Essenciais ---
    colunas_esperadas = {
        'Nome': COLUNA_NOME, 'Telefone': COLUNA_TELEFONE,
        'Valor': COLUNA_VALOR, 'Vencimento': COLUNA_VENCIMENTO
    }
    # *** BLOCO RESTAURADO ABAIXO ***
    colunas_faltantes_nomes = []
    colunas_faltantes_keys = []
    colunas_presentes = []

    logger.info("Verificando colunas essenciais (definidas no .env):")
    for key, col_name in colunas_esperadas.items():
        if col_name:
            logger.debug(f"  - Esperando '{key}' na coluna: '{col_name}'")
            if col_name in df.columns:
                colunas_presentes.append(col_name)
            else:
                colunas_faltantes_nomes.append(col_name)
                colunas_faltantes_keys.append(key)
        else:
            # Permite que colunas como VALOR e VENCIMENTO sejam opcionais no .env
            # Mas NOME e TELEFONE são validados no script principal
            if key in ['Nome', 'Telefone']:
                 logger.warning(f"  - Variável de ambiente OBRIGATÓRIA para '{key}' não definida no .env (Ex: COLUNA_{key.upper()}=...)")
                 # Adiciona à lista de faltantes para o erro ser acionado abaixo
                 colunas_faltantes_nomes.append(f"Variável COLUNA_{key.upper()} não definida no .env")
                 colunas_faltantes_keys.append(key)
            else:
                 logger.info(f"  - Variável de ambiente OPCIONAL para '{key}' não definida no .env (Ex: COLUNA_{key.upper()}). Coluna não será processada.")
    # *** FIM DO BLOCO RESTAURADO ***

    # Agora esta verificação funcionará, pois as listas foram populadas
    if colunas_faltantes_nomes: # Verifica se a lista não está vazia
        logger.error("Erro Crítico: Colunas essenciais NÃO encontradas ou variáveis .env ausentes:")
        for i in range(len(colunas_faltantes_nomes)):
             logger.error(f"  - {colunas_faltantes_keys[i]}: Esperava coluna '{colunas_faltantes_nomes[i]}' ou variável definida")
        logger.error(f"Colunas encontradas na planilha: {df.columns.to_list()}")
        logger.error("Verifique o .env e os cabeçalhos da planilha.")
        return None
    # Verifica se pelo menos as colunas obrigatórias (Nome, Telefone) foram encontradas
    if not COLUNA_NOME in colunas_presentes or not COLUNA_TELEFONE in colunas_presentes:
         logger.error("Erro Crítico: Colunas de Nome ou Telefone não encontradas na planilha (verificar nomes no .env e cabeçalhos).")
         return None


    logger.info("Colunas essenciais encontradas/configuradas.")
    logger.info(f"Iniciando processamento e limpeza de {len(df)} registros...")

    # --- Limpeza e Conversões (sem alterações nesta parte) ---
    logger.debug("Iniciando limpeza/conversão...")

    # Telefone: Limpeza inicial + Formatação para +55
    if COLUNA_TELEFONE and COLUNA_TELEFONE in df.columns:
        logger.debug(f"Tentando limpar e formatar coluna '{COLUNA_TELEFONE}'...")
        try:
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].apply(
                lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if pd.notna(x) else ''
            ).astype(str)
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].apply(formatar_telefone_br)
            logger.debug(f"Coluna '{COLUNA_TELEFONE}' processada para formato E.164 (+55 quando aplicável).")
        except Exception as e:
            logger.exception(f"Erro ao limpar/formatar coluna de telefone '{COLUNA_TELEFONE}'")
            return None

    # Data de Vencimento
    if COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in df.columns:
        logger.debug(f"Tentando converter coluna '{COLUNA_VENCIMENTO}' para data...")
        try:
            df[COLUNA_VENCIMENTO] = pd.to_datetime(df[COLUNA_VENCIMENTO], errors='coerce')
            nat_count = df[COLUNA_VENCIMENTO].isnull().sum()
            if nat_count > 0:
                logger.warning(f"{nat_count} valores na coluna '{COLUNA_VENCIMENTO}' não puderam ser convertidos para data (viraram NaT).")
            logger.debug(f"Coluna '{COLUNA_VENCIMENTO}' convertida para data (erros viraram NaT).")
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de vencimento '{COLUNA_VENCIMENTO}'")
            # Pode ser não crítico se a data não for essencial, mas por segurança vamos parar
            return None

    # Valor
    if COLUNA_VALOR and COLUNA_VALOR in df.columns:
        logger.debug(f"Tentando converter coluna '{COLUNA_VALOR}' para numérico...")
        try:
            original_dtype_is_numeric = pd.api.types.is_numeric_dtype(df[COLUNA_VALOR])
            if not original_dtype_is_numeric:
                 logger.debug(f"Coluna '{COLUNA_VALOR}' não é numérica, tentando limpar string...")
                 df[COLUNA_VALOR] = df[COLUNA_VALOR].astype(str).str.replace(r'[R$\s.]', '', regex=True).str.replace(',', '.', regex=False)

            df[COLUNA_VALOR] = pd.to_numeric(df[COLUNA_VALOR], errors='coerce')
            nan_count = df[COLUNA_VALOR].isnull().sum()
            if nan_count > 0:
                logger.warning(f"{nan_count} valores na coluna '{COLUNA_VALOR}' não puderam ser convertidos para número (viraram NaN).")
            logger.debug(f"Coluna '{COLUNA_VALOR}' convertida para numérico (erros viraram NaN).")
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de valor '{COLUNA_VALOR}'")
            # Pode ser não crítico, mas paramos por segurança
            return None

    logger.info("--- Leitura e pré-processamento do arquivo Excel concluídos ---")
    return df