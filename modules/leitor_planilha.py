# modules/leitor_planilha.py
import pandas as pd
import os
from dotenv import load_dotenv
import logging

load_dotenv() 

COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')
COLUNA_VALOR = os.getenv('COLUNA_VALOR')
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO')
COLUNA_LOTE = os.getenv('COLUNA_LOTE')
COLUNA_LOTEAMENTO = os.getenv('COLUNA_LOTEAMENTO')

logger = logging.getLogger(__name__)

def formatar_telefone_br(tel_original):
    if not isinstance(tel_original, str) or not tel_original.strip(): return tel_original
    tel_digits = ''.join(filter(str.isdigit, tel_original))
    if len(tel_digits) == 10 or len(tel_digits) == 11: return "+55" + tel_digits
    elif (len(tel_digits) == 12 or len(tel_digits) == 13) and tel_digits.startswith('55'): return "+" + tel_digits
    elif isinstance(tel_original, str) and tel_original.startswith('+55') and (len(tel_digits) == 12 or len(tel_digits) == 13): return tel_original
    else: logger.warning(f"LEITOR_LOG: Telefone '{tel_original}' (dígitos: '{tel_digits}') com formato/comprimento inesperado."); return tel_original

def extrair_e_formatar_nome(nome_completo):
    nome_extraido = nome_completo
    if isinstance(nome_completo, str) and ' - ' in nome_completo:
        try:
            partes = nome_completo.split(' - ', 1); nome_extraido = partes[1].strip() if len(partes) > 1 else nome_completo
        except Exception as e: logger.warning(f"LEITOR_LOG: Erro ao extrair nome de '{nome_completo}': {e}.")
    if isinstance(nome_extraido, str):
        nome_formatado = nome_extraido.lower().title()
        for exc in {' Da ', ' De ', ' Do ', ' Dos ', ' Das '}:
            if exc in nome_formatado: nome_formatado = nome_formatado.replace(exc, exc.lower())
        return nome_formatado
    return str(nome_extraido)

def carregar_inadimplentes(arquivo_input):
    display_name = "arquivo_desconhecido.xlsx"
    if isinstance(arquivo_input, str):
        display_name = os.path.basename(arquivo_input)
        logger.info(f"LEITOR_LOG: Lendo arquivo Excel (caminho): {display_name}")
        if not os.path.exists(arquivo_input): logger.error(f"LEITOR_LOG: Arquivo não encontrado: '{arquivo_input}'"); return None
        file_to_read = arquivo_input
    else: 
        if hasattr(arquivo_input, 'name'): display_name = arquivo_input.name
        logger.info(f"LEITOR_LOG: Lendo arquivo Excel (objeto em memória): {display_name}")
        if hasattr(arquivo_input, 'seek') and callable(arquivo_input.seek):
            try: arquivo_input.seek(0)
            except Exception as e_seek: logger.warning(f"LEITOR_LOG: Não foi possível usar seek(0): {e_seek}")
        file_to_read = arquivo_input
    try:
        df = pd.read_excel(file_to_read, header=0, dtype=str)
        colunas_obrigatorias_para_linha = [col for col in [COLUNA_NOME, COLUNA_TELEFONE] if col]
        if not colunas_obrigatorias_para_linha:
            logger.error("LEITOR_LOG: ERRO CRÍTICO - COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env.")
            return None
        df.dropna(subset=colunas_obrigatorias_para_linha, how='all', inplace=True)
        logger.info(f"LEITOR_LOG: Arquivo '{display_name}' lido. {len(df)} linhas após remoção de vazias.")
        if df.empty: logger.warning(f"LEITOR_LOG: Planilha '{display_name}' vazia."); return df
    except Exception as e: logger.exception(f"LEITOR_LOG: Erro Crítico ao ler Excel '{display_name}'"); return None

    colunas_criticas = {'Nome': COLUNA_NOME, 'Telefone': COLUNA_TELEFONE}
    for key, nome_col in colunas_criticas.items():
        if not nome_col: logger.error(f"LEITOR_LOG: Variável para coluna OBRIGATÓRIA '{key}' não definida no .env."); return None
        if nome_col not in df.columns: logger.error(f"LEITOR_LOG: Coluna OBRIGATÓRIA '{nome_col}' (para '{key}') não encontrada. Colunas: {df.columns.to_list()}"); return None
    
    logger.info(f"LEITOR_LOG: Colunas essenciais verificadas. Limpando {len(df)} registros...")
    if COLUNA_NOME and COLUNA_NOME in df.columns: df[COLUNA_NOME] = df[COLUNA_NOME].apply(extrair_e_formatar_nome)
    if COLUNA_TELEFONE and COLUNA_TELEFONE in df.columns: df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].apply(formatar_telefone_br)
    if COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in df.columns:
        try:
            df[COLUNA_VENCIMENTO] = pd.to_datetime(df[COLUNA_VENCIMENTO], errors='coerce', dayfirst=True)
            if df[COLUNA_VENCIMENTO].isnull().any(): logger.warning(f"LEITOR_LOG: {df[COLUNA_VENCIMENTO].isnull().sum()} NaTs em '{COLUNA_VENCIMENTO}'.")
        except Exception: logger.exception(f"LEITOR_LOG: Erro ao converter '{COLUNA_VENCIMENTO}'.")
    if COLUNA_VALOR and COLUNA_VALOR in df.columns:
        try:
            col_valor_str = df[COLUNA_VALOR].astype(str).str.strip().str.replace(r'^R\$\s*', '', regex=True).str.replace(r'\.(?=\d{3},)', '', regex=True).str.replace(',', '.', regex=False).str.replace(r'\s+', '', regex=True)
            df[COLUNA_VALOR] = pd.to_numeric(col_valor_str, errors='coerce')
            if df[COLUNA_VALOR].isnull().any(): logger.warning(f"LEITOR_LOG: {df[COLUNA_VALOR].isnull().sum()} NaNs em '{COLUNA_VALOR}'.")
        except Exception: logger.exception(f"LEITOR_LOG: Erro ao converter '{COLUNA_VALOR}'.")
    
    logger.info(f"LEITOR_LOG: Pré-processamento de '{display_name}' concluído.")
    return df