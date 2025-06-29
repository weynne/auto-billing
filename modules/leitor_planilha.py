# modules/leitor_planilha.py
import pandas as pd
import os
import logging
import phonenumbers 

logger = logging.getLogger(__name__)

def _formatar_telefone_br(tel_str):

    if not isinstance(tel_str, str) or not tel_str.strip():
        return None
    try:
        phone_number = phonenumbers.parse(tel_str, "BR")
        if phonenumbers.is_valid_number(phone_number):
            return phonenumbers.format_number(phone_number, phonenumbers.PhoneNumberFormat.E164)
        else:
            logger.warning(f"Número de telefone '{tel_str}' é considerado inválido.")
            return None
    except phonenumbers.phonenumberutil.NumberParseException as e:
        logger.warning(f"Não foi possível analisar o telefone '{tel_str}': {e}")
        return None

def _formatar_nome_proprio(nome_str):

    if not isinstance(nome_str, str):
        return nome_str
    
    nome_limpo = nome_str.split(' - ', 1)[-1].strip()
    
    nome_formatado = nome_limpo.title()
    preposicoes = {' Da ', ' De ', ' Do ', ' Dos ', ' Das '}
    for prep in preposicoes:
        if prep in nome_formatado:
            nome_formatado = nome_formatado.replace(prep, prep.lower())
    return nome_formatado

def _converter_valor_para_float(valor_str):
    if pd.isna(valor_str):
        return None
    try:
        s = str(valor_str).strip()
        s = ''.join(filter(lambda char: char in '0123456789,.', s))

        if ',' in s and '.' in s:
            s = s.replace('.', '') 

        s = s.replace(',', '.')
        
        return float(s)
    except (ValueError, TypeError):
        logger.warning(f"Não foi possível converter o valor '{valor_str}' para numérico.")
        return None


def carregar_inadimplentes(arquivo_input, config):

    nome_arquivo = "Streamlit Upload"
    if isinstance(arquivo_input, str):
        nome_arquivo = os.path.basename(arquivo_input)
    elif hasattr(arquivo_input, 'name'):
        nome_arquivo = arquivo_input.name
    
    logger.info(f"Lendo e processando arquivo: {nome_arquivo}")

    try:

        df = pd.read_excel(arquivo_input, dtype=str)

        df.dropna(how='all', inplace=True)
    except Exception as e:
        logger.exception(f"Erro crítico ao ler o arquivo Excel '{nome_arquivo}'")
        return None

    colunas_necessarias = { 'col_nome', 'col_telefone' }
    for col_key in colunas_necessarias:
        if config[col_key] not in df.columns:
            logger.error(f"Coluna obrigatória '{config[col_key]}' (de '{col_key}') não encontrada no arquivo.")
            return None

    logger.info(f"Arquivo '{nome_arquivo}' lido com {len(df)} linhas. Iniciando limpeza e formatação...")

    df[config['col_nome']] = df[config['col_nome']].apply(_formatar_nome_proprio)
    df[config['col_telefone']] = df[config['col_telefone']].apply(_formatar_telefone_br)

    if config['col_vencimento'] in df.columns:
        df[config['col_vencimento']] = pd.to_datetime(df[config['col_vencimento']], errors='coerce', dayfirst=True)
        nulos_venc = df[config['col_vencimento']].isnull().sum()
        if nulos_venc > 0:
            logger.warning(f"{nulos_venc} datas de vencimento inválidas foram convertidas para Nulo (NaT).")

    if config['col_valor'] in df.columns:
        df[config['col_valor']] = df[config['col_valor']].apply(_converter_valor_para_float)
        nulos_valor = df[config['col_valor']].isnull().sum()
        if nulos_valor > 0:
            logger.warning(f"{nulos_valor} valores monetários inválidos foram convertidos para Nulo.")
            
    logger.info(f"Pré-processamento de '{nome_arquivo}' concluído.")
    return df