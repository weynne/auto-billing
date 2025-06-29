# modules/leitor_planilha.py
import pandas as pd
import os
import logging
from datetime import datetime
import phonenumbers
from config import MAPEAMENTO_LOTEAMENTO

logger = logging.getLogger(__name__)

def _formatar_telefone_br(tel_str):
    if not isinstance(tel_str, str) or not tel_str.strip(): return None
    try:
        phone_number = phonenumbers.parse(tel_str, "BR")
        if phonenumbers.is_valid_number(phone_number):
            return phonenumbers.format_number(phone_number, phonenumbers.PhoneNumberFormat.E164)
        return None
    except phonenumbers.phonenumberutil.NumberParseException:
        return None

def _formatar_nome_proprio(nome_str):
    if not isinstance(nome_str, str): return nome_str
    nome_limpo = nome_str.split(' - ', 1)[-1].strip()
    nome_formatado = nome_limpo.title()
    for prep in {' Da ', ' De ', ' Do ', ' Dos ', ' Das '}:
        if prep in nome_formatado: nome_formatado = nome_formatado.replace(prep, prep.lower())
    return nome_formatado

def _converter_valor_para_float(valor_str):
    if pd.isna(valor_str): return None
    try:
        s = str(valor_str).strip()
        s = ''.join(filter(lambda char: char in '0123456789,.', s))
        if ',' in s and '.' in s: s = s.replace('.', '')
        s = s.replace(',', '.')
        return float(s)
    except (ValueError, TypeError):
        return None

def carregar_inadimplentes(arquivo_input, config):
    try:
        df = pd.read_excel(arquivo_input, dtype=str)
        df.dropna(how='all', inplace=True)
    except Exception:
        logger.exception(f"Erro crítico ao ler o arquivo Excel.")
        return None

    for col_key in ['col_nome', 'col_telefone']:
        if config.get(col_key) not in df.columns:
            logger.error(f"Coluna obrigatória '{config.get(col_key)}' não encontrada.")
            return None
    
    if config.get('col_nome') in df.columns: df[config['col_nome']] = df[config['col_nome']].apply(_formatar_nome_proprio)
    if config.get('col_telefone') in df.columns: df[config['col_telefone']] = df[config['col_telefone']].apply(_formatar_telefone_br)
    if config.get('col_vencimento') in df.columns: df[config['col_vencimento']] = pd.to_datetime(df[config['col_vencimento']], errors='coerce', dayfirst=True)
    if config.get('col_valor') in df.columns: df[config['col_valor']] = df[config['col_valor']].apply(_converter_valor_para_float)
    
    col_loteamento = config.get('col_loteamento')
    if col_loteamento and col_loteamento in df.columns:
        df['cod_loteamento'] = df[col_loteamento].astype(str).str.split('/').str[0].str.strip()
        df['Nome Loteamento'] = df['cod_loteamento'].map(MAPEAMENTO_LOTEAMENTO).fillna(df['cod_loteamento'])
    else:
        df['Nome Loteamento'] = 'Não especificado'
        
    return df

def gerar_dados_dashboard_aging(df, config, data_base=None, loteamento_selecionado="Todos"):
    col_vencimento = config.get('col_vencimento')
    col_valor = config.get('col_valor')

    if not all([col_vencimento, col_valor]) or not all(c in df.columns for c in [col_vencimento, col_valor]):
        logging.error("Colunas essenciais (vencimento, valor) não encontradas.")
        return None

    df_dash = df.copy()
    
    if loteamento_selecionado != "Todos":
        df_dash = df_dash[df_dash['Nome Loteamento'] == loteamento_selecionado]

    data_base = pd.to_datetime(data_base if data_base else datetime.now().date())
    df_dash[col_vencimento] = pd.to_datetime(df_dash[col_vencimento], errors='coerce')
    df_vencidas = df_dash[df_dash[col_vencimento] < data_base].copy()
    
    if df_vencidas.empty:
        return {'kpis': {'total_vencido': 0}, 'dados_donut': pd.DataFrame()}

    df_vencidas['dias_atraso'] = (data_base - df_vencidas[col_vencimento]).dt.days
    
    bins = [-1, 30, 60, 90, 120, 180, 360, 720, float('inf')]
    labels = ['Até 30 dias', 'Vencido +30', 'Vencido +60', 'Vencido +90', 'Vencido +120', 'Vencido +180', 'Vencido +360', 'Vencido +720']
    df_vencidas['Faixa de Atraso'] = pd.cut(df_vencidas['dias_atraso'], bins=bins, labels=labels, right=True)
    
    total_vencido = df_vencidas[col_valor].sum()
    
    dados_donut = df_vencidas.groupby('Faixa de Atraso', observed=True)[col_valor].sum().reset_index()
    dados_donut.columns = ['Faixa de Atraso', 'Valor']
    
    if total_vencido > 0:
        dados_donut['Percentual'] = (dados_donut['Valor'] / total_vencido) * 100
    else:
        dados_donut['Percentual'] = 0

    return {
        'kpis': {'total_vencido': total_vencido},
        'dados_donut': dados_donut.sort_values(by='Faixa de Atraso').reset_index(drop=True)
    }