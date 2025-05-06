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
COLUNA_LOTE = os.getenv('LOTE')
COLUNA_LOTEAMENTO = os.getenv('COLUNA_LOTEAMENTO') # Ex: DUPLICATA

# Obtém logger
logger = logging.getLogger(__name__)

# --- Função formatar_telefone_br (sem mudanças) ---
def formatar_telefone_br(tel_original):
    if not isinstance(tel_original, str) or not tel_original.strip():
         return tel_original
    tel_digits = ''.join(filter(str.isdigit, tel_original))
    if len(tel_digits) == 10 or len(tel_digits) == 11: # Ex: 8199998888 ou 8133334444
        logger.debug(f"Formatando tel {tel_digits} (len {len(tel_digits)}) para +55...")
        return "+55" + tel_digits
    elif (len(tel_digits) == 12 or len(tel_digits) == 13) and tel_digits.startswith('55'): # Ex: 558199998888
        logger.debug(f"Formatando tel {tel_digits} (len {len(tel_digits)}, starts '55') para +...")
        return "+" + tel_digits
    elif isinstance(tel_original, str) and tel_original.startswith('+55') and (len(tel_digits) == 12 or len(tel_digits) == 13): # Já formatado +5581...
         logger.debug(f"Telefone '{tel_original}' já parece estar no formato +55. Mantendo.")
         return tel_original
    else: # Formatos inesperados
        logger.warning(f"Telefone '{tel_original}' (dígitos: '{tel_digits}') com formato/comprimento inesperado. Não foi possível formatar para +55.")
        # Retorna o original ou só dígitos? Retornar original pode ser mais informativo no log principal
        return tel_original

# --- Função para limpar e FORMATAR o nome ---
def extrair_e_formatar_nome(nome_completo):
    nome_extraido = nome_completo
    if isinstance(nome_completo, str) and ' - ' in nome_completo:
        try:
            partes = nome_completo.split(' - ', 1)
            if len(partes) > 1:
                nome_extraido = partes[1].strip()
        except Exception as e:
            logger.warning(f"Erro ao tentar extrair nome de '{nome_completo}': {e}. Usando original.")
            nome_extraido = nome_completo

    if isinstance(nome_extraido, str):
        nome_formatado = nome_extraido.lower().title()
        palavras_excecao = {' Da ', ' De ', ' Do ', ' Dos ', ' Das '}
        for exc in palavras_excecao:
            if exc in nome_formatado:
                 nome_formatado = nome_formatado.replace(exc, exc.lower())
        return nome_formatado
    else:
        return str(nome_extraido)

# --- Função carregar_inadimplentes ---
def carregar_inadimplentes(caminho_arquivo_planilha):
    """ Carrega e pré-processa dados da planilha Excel TRATADA. """
    logger.info(f"--- Iniciando leitura do arquivo Excel: {os.path.basename(caminho_arquivo_planilha)} ---")
    if not caminho_arquivo_planilha or not os.path.exists(caminho_arquivo_planilha):
        logger.error(f"Arquivo da planilha não encontrado em '{caminho_arquivo_planilha}'")
        return None

    try:
        df = pd.read_excel(caminho_arquivo_planilha, header=0, dtype=str) # Lê tudo como string
        colunas_obrigatorias_para_linha = [col for col in [COLUNA_NOME, COLUNA_TELEFONE] if col]
        if not colunas_obrigatorias_para_linha:
             logger.error("ERRO CRÍTICO: COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env.")
             return None
        df.dropna(subset=colunas_obrigatorias_para_linha, how='all', inplace=True) # Remove linhas sem nome NEM telefone

        logger.info(f"Arquivo Excel lido. {len(df)} linhas encontradas após remoção de linhas vazias iniciais.")
        if df.empty:
            logger.warning("Planilha parece vazia após leitura inicial ou remoção de linhas vazias.")
            return df
    except FileNotFoundError:
       logger.error(f"Arquivo não encontrado (verificado novamente) em '{caminho_arquivo_planilha}'")
       return None
    except KeyError as e:
        logger.error(f"Erro Crítico: Coluna essencial '{e}' definida no .env não encontrada na planilha.")
        logger.error(f"Verifique .env e cabeçalhos: {df.columns.to_list() if 'df' in locals() else 'Erro antes de ler colunas'}")
        return None
    except Exception as e:
       logger.exception(f"Erro Crítico ao ler o arquivo Excel '{os.path.basename(caminho_arquivo_planilha)}'")
       return None

    # --- Validação de Colunas ---
    colunas_esperadas = {
        'Nome': COLUNA_NOME,
        'Telefone': COLUNA_TELEFONE,
        'Valor': COLUNA_VALOR,
        'Vencimento': COLUNA_VENCIMENTO,
        'Lote': COLUNA_LOTE, 
        'LoteamentoInfo': COLUNA_LOTEAMENTO 
    }
    colunas_faltantes_nomes = []
    colunas_faltantes_keys = []
    colunas_presentes = []

    logger.info("Verificando colunas (definidas no .env):")
    for key, nome_real_coluna in colunas_esperadas.items():
        is_obrigatoria_leitor = key in ['Nome', 'Telefone']
        # Define o nome da variável de ambiente correspondente
        if key == 'Lote': env_var_name = 'LOTE'
        # <<< ALTERADO: Verifica a variável COLUNA_LOTEAMENTO >>>
        elif key == 'LoteamentoInfo': env_var_name = 'COLUNA_LOTEAMENTO'
        else: env_var_name = f"COLUNA_{key.upper()}"

        if nome_real_coluna: # Var do .env definida
            logger.debug(f"  - Esperando '{key}' na coluna: '{nome_real_coluna}' (via var {env_var_name})")
            if nome_real_coluna in df.columns:
                colunas_presentes.append(nome_real_coluna)
            else: # Coluna não encontrada
                if is_obrigatoria_leitor:
                    logger.error(f"  - Coluna OBRIGATÓRIA '{nome_real_coluna}' para '{key}' NÃO encontrada.")
                    colunas_faltantes_nomes.append(nome_real_coluna)
                    colunas_faltantes_keys.append(key)
                else:
                    logger.warning(f"  - Coluna OPCIONAL '{nome_real_coluna}' para '{key}' não encontrada (Var {env_var_name} definida mas coluna ausente).")
        else: # Var do .env NÃO definida
            if is_obrigatoria_leitor:
                 logger.error(f"  - Variável de ambiente OBRIGATÓRIA {env_var_name} não definida no .env")
                 colunas_faltantes_nomes.append(f"Variável {env_var_name} não definida")
                 colunas_faltantes_keys.append(key)
            else:
                 logger.info(f"  - Variável de ambiente OPCIONAL {env_var_name} não definida. Campo '{key}' não será lido/usado.")

    # --- Verificação de Erros de Coluna OBRIGATÓRIA ---
    if any(key in ['Nome', 'Telefone'] for key in colunas_faltantes_keys):
        logger.error("Erro Crítico: Colunas/Variáveis obrigatórias (Nome/Telefone) ausentes:")
        for i in range(len(colunas_faltantes_nomes)):
             logger.error(f"  - {colunas_faltantes_keys[i]}: {colunas_faltantes_nomes[i]}")
        logger.error(f"Colunas encontradas na planilha: {df.columns.to_list()}")
        return None
    if (COLUNA_NOME and COLUNA_NOME not in colunas_presentes) or \
       (COLUNA_TELEFONE and COLUNA_TELEFONE not in colunas_presentes):
         logger.error("Erro Crítico: Colunas de Nome ou Telefone definidas no .env não foram encontradas.")
         logger.error(f"Colunas encontradas: {df.columns.to_list()}")
         return None

    logger.info("Colunas essenciais encontradas/configuradas.")
    logger.info(f"Iniciando processamento e limpeza de {len(df)} registros...")

    # --- Limpeza e Conversões ---
    # NOME
    if COLUNA_NOME in colunas_presentes:
        logger.debug(f"Limpando coluna de Nome '{COLUNA_NOME}'...")
        df[COLUNA_NOME] = df[COLUNA_NOME].apply(extrair_e_formatar_nome)

    # TELEFONE
    if COLUNA_TELEFONE in colunas_presentes:
        logger.debug(f"Processando coluna '{COLUNA_TELEFONE}'...")
        try:
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].apply(formatar_telefone_br)
        except Exception as e:
            logger.exception(f"Erro ao processar coluna de telefone '{COLUNA_TELEFONE}'")
            return None # Erro crítico na formatação do telefone

    # VENCIMENTO
    if COLUNA_VENCIMENTO in colunas_presentes:
        logger.debug(f"Processando coluna '{COLUNA_VENCIMENTO}'...")
        try:
            df[COLUNA_VENCIMENTO] = pd.to_datetime(df[COLUNA_VENCIMENTO], errors='coerce', dayfirst=True)
            nat_count = df[COLUNA_VENCIMENTO].isnull().sum()
            if nat_count > 0: logger.warning(f"{nat_count} valores em '{COLUNA_VENCIMENTO}' não convertidos para data (viraram NaT).")
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de vencimento '{COLUNA_VENCIMENTO}'")
 

    # VALOR
    if COLUNA_VALOR in colunas_presentes:
        logger.debug(f"Processando coluna '{COLUNA_VALOR}'...")
        try:
            # Tenta limpar R$, espaços, separador de milhar e ajustar decimal
            col_valor_str = df[COLUNA_VALOR].astype(str).str.strip()
            col_valor_limpa = col_valor_str.str.replace(r'^R\$\s*', '', regex=True) # Tira R$ do início
            # Remove pontos de milhar (APENAS se seguido por 3 dígitos e vírgula)
            col_valor_limpa = col_valor_limpa.str.replace(r'\.(?=\d{3},)', '', regex=True)
            # Troca vírgula decimal por ponto
            col_valor_limpa = col_valor_limpa.str.replace(',', '.', regex=False)
            # Remove espaços restantes (caso haja algo como '1 234.56')
            col_valor_limpa = col_valor_limpa.str.replace(r'\s+', '', regex=True)

            df[COLUNA_VALOR] = pd.to_numeric(col_valor_limpa, errors='coerce')
            nan_count = df[COLUNA_VALOR].isnull().sum()
            if nan_count > 0:
                logger.warning(f"{nan_count} valores em '{COLUNA_VALOR}' não convertidos para número (viraram NaN) após limpeza.")
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de valor '{COLUNA_VALOR}'")
            # Permite continuar com valores inválidos (NaN), loga o erro.
 
    logger.info("--- Leitura e pré-processamento do arquivo Excel concluídos ---")
    return df