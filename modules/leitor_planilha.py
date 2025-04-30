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

# Obtém logger
logger = logging.getLogger(__name__)

# --- Função formatar_telefone_br ---
def formatar_telefone_br(tel_original):
    """ Tenta formatar um número de telefone para o padrão E.164 (+55). """
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
        return tel_digits # Retorna só dígitos para validação posterior

# --- Função para limpar e FORMATAR o nome ---
def extrair_e_formatar_nome(nome_completo):
    """ Extrai o nome de 'COD - NOME' e aplica Title Case. """
    nome_extraido = nome_completo # Default to original

    # 1. Extrai o nome após ' - ' (se existir)
    if isinstance(nome_completo, str) and ' - ' in nome_completo:
        try:
            partes = nome_completo.split(' - ', 1)
            if len(partes) > 1:
                nome_extraido = partes[1].strip()
        except Exception as e:
            logger.warning(f"Erro ao tentar extrair nome de '{nome_completo}': {e}. Usando original.")
            nome_extraido = nome_completo # Garante que não falhe

    # 2. Aplica Title Case (Primeiro Minúsculo, depois Title)
    if isinstance(nome_extraido, str):
         # Converte tudo para minúsculo e depois aplica Title Case
         # Isso garante que "DE" ou "DA" se tornem "De", "Da"
         # e que "NOME TODO MAIUSCULO" vire "Nome Todo Maiusculo"
         nome_formatado = nome_extraido.lower().title()
         # Pequena correção para 'Da', 'De', 'Do', 'Dos', 'Das' (comum em nomes)
         # Pode adicionar mais exceções se necessário
         palavras_excecao = {' Da ', ' De ', ' Do ', ' Dos ', ' Das '}
         for exc in palavras_excecao:
              if exc in nome_formatado:
                   nome_formatado = nome_formatado.replace(exc, exc.lower())
         return nome_formatado
    else:
         # Se não for string (ex: número lido como nome), apenas converte para string
         return str(nome_extraido)

# --- Função carregar_inadimplentes ---
def carregar_inadimplentes(caminho_arquivo_planilha):
    """ Carrega e pré-processa dados da planilha Excel TRATADA. """
    logger.info(f"--- Iniciando leitura do arquivo Excel: {os.path.basename(caminho_arquivo_planilha)} ---")
    if not caminho_arquivo_planilha or not os.path.exists(caminho_arquivo_planilha):
        logger.error(f"Arquivo da planilha não encontrado em '{caminho_arquivo_planilha}'")
        return None

    try:
        # Especifica dtype=str para ler tudo como texto inicialmente, evitando problemas com tipos mistos
        df = pd.read_excel(caminho_arquivo_planilha, header=0, dtype=str)
        # Remove linhas onde colunas essenciais (Nome, Telefone) possam ser completamente vazias após leitura
        df.dropna(subset=[COLUNA_NOME, COLUNA_TELEFONE], how='all', inplace=True)
        logger.info(f"Arquivo Excel lido. {len(df)} linhas encontradas após remoção de linhas vazias iniciais.")
        if df.empty:
             logger.warning("Planilha parece vazia após leitura inicial ou remoção de linhas vazias.")
             return df # Retorna dataframe vazio
    except FileNotFoundError:
         logger.error(f"Arquivo não encontrado (verificado novamente) em '{caminho_arquivo_planilha}'")
         return None
    except KeyError as e:
         logger.error(f"Erro Crítico: Coluna essencial '{e}' (definida em COLUNA_NOME ou COLUNA_TELEFONE no .env) não encontrada na planilha ao tentar remover linhas vazias.")
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
        'Lote': COLUNA_LOTE
    }
    colunas_faltantes_nomes = []
    colunas_faltantes_keys = []
    colunas_presentes = []

    logger.info("Verificando colunas (definidas no .env):")
    for key, nome_real_coluna in colunas_esperadas.items():
        is_obrigatoria_leitor = key in ['Nome', 'Telefone']
        env_var_name = 'LOTE' if key == 'Lote' else f"COLUNA_{key.upper()}"

        if nome_real_coluna: # Var do .env definida
            logger.debug(f"  - Esperando '{key}' na coluna: '{nome_real_coluna}' (via var {env_var_name})")
            if nome_real_coluna in df.columns:
                colunas_presentes.append(nome_real_coluna)
            else: # Coluna não encontrada
                if is_obrigatoria_leitor:
                    logger.error(f"  - Coluna OBRIGATÓRIA '{nome_real_coluna}' para '{key}' NÃO encontrada.")
                    colunas_faltantes_nomes.append(nome_real_coluna)
                    colunas_faltantes_keys.append(key)
                else:
                    logger.warning(f"  - Coluna OPCIONAL '{nome_real_coluna}' para '{key}' não encontrada (Var {env_var_name} definida mas coluna ausente).")
        else: # Var do .env NÃO definida
            if is_obrigatoria_leitor:
                 logger.error(f"  - Variável de ambiente OBRIGATÓRIA {env_var_name} não definida no .env")
                 colunas_faltantes_nomes.append(f"Variável {env_var_name} não definida")
                 colunas_faltantes_keys.append(key)
            else:
                 logger.info(f"  - Variável de ambiente OPCIONAL {env_var_name} não definida. Campo '{key}' não será lido/processado.")

    # --- Verificação de Erros de Coluna ---
    if colunas_faltantes_nomes:
        logger.error("Erro Crítico: Colunas/Variáveis obrigatórias ausentes:")
        for i in range(len(colunas_faltantes_nomes)):
             logger.error(f"  - {colunas_faltantes_keys[i]}: {colunas_faltantes_nomes[i]}")
        logger.error(f"Colunas encontradas na planilha: {df.columns.to_list()}")
        return None
    if not COLUNA_NOME in colunas_presentes or not COLUNA_TELEFONE in colunas_presentes:
         logger.error("Erro Crítico: Colunas de Nome ou Telefone não encontradas (verificar .env e cabeçalhos).")
         return None

    logger.info("Colunas essenciais encontradas/configuradas.")
    logger.info(f"Iniciando processamento e limpeza de {len(df)} registros...")

    # --- Limpeza e Conversões ---
    logger.debug("Iniciando limpeza/conversão...")

    # Limpeza do NOME
    if COLUNA_NOME in colunas_presentes:
         logger.debug(f"Limpando coluna de Nome '{COLUNA_NOME}' (removendo código e ' - ')...")
         # Aplica a função de limpeza (já garante ser string)
         df[COLUNA_NOME] = df[COLUNA_NOME].apply(extrair_e_formatar_nome)
         logger.debug(f"Coluna '{COLUNA_NOME}' limpa.")

    # Limpeza/Formatação Telefone
    if COLUNA_TELEFONE in colunas_presentes:
        logger.debug(f"Processando coluna '{COLUNA_TELEFONE}'...")
        try:
            # Aplica a função de formatação (que já lida com string/None)
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].apply(formatar_telefone_br)
            logger.debug(f"Coluna '{COLUNA_TELEFONE}' processada para formato E.164 (+55 quando aplicável).")
        except Exception as e:
            logger.exception(f"Erro ao processar coluna de telefone '{COLUNA_TELEFONE}'")
            return None # Erro crítico
        
    # Conversão Data de Vencimento
    if COLUNA_VENCIMENTO in colunas_presentes:
        logger.debug(f"Processando coluna '{COLUNA_VENCIMENTO}'...")
        try:
            # --- ALTERAÇÃO AQUI: Adicionado dayfirst=True ---
            df[COLUNA_VENCIMENTO] = pd.to_datetime(df[COLUNA_VENCIMENTO], errors='coerce', dayfirst=True)
            # ----------------------------------------------
            nat_count = df[COLUNA_VENCIMENTO].isnull().sum()
            if nat_count > 0: logger.warning(f"{nat_count} valores em '{COLUNA_VENCIMENTO}' não convertidos para data (viraram NaT).")
            logger.debug(f"Coluna '{COLUNA_VENCIMENTO}' convertida para data (erros viraram NaT).")
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de vencimento '{COLUNA_VENCIMENTO}'")
            return None # Parar por segurança

    # Conversão Valor
    if COLUNA_VALOR in colunas_presentes:
        logger.debug(f"Processando coluna '{COLUNA_VALOR}'...")
        try:
            # 1. Garante que é string para manipulação
            col_valor_str = df[COLUNA_VALOR].astype(str)
            # 2. Remove R$ e espaços em branco
            col_valor_str = col_valor_str.str.replace(r'[R$\s]', '', regex=True)
            # 3. Remove pontos de milhar (APENAS se existir vírgula depois)
            col_valor_limpa = col_valor_str.str.replace(r'\.(?=.*\d{3},)', '', regex=True) # Tira ponto se seguido por 3 digitos e vírgula
            # 4. Substitui a vírgula decimal por ponto (se houver)
            col_valor_limpa = col_valor_limpa.str.replace(',', '.', regex=False)

            # 5. Converte para numérico
            df[COLUNA_VALOR] = pd.to_numeric(col_valor_limpa, errors='coerce')

            nan_count = df[COLUNA_VALOR].isnull().sum()
            if nan_count > 0:
                logger.warning(f"{nan_count} valores em '{COLUNA_VALOR}' não convertidos para número (viraram NaN) após limpeza.")
            logger.debug(f"Coluna '{COLUNA_VALOR}' convertida para numérico.")
        except Exception as e:
            logger.exception(f"Erro ao converter coluna de valor '{COLUNA_VALOR}'")
            return None # Parar por segurança

    # Coluna Lote (COLUNA_LOTE) - Já lida como string, nenhuma conversão extra necessária aqui

    logger.info("--- Leitura e pré-processamento do arquivo Excel concluídos ---")
    return df