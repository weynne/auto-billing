# -*- coding: utf-8 -*-
# modules/leitor_planilha.py

import pandas as pd
import os
from dotenv import load_dotenv

# Carrega variáveis de ambiente do arquivo .env no diretório raiz do projeto
load_dotenv()

# Pega os nomes das colunas esperadas do .env
COLUNA_NOME = os.getenv('COLUNA_NOME') # Ex: CLIENTE
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE') # Ex: TELEFONE
COLUNA_VALOR = os.getenv('COLUNA_VALOR') # Ex: SALDO
COLUNA_VENCIMENTO = os.getenv('COLUNA_VENCIMENTO') # Ex: VENCTO.

def carregar_inadimplentes(caminho_arquivo_planilha):
    """
    Carrega os dados da planilha de inadimplentes de um arquivo Excel (.xls ou .xlsx) TRATADO.
    Assume cabeçalho na linha 1 e colunas definidas no .env. Inclui debug prints.

    Args:
        caminho_arquivo_planilha (str): Caminho completo para o arquivo Excel.

    Returns:
        pandas.DataFrame: DataFrame processado ou None se ocorrer erro.
    """
    print(f"--- Iniciando leitura do arquivo Excel: {os.path.basename(caminho_arquivo_planilha)} ---")
    if not caminho_arquivo_planilha or not os.path.exists(caminho_arquivo_planilha):
        print(f"Erro Crítico: Arquivo da planilha não encontrado em '{caminho_arquivo_planilha}'")
        return None

    try:
        # --- Leitura Direta de Excel ---
        df = pd.read_excel(caminho_arquivo_planilha, header=0)
        print(f"Arquivo Excel lido com sucesso. {len(df)} linhas encontradas antes da validação.")

    except FileNotFoundError:
        print(f"Erro Crítico: Arquivo não encontrado (verificado novamente) em '{caminho_arquivo_planilha}'")
        return None
    except Exception as e:
        print(f"Erro Crítico ao ler o arquivo Excel '{os.path.basename(caminho_arquivo_planilha)}': {e}")
        print("Verifique se o arquivo é um Excel válido (.xls ou .xlsx) e se 'openpyxl'/'xlrd' está instalado.")
        return None

    # --- Validação de Colunas Essenciais ---
    colunas_esperadas = {
        'Nome': COLUNA_NOME,
        'Telefone': COLUNA_TELEFONE,
        'Valor': COLUNA_VALOR,
        'Vencimento': COLUNA_VENCIMENTO
    }
    colunas_faltantes_nomes = []
    colunas_faltantes_keys = []
    colunas_presentes = []

    print("Verificando colunas essenciais (definidas no .env):")
    for key, col_name in colunas_esperadas.items():
        if col_name:
            print(f"  - Esperando '{key}' na coluna: '{col_name}'")
            if col_name in df.columns:
                 colunas_presentes.append(col_name)
            else:
                colunas_faltantes_nomes.append(col_name)
                colunas_faltantes_keys.append(key)
        else:
             print(f"  - Aviso: Variável de ambiente para '{key}' não definida no .env (Ex: COLUNA_{key.upper()})")

    if colunas_faltantes_nomes:
        print(f"\nErro Crítico: Colunas essenciais NÃO encontradas:")
        for i in range(len(colunas_faltantes_nomes)):
             print(f"  - {colunas_faltantes_keys[i]}: Esperava '{colunas_faltantes_nomes[i]}'")
        print(f"\nColunas encontradas: {df.columns.to_list()}")
        print("Verifique .env e cabeçalhos da planilha.")
        return None

    if not colunas_presentes:
         print("\nErro Crítico: Nenhuma coluna essencial encontrada/definida. Não é possível continuar.")
         return None

    print("Todas as colunas essenciais definidas no .env foram encontradas.")
    print(f"Iniciando processamento e limpeza de {len(df)} registros...")

    # --- Limpeza e Conversões (COM DEBUGGING ADICIONAL) ---
    print("DEBUG: Iniciando limpeza/conversão...") # <-- DEBUG ADD

    # Telefone: Converter para string e limpar não-numéricos
    if COLUNA_TELEFONE and COLUNA_TELEFONE in df.columns:
        print(f"DEBUG: Tentando limpar coluna '{COLUNA_TELEFONE}'...") # <-- DEBUG ADD
        try:
            # Garante que é string antes de usar métodos .str, tratando NaNs e números
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].apply(lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if pd.notna(x) else '').astype(str)
            df[COLUNA_TELEFONE] = df[COLUNA_TELEFONE].str.replace(r'[^\d]', '', regex=True) # Remove tudo que não for dígito
            print(f"DEBUG: Coluna '{COLUNA_TELEFONE}' limpa.") # <-- DEBUG MOD
        except Exception as e:
             # Tornar o erro mais óbvio no log
             print(f"!! ERRO DETALHADO !! ao limpar coluna de telefone '{COLUNA_TELEFONE}': {e}") # <-- DEBUG MOD
             # print(f"Erro ocorreu processando dados como: {df[[COLUNA_TELEFONE]].head()}") # Descomente CUIDADOSAMENTE para ver dados
             return None # Retorna None imediatamente se a limpeza falhar

    # Data de Vencimento: Converter para datetime
    if COLUNA_VENCIMENTO and COLUNA_VENCIMENTO in df.columns:
        print(f"DEBUG: Tentando converter coluna '{COLUNA_VENCIMENTO}' para data...") # <-- DEBUG ADD
        try:
            original_dtype = df[COLUNA_VENCIMENTO].dtype
            # Tenta converter, erros viram NaT (Not a Time)
            df[COLUNA_VENCIMENTO] = pd.to_datetime(df[COLUNA_VENCIMENTO], errors='coerce')
            nat_count = df[COLUNA_VENCIMENTO].isnull().sum()
            if nat_count > 0:
                 print(f"Aviso: {nat_count} valores na coluna '{COLUNA_VENCIMENTO}' não puderam ser convertidos para data.")
            # Só imprime sucesso se nenhum erro ocorreu E se o tipo mudou
            elif not pd.api.types.is_datetime64_any_dtype(original_dtype):
                 print(f"DEBUG: Coluna '{COLUNA_VENCIMENTO}' convertida para data.") # <-- DEBUG MOD
        except Exception as e:
             print(f"!! ERRO DETALHADO !! ao converter coluna de vencimento '{COLUNA_VENCIMENTO}': {e}") # <-- DEBUG MOD
             # print(f"Erro ocorreu processando dados como: {df[[COLUNA_VENCIMENTO]].head()}")
             return None

    # Valor: Converter para numérico (float)
    if COLUNA_VALOR and COLUNA_VALOR in df.columns:
        print(f"DEBUG: Tentando converter coluna '{COLUNA_VALOR}' para numérico...") # <-- DEBUG ADD
        try:
            original_dtype = df[COLUNA_VALOR].dtype
            # Só tenta limpar a string se o tipo não for numérico
            if not pd.api.types.is_numeric_dtype(df[COLUNA_VALOR]):
                print(f"DEBUG: Coluna '{COLUNA_VALOR}' não é numérica (tipo: {original_dtype}), tentando limpar...") # <-- DEBUG MOD
                # Limpa: remove R$, espaço, ponto milhar. Troca vírgula decimal por ponto.
                df[COLUNA_VALOR] = df[COLUNA_VALOR].astype(str).str.replace(r'[R$\s.]', '', regex=True).str.replace(',', '.', regex=False)

            # Converte para numérico, erros viram NaN (Not a Number)
            df[COLUNA_VALOR] = pd.to_numeric(df[COLUNA_VALOR], errors='coerce')
            nan_count = df[COLUNA_VALOR].isnull().sum()
            if nan_count > 0:
                print(f"Aviso: {nan_count} valores na coluna '{COLUNA_VALOR}' não puderam ser convertidos para número.")
            # Só imprime sucesso se nenhum erro ocorreu E se o tipo mudou
            elif not pd.api.types.is_numeric_dtype(original_dtype):
                 print(f"DEBUG: Coluna '{COLUNA_VALOR}' convertida para numérico.") # <-- DEBUG MOD
        except Exception as e:
              print(f"!! ERRO DETALHADO !! ao converter coluna de valor '{COLUNA_VALOR}': {e}") # <-- DEBUG MOD
              # print(f"Erro ocorreu processando dados como: {df[[COLUNA_VALOR]].head()}")
              return None

    print("DEBUG: Fim da limpeza/conversão.") # <-- DEBUG ADD
    print("--- Leitura e pré-processamento do arquivo Excel concluídos ---")
    return df # Retorna o DataFrame se tudo correu bem até aqui

# Bloco para teste rápido (opcional, requer um arquivo de teste)
# if __name__ == '__main__':
#     print("Executando teste do leitor_planilha...")
#     # Crie um arquivo 'teste_tratado.xlsx' com as colunas corretas para testar
#     arquivo_teste = 'teste_tratado.xlsx'
#     if os.path.exists(arquivo_teste):
#         df_teste = carregar_inadimplentes(arquivo_teste)
#         if df_teste is not None:
#             print("\n--- DataFrame de Teste Carregado ---")
#             print(df_teste.head())
#             print("\n--- Informações do DataFrame de Teste ---")
#             df_teste.info()
#             # Verifique os tipos de dados das colunas VALOR, VENCIMENTO, TELEFONE
#         else:
#             print("\nTeste falhou. DataFrame não foi carregado.")
#     else:
#          print(f"\nArquivo de teste '{arquivo_teste}' não encontrado. Crie-o para testar.")