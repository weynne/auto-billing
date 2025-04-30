#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# main.py

import time
import os
from dotenv import load_dotenv
import tkinter as tk
from tkinter import filedialog
import pandas as pd # Importado para usar pd.notna se necessário

# Carrega os módulos locais da aplicação
# Garanta que a pasta 'modules' está no mesmo nível que main.py
# e que contém os arquivos __init__.py (vazio), leitor_planilha.py, etc.
try:
    from modules import leitor_planilha
    from modules import construtor_mensagem
    from modules import arquivo_txt_sender # <--- USA O SENDER PARA ARQUIVOS .TXT
except ImportError as e:
     print("Erro Crítico: Não foi possível importar módulos da pasta 'modules'.")
     print(f"Detalhe: {e}")
     print("Verifique se a pasta 'modules' existe, contém um arquivo __init__.py (pode ser vazio)")
     print("e os outros arquivos .py necessários (leitor_planilha, etc.).")
     exit() # Encerra o script se módulos não puderem ser importados


# Carrega variáveis de ambiente do arquivo .env
# A função load_dotenv() procura pelo arquivo .env no diretório atual ou nos pais
if load_dotenv():
     print("Arquivo .env carregado com sucesso.")
else:
     print("Aviso: Arquivo .env não encontrado na raiz do projeto.")
     print("As configurações de nome de coluna e API dependerão das variáveis de ambiente do sistema (se existirem).")


# --- Configurações ---
# Pega nomes das colunas do .env que são usados diretamente em main.py
COLUNA_NOME = os.getenv('COLUNA_NOME')
COLUNA_TELEFONE = os.getenv('COLUNA_TELEFONE')

# Tempo de pausa entre o processamento de cada linha (em segundos)
# Pode ser útil para não sobrecarregar o sistema ou para simular um envio mais cadenciado
try:
    # Tenta ler do .env, senão usa o padrão 1
    DELAY_ENTRE_MENSAGENS_SEGUNDOS = int(os.getenv('DELAY_SEGUNDOS', 1))
except ValueError:
     print("Aviso: Valor inválido para DELAY_SEGUNDOS no .env. Usando padrão de 1 segundo.")
     DELAY_ENTRE_MENSAGENS_SEGUNDOS = 1


def selecionar_arquivo_planilha():
    """
    Abre uma janela gráfica para o usuário selecionar o arquivo da planilha Excel tratada.

    Returns:
        str: O caminho completo para o arquivo selecionado, ou None se o usuário cancelar.
    """
    print("\nAbrindo janela para seleção da planilha...")
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter
    # Força a janela de diálogo a aparecer na frente de outras janelas
    root.attributes('-topmost', True)

    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha Excel TRATADA",
        # Prioriza arquivos Excel, mas permite outros
        filetypes=[
            ("Arquivos Excel", "*.xlsx *.xls"),
            ("Arquivos CSV (Fallback)", "*.csv"), # Mantém CSV caso o usuário trate para CSV
            ("Todos os arquivos", "*.*")
        ]
    )
    root.destroy() # Fecha a janela tk principal oculta após seleção

    if caminho_arquivo:
        print(f"Arquivo selecionado: {caminho_arquivo}")
        return caminho_arquivo
    else:
        print("Nenhum arquivo foi selecionado.")
        return None

def processar_cobrancas():
    """
    Função principal para orquestrar a leitura, construção e salvamento das mensagens.
    """
    print("\n--- Iniciando Processo de 'Envio' de Cobranças (Salvando em .txt) ---")

    # Verifica se as colunas essenciais do .env foram carregadas para uso em main.py
    if not COLUNA_NOME or not COLUNA_TELEFONE:
         print("\nErro Crítico: Variáveis COLUNA_NOME ou COLUNA_TELEFONE não definidas no .env.")
         print("Verifique se o arquivo .env existe e contém essas definições (Ex: COLUNA_NOME=CLIENTE).")
         return # Encerra a função

    # 1. Selecionar a planilha via interface gráfica
    caminho_planilha_selecionada = selecionar_arquivo_planilha()

    if not caminho_planilha_selecionada:
        print("\nProcesso cancelado pelo usuário. Encerrando.")
        return

    # 2. Carregar e pré-processar dados da planilha usando o módulo leitor
    print(f"\n--- Carregando e Processando Planilha ---")
    dados_inadimplentes = leitor_planilha.carregar_inadimplentes(caminho_planilha_selecionada)

    # Verifica se o carregamento foi bem-sucedido
    if dados_inadimplentes is None:
        print("\nFalha ao carregar dados da planilha (verifique os erros acima). Encerrando.")
        return
    if dados_inadimplentes.empty:
        print("\nPlanilha carregada está vazia. Nenhum registro para processar. Encerrando.")
        return

    # 3. Iniciar processamento dos registros
    total_registros = len(dados_inadimplentes)
    salvos_sucesso = 0
    falhas = 0
    print(f"\n--- Iniciando processamento de {total_registros} registros ---")

    # Itera sobre cada linha do DataFrame
    for indice, cliente in dados_inadimplentes.iterrows():
        print("-" * 20) # Separador visual para cada registro
        print(f"Processando Registro {indice + 1}/{total_registros}...")

        # Extrai nome e telefone para validação e para o sender
        # Usar .get() com valor padrão é mais seguro caso a coluna exista mas a célula esteja vazia
        nome_cliente = cliente.get(COLUNA_NOME, "Nome Ausente")
        # O telefone já deve vir limpo (só dígitos) do leitor_planilha
        telefone = cliente.get(COLUNA_TELEFONE, "")

        # Validação básica do telefone (essencial para nome do arquivo e futuro envio)
        # Considera telefones brasileiros (fixo ou móvel) após limpeza
        if not telefone or not isinstance(telefone, str) or not (10 <= len(telefone) <= 11):
            print(f"AVISO: Telefone inválido ou ausente para '{nome_cliente}' (Valor: '{telefone}', Comprimento: {len(str(telefone))}). Pulando.")
            falhas += 1
            continue # Pula para o próximo cliente na planilha

        print(f"  - Cliente: {nome_cliente}")
        print(f"  - Telefone (limpo): {telefone}")

        # 4. Construir a mensagem personalizada usando o módulo construtor
        try:
             mensagem = construtor_mensagem.criar_mensagem_cobranca(cliente)
             # print(f"  - Mensagem Gerada:\n{mensagem[:150]}...") # Descomente para pré-visualizar
        except Exception as e_msg:
             print(f"ERRO: Falha ao construir mensagem para '{nome_cliente}': {e_msg}. Pulando.")
             falhas += 1
             continue

        # 5. Salvar a mensagem em arquivo .txt usando o módulo sender
        print(f"  - Tentando salvar mensagem em arquivo .txt...")
        try:
             sucesso_salvar = arquivo_txt_sender.salvar_mensagem_em_txt(telefone, nome_cliente, mensagem)
        except Exception as e_save:
             print(f"ERRO: Falha inesperada ao tentar salvar arquivo para '{nome_cliente}': {e_save}")
             sucesso_salvar = False # Marca como falha

        # Atualiza contadores e loga resultado
        if sucesso_salvar:
            salvos_sucesso += 1
            print("  - Mensagem salva com sucesso.")
        else:
            falhas += 1
            print("  - Falha ao salvar a mensagem (ver logs do arquivo_txt_sender).")

        # 6. Pausa entre processamentos
        if total_registros > 1 and indice < total_registros - 1: # Só pausa se houver mais de um e não for o último
             print(f"  - Aguardando {DELAY_ENTRE_MENSAGENS_SEGUNDOS} segundo(s)...")
             time.sleep(DELAY_ENTRE_MENSAGENS_SEGUNDOS)

    # Fim do loop for

    # 7. Resumo final do processamento
    print("\n" + "="*40)
    print("--- Processo Concluído ---")
    print(f"Planilha processada: {os.path.basename(caminho_planilha_selecionada)}")
    print(f"Total de registros na planilha: {total_registros}")
    print(f"Mensagens salvas com sucesso em arquivos .txt: {salvos_sucesso}")
    print(f"Registros com falha ou pulados (telefone inválido/erro): {falhas}")
    if salvos_sucesso > 0:
        print(f"Verifique os arquivos na pasta: '{arquivo_txt_sender.DIRETORIO_SAIDA}'")
    print("="*40)

# Ponto de entrada principal do script
if __name__ == "__main__":
    # Verifica se as variáveis de ambiente essenciais foram carregadas antes de iniciar
    # Isso ajuda a pegar erros de .env ausente ou mal configurado logo no início.
    if not COLUNA_NOME or not COLUNA_TELEFONE:
        print("ERRO CRÍTICO: Variáveis COLUNA_NOME ou COLUNA_TELEFONE não carregadas do .env.")
        print("Certifique-se que o arquivo .env existe na raiz do projeto e contém as configurações corretas.")
        print("Exemplo:")
        print("COLUNA_NOME=CLIENTE")
        print("COLUNA_TELEFONE=TELEFONE")
        print("COLUNA_VALOR=SALDO")
        print("COLUNA_VENCIMENTO=VENCTO.")
    else:
        # Se as configurações básicas parecem OK, inicia o processo
        processar_cobrancas()