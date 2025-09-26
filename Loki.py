import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.ttk import Progressbar
from PIL import Image, ImageTk
from threading import Thread
from functools import partial
from configparser import ConfigParser
import gdown
import logging
import urllib.request
import requests
import time
import shutil
import random
import string
import subprocess
import pyodbc

# Redirecionar stdout e stderr para arquivos temporários
if not sys.stdout:
    sys.stdout = open(os.devnull, 'w')
if not sys.stderr:
    sys.stderr = open(os.devnull, 'w')

# URL do arquivo de versão no Google Drive
VERSION_URL = "https://drive.google.com/uc?export=download&id=1OvC1v7RFKFoYk0rLLIC84lNSwLn2RLAq"

# URL direta de download do Google Drive para o pacote de atualização
UPDATE_URL = "https://drive.google.com/uc?export=download&id=1tz6qQPt-SUCc8K4JDIz-mGxgBjT5-ONI&confirm=t"


# Definir a versão atual da aplicação no código
CURRENT_VERSION = "8.2.3"

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Configuração adicional para evitar o erro de codificação
logging.getLogger("patoolib").setLevel(logging.DEBUG)
for handler in logging.getLogger("patoolib").handlers:
    if isinstance(handler, logging.StreamHandler):
        handler.setStream(sys.stdout)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        handler.encoding = 'utf-8'



# Printando retornos de teste
def print_status(message):
    print("="*60)
    print(f"Status: {message}")
    print("="*60)

# Exibir mensagens de log na interface gráfica
def print_status_gui(message):

    status_label.config(text=message, fg="")
    root.update()

# Exibir mensagens de Atualização na interface gráfica
def atualizar_status(message,root):

    print_status(message)
    status_label.config(text=message)
    root.update()

    pass

# Criando as pastas em seus respectivos locais D:\BDS\Dados e D:\BDS\Logs
def criar_pastas():
    nome_pasta = nome_pasta_entry.get()

    caminho_base_logs = os.path.join("D:\\BDS\\Logs", nome_pasta)
    caminho_base_dados = os.path.join("D:\\BDS\\Dados", nome_pasta)

    # Verificar se o nome da pasta contém "_"
    if "_" not in nome_pasta:
        messagebox.showerror("Erro", "O nome da pasta deve conter um caractere '_' (underline).")
        return

    if not os.path.exists(caminho_base_logs) and not os.path.exists(caminho_base_dados):
        os.makedirs(caminho_base_logs)

        opcao = opcao_combobox.get()

        if opcao == "Total Contador Fit":
            caminho_ac_logs = os.path.join(caminho_base_logs, "AC")
            os.makedirs(caminho_ac_logs)

            caminho_ac_dados = os.path.join(caminho_base_dados, "AC")
            os.makedirs(caminho_ac_dados)

        elif opcao == "Total Contador":
            caminho_ac_logs = os.path.join(caminho_base_logs, "AC")
            os.makedirs(caminho_ac_logs)

            caminho_ac_dados = os.path.join(caminho_base_dados, "AC")
            os.makedirs(caminho_ac_dados)

            caminho_patrio_logs = os.path.join(caminho_base_logs, "PATRIO")
            os.makedirs(caminho_patrio_logs)

            caminho_patrio_dados = os.path.join(caminho_base_dados, "PATRIO")
            os.makedirs(caminho_patrio_dados)

        elif opcao == "Apenas AG":
            caminho_ag_logs = os.path.join(caminho_base_logs, "AG")
            os.makedirs(caminho_ag_logs)

            caminho_ag_dados = os.path.join(caminho_base_dados, "AG")
            os.makedirs(caminho_ag_dados) 

        elif opcao == "Total Contador + AG":
            caminho_ac_logs = os.path.join(caminho_base_logs, "AC")
            os.makedirs(caminho_ac_logs)

            caminho_ac_dados = os.path.join(caminho_base_dados, "AC")
            os.makedirs(caminho_ac_dados)

            caminho_patrio_logs = os.path.join(caminho_base_logs, "PATRIO")
            os.makedirs(caminho_patrio_logs)

            caminho_patrio_dados = os.path.join(caminho_base_dados, "PATRIO")
            os.makedirs(caminho_patrio_dados)

            caminho_ag_logs = os.path.join(caminho_base_logs, "AG")
            os.makedirs(caminho_ag_logs)

            caminho_ag_dados = os.path.join(caminho_base_dados, "AG")
            os.makedirs(caminho_ag_dados)

        elif opcao == "AC e AG":
            caminho_ac_logs = os.path.join(caminho_base_logs, "AC")
            os.makedirs(caminho_ac_logs)

            caminho_ac_dados = os.path.join(caminho_base_dados, "AC")
            os.makedirs(caminho_ac_dados)

            caminho_ag_logs = os.path.join(caminho_base_logs, "AG")
            os.makedirs(caminho_ag_logs)

            caminho_ag_dados = os.path.join(caminho_base_dados, "AG")
            os.makedirs(caminho_ag_dados)

        elif opcao == "Apenas PONTO":
            caminho_ag_logs = os.path.join(caminho_base_logs, "PONTO")
            os.makedirs(caminho_ag_logs)

            caminho_ag_dados = os.path.join(caminho_base_dados, "PONTO")
            os.makedirs(caminho_ag_dados)

        elif opcao == "AC e PONTO":
            caminho_ac_logs = os.path.join(caminho_base_logs, "AC")
            os.makedirs(caminho_ac_logs)

            caminho_ac_dados = os.path.join(caminho_base_dados, "AC")
            os.makedirs(caminho_ac_dados)

            caminho_ag_logs = os.path.join(caminho_base_logs, "PONTO")
            os.makedirs(caminho_ag_logs)

            caminho_ag_dados = os.path.join(caminho_base_dados, "PONTO")
            os.makedirs(caminho_ag_dados)

        elif opcao == "Total Contador + PONTO":
            caminho_ac_logs = os.path.join(caminho_base_logs, "AC")
            os.makedirs(caminho_ac_logs)

            caminho_ac_dados = os.path.join(caminho_base_dados, "AC")
            os.makedirs(caminho_ac_dados)

            caminho_patrio_logs = os.path.join(caminho_base_logs, "PATRIO")
            os.makedirs(caminho_patrio_logs)

            caminho_patrio_dados = os.path.join(caminho_base_dados, "PATRIO")
            os.makedirs(caminho_patrio_dados)

            caminho_ag_logs = os.path.join(caminho_base_logs, "PONTO")
            os.makedirs(caminho_ag_logs)

            caminho_ag_dados = os.path.join(caminho_base_dados, "PONTO")
            os.makedirs(caminho_ag_dados)

        messagebox.showinfo("Sucesso", "Pastas criadas com sucesso. Agora, por favor, insira as letras, dígitos e senha.")
        habilitar_entradas()
        print_status("Pastas criadas com sucesso.")

    else:
        messagebox.showerror("Erro", f"A pasta '{nome_pasta}' já existe em D:\\BDS\\Logs e D:\\BDS\\Dados.")

# Habilitando entradas inicialmentes bloqueadas
def habilitar_entradas():
    letras_label.config(state=tk.NORMAL)
    letras_entry.config(state=tk.NORMAL)
    cpf_cnpj_label.config(state=tk.NORMAL)
    cpf_cnpj_entry.config(state=tk.NORMAL)
    senha_label.config(state=tk.NORMAL)
    senha_cliente.config(state=tk.NORMAL)
    criar_button.config(state=tk.NORMAL)
    gerar_senha_button.config(state=tk.NORMAL)

# Gerador de senhas
def gerar_senha():
    caracteres = string.ascii_letters + string.digits
    senha = ''.join(random.choice(caracteres) for _ in range(13))
    senha_cliente.delete(0, tk.END)  # Limpa o campo de senha existente
    senha_cliente.insert(0, senha)  # Insere a senha gerada no campo de senha

    # Copie a senha para a área de transferência
    root.clipboard_clear()  # Limpa a área de transferência
    root.clipboard_append(senha)  # Copia a senha para a área de transferência
    root.update()  # Atualiza a interface gráfica após a cópia

    # Notifique o usuário
    messagebox.showinfo("Senha gerada", f"Senha gerada: {senha}\nCopiada para a área de transferência")

# Criando os arquivos '.ini' dos clientes
def criar_ini(progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    if opcao_combobox.get() == "Total Contador + AG":
        criar_ini_ac_patrio_ag(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)
    
    elif opcao_combobox.get() == "AC e AG":
        criar_ini_ac_ag(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Total Contador":
        criar_ini_ac_patrio(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Apenas AG":
        criar_ini_ag(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Apenas PONTO":
        criar_ini_ponto(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "AC e PONTO":
        criar_ini_ac_ponto(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Total Contador + PONTO":
        criar_ini_ac_patrio_ponto(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    else:

        letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
        letras_minusculas = letras.lower()
        cpf_cnpj = cpf_cnpj_entry.get()
        senha = senha_cliente.get()

        # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
        if not letras or not cpf_cnpj or not senha:
            messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
            return

        # Crie o arquivo .ini no local desejado
        nome_pasta = nome_pasta_entry.get()
        nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
        caminho_ac_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", nome_cliente)
        os.makedirs(caminho_ac_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

        ini_path = os.path.join(caminho_ac_ini, f'{nome_cliente}.ini')
        with open(ini_path, 'w') as ini_file:
            ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AC\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AC\nDatabaseFile = 127.0.0.1:{letras}_AC_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

        nome_pasta = nome_pasta_entry.get()
        modificar_config_ini(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def criar_ini_ac_patrio_ag(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
    letras_minusculas = letras.lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
    if not letras or not cpf_cnpj or not senha:
        messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
        return
    

    # Crie o arquivo .ini AC no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ac_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", nome_cliente)
    os.makedirs(caminho_ac_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ac_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AC\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AC\nDatabaseFile = 127.0.0.1:{letras}_AC_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    # Crie o arquivo .ini do PATRIO no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ag_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", nome_cliente)
    os.makedirs(caminho_ag_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ag_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\PATRIO\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\PATRIO\nDatabaseFile = 127.0.0.1:{letras}_PATRIO_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    # Crie o arquivo .ini do AG no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ag_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", nome_cliente)
    os.makedirs(caminho_ag_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ag_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AG\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AG\nDatabaseFile = 127.0.0.1:{letras}_AG_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    nome_pasta = nome_pasta_entry.get()
    modificar_config_ini_ac_patrio_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def criar_ini_ac_patrio(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
    letras_minusculas = letras.lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
    if not letras or not cpf_cnpj or not senha:
        messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
        return
    

    # Crie o arquivo .ini AC no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ac_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", nome_cliente)
    os.makedirs(caminho_ac_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ac_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AC\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AC\nDatabaseFile = 127.0.0.1:{letras}_AC_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    # Crie o arquivo .ini do PATRIO no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ag_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", nome_cliente)
    os.makedirs(caminho_ag_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ag_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\PATRIO\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\PATRIO\nDatabaseFile = 127.0.0.1:{letras}_PATRIO_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    nome_pasta = nome_pasta_entry.get()
    modificar_config_ini_ac_patrio(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def criar_ini_ac_ag(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):
    letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
    letras_minusculas = letras.lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
    if not letras or not cpf_cnpj or not senha:
        messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
        return
    

    # Crie o arquivo .ini no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ac_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", nome_cliente)
    os.makedirs(caminho_ac_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ac_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AC\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AC\nDatabaseFile = 127.0.0.1:{letras}_AC_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    # Crie o arquivo .ini do AG no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ag_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", nome_cliente)
    os.makedirs(caminho_ag_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ag_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AG\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AG\nDatabaseFile = 127.0.0.1:{letras}_AG_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    nome_pasta = nome_pasta_entry.get()
    modificar_config_ini_ac_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto) 

def criar_ini_ag(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):
        
        letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
        letras_minusculas = letras.lower()
        cpf_cnpj = cpf_cnpj_entry.get()
        senha = senha_cliente.get()

        # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
        if not letras or not cpf_cnpj or not senha:
            messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
            return

        # Crie o arquivo .ini no local desejado
        nome_pasta = nome_pasta_entry.get()
        nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
        caminho_ag_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", nome_cliente)
        os.makedirs(caminho_ag_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

        ini_path = os.path.join(caminho_ag_ini, f'{nome_cliente}.ini')
        with open(ini_path, 'w') as ini_file:
            ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AG\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AG\nDatabaseFile = 127.0.0.1:{letras}_AG_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

        nome_pasta = nome_pasta_entry.get()
        modificar_config_ini_AG(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_patrio, caminho_arquivo_ponto)

def criar_ini_ponto(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

        letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
        letras_minusculas = letras.lower()
        cpf_cnpj = cpf_cnpj_entry.get()
        senha = senha_cliente.get()

        # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
        if not letras or not cpf_cnpj or not senha:
            messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
            return

        # Crie o arquivo .ini no local desejado
        nome_pasta = nome_pasta_entry.get()
        nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
        caminho_ponto_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", nome_cliente)
        os.makedirs(caminho_ponto_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

        ini_path = os.path.join(caminho_ponto_ini, f'{nome_cliente}.ini')
        with open(ini_path, 'w') as ini_file:
            ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\PONTO\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\PONTO\nDatabaseFile = 127.0.0.1:{letras}_PONTO_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

        nome_pasta = nome_pasta_entry.get()
        modificar_config_ini_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_patrio, caminho_arquivo_ponto)

def criar_ini_ac_ponto(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
    letras_minusculas = letras.lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
    if not letras or not cpf_cnpj or not senha:
        messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
        return
    

    # Crie o arquivo .ini no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ac_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", nome_cliente)
    os.makedirs(caminho_ac_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ac_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AC\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AC\nDatabaseFile = 127.0.0.1:{letras}_AC_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    # Crie o arquivo .ini do PONTO no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ponto_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", nome_cliente)
    os.makedirs(caminho_ponto_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ponto_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\PONTO\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\PONTO\nDatabaseFile = 127.0.0.1:{letras}_PONTO_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    nome_pasta = nome_pasta_entry.get()
    modificar_config_ini_ac_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)   

def criar_ini_ac_patrio_ponto(progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
    letras_minusculas = letras.lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Verificar se os campos letras, cpf_cnpj e senha não estão vazios
    if not letras or not cpf_cnpj or not senha:
        messagebox.showerror("Erro", "Os campos letras, digitos e senha são obrigatórios.")
        return
    

    # Crie o arquivo .ini AC no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ac_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", nome_cliente)
    os.makedirs(caminho_ac_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ac_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\AC\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\AC\nDatabaseFile = 127.0.0.1:{letras}_AC_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    # Crie o arquivo .ini do PATRIO no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_patrio_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", nome_cliente)
    os.makedirs(caminho_patrio_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_patrio_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\PATRIO\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\PATRIO\nDatabaseFile = 127.0.0.1:{letras}_PATRIO_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    # Crie o arquivo .ini do AG no local desejado
    nome_pasta = nome_pasta_entry.get()
    nome_cliente = nome_pasta  # Usando o mesmo nome para a pasta do .ini
    caminho_ponto_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", nome_cliente)
    os.makedirs(caminho_ponto_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

    ini_path = os.path.join(caminho_ponto_ini, f'{nome_cliente}.ini')
    with open(ini_path, 'w') as ini_file:
        ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\PONTO\nDataFolder = D:\\BDS\\Dados\\{nome_pasta}\\PONTO\nDatabaseFile = 127.0.0.1:{letras}_PONTO_{cpf_cnpj}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{cpf_cnpj}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')

    nome_pasta = nome_pasta_entry.get()
    modificar_config_ini_ac_patrio_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)


# Modificar o arquivo 'Config.ini' adicionando a linha de atualização do cliente

def modificar_config_ini(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    # Verifique se o arquivo existe antes de tentar abri-lo
    if not os.path.exists(config_ini_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {config_ini_path}")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AC,ExeDirName=AC,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AC_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # exibição da messagebox.showinfo
    exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def modificar_config_ini_ac_patrio(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AC,ExeDirName=AC,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AC_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Configurar .ini para PATRIO
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=PATRIO,ExeDirName=PATRIO,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_PATRIO_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # exibição da messagebox.showinfo
    exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def modificar_config_ini_ac_patrio_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AC,ExeDirName=AC,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AC_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Configurar .ini para PATRIO
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=PATRIO,ExeDirName=PATRIO,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_PATRIO_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AG,ExeDirName=AG,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AG_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # exibição da messagebox.showinfo
    exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def modificar_config_ini_AG(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_ponto):
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AG,ExeDirName=AG,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AG_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # exibição da messagebox.showinfo
    exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def modificar_config_ini_ac_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AC,ExeDirName=AC,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AC_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Configurar .ini para AG
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AG,ExeDirName=AG,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AG_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # exibição da messagebox.showinfo
    exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def modificar_config_ini_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AC,ExeDirName=AC,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AC_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # exibição da messagebox.showinfo
    exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def modificar_config_ini_ac_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

        config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

        with open(config_ini_path, 'r') as config_file:
            lines = config_file.readlines()

        # Encontre a seção [Operations] no arquivo
        operations_section_start = -1
        for i, line in enumerate(lines):
            if line.strip() == '[Operations]':
                operations_section_start = i
                break

        # Encontre o final da seção [Operations]
        operations_section_end = len(lines)
        for i in range(operations_section_start + 1, len(lines)):
            if lines[i].strip().startswith('['):
                operations_section_end = i
                break

        # Verifique se há linhas não comentadas dentro da seção [Operations]
        linhas_nao_comentadas = False
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                linhas_nao_comentadas = True
                break

        # Comente todas as linhas não comentadas dentro da seção [Operations]
        if linhas_nao_comentadas:
            for i in range(operations_section_start + 1, operations_section_end):
                if not lines[i].strip().startswith(';'):
                    lines[i] = ';' + lines[i]

        # Construa a nova linha com os valores especificados
        nova_linha = f"Customer={nome_pasta},ExeName=AC,ExeDirName=AC,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AC_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

        # Encontre a última linha comentada dentro da seção [Operations]
        ultima_linha_comentada_index = -1
        for i in range(operations_section_end - 1, operations_section_start, -1):
            if lines[i].strip().startswith(';'):
                ultima_linha_comentada_index = i
                break

        # Insira a nova linha após a última linha comentada ou no final da seção
        if ultima_linha_comentada_index != -1:
            lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
            lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
        else:
            # Se não houver linhas comentadas, insira no final da seção
            lines.insert(operations_section_end, nova_linha + '\n')

        # Modificar o arquivo para inserir a nova linha
        with open(config_ini_path, 'w') as config_file:
            config_file.writelines(lines)

        #Configurar .ini para PONTO
        config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", "config.ini")

        with open(config_ini_path, 'r') as config_file:
            lines = config_file.readlines()

        # Encontre a seção [Operations] no arquivo
        operations_section_start = -1
        for i, line in enumerate(lines):
            if line.strip() == '[Operations]':
                operations_section_start = i
                break

        # Encontre o final da seção [Operations]
        operations_section_end = len(lines)
        for i in range(operations_section_start + 1, len(lines)):
            if lines[i].strip().startswith('['):
                operations_section_end = i
                break

        # Verifique se há linhas não comentadas dentro da seção [Operations]
        linhas_nao_comentadas = False
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                linhas_nao_comentadas = True
                break

        # Comente todas as linhas não comentadas dentro da seção [Operations]
        if linhas_nao_comentadas:
            for i in range(operations_section_start + 1, operations_section_end):
                if not lines[i].strip().startswith(';'):
                    lines[i] = ';' + lines[i]

        # Construa a nova linha com os valores especificados
        nova_linha = f"Customer={nome_pasta},ExeName=PONTO,ExeDirName=PONTO,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_PONTO_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

        # Encontre a última linha comentada dentro da seção [Operations]
        ultima_linha_comentada_index = -1
        for i in range(operations_section_end - 1, operations_section_start, -1):
            if lines[i].strip().startswith(';'):
                ultima_linha_comentada_index = i
                break

        # Insira a nova linha após a última linha comentada ou no final da seção
        if ultima_linha_comentada_index != -1:
            lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
            lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
        else:
            # Se não houver linhas comentadas, insira no final da seção
            lines.insert(operations_section_end, nova_linha + '\n')

        # Modificar o arquivo para inserir a nova linha
        with open(config_ini_path, 'w') as config_file:
            config_file.writelines(lines)

        # exibição da messagebox.showinfo
        exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

def modificar_config_ini_ac_patrio_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=AC,ExeDirName=AC,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_AC_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Configurar .ini para PATRIO
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=PATRIO,ExeDirName=PATRIO,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_PATRIO_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Verifique se há linhas não comentadas dentro da seção [Operations]
    linhas_nao_comentadas = False
    for i in range(operations_section_start + 1, operations_section_end):
        if not lines[i].strip().startswith(';'):
            linhas_nao_comentadas = True
            break

    # Comente todas as linhas não comentadas dentro da seção [Operations]
    if linhas_nao_comentadas:
        for i in range(operations_section_start + 1, operations_section_end):
            if not lines[i].strip().startswith(';'):
                lines[i] = ';' + lines[i]

    # Construa a nova linha com os valores especificados
    nova_linha = f"Customer={nome_pasta},ExeName=PONTO,ExeDirName=PONTO,AppSubdir={nome_pasta},DbInstance=,DbName={letras}_PONTO_{cpf_cnpj},DbUser={letras_minusculas}_{cpf_cnpj}.sql,DbPass={senha}"

    # Encontre a última linha comentada dentro da seção [Operations]
    ultima_linha_comentada_index = -1
    for i in range(operations_section_end - 1, operations_section_start, -1):
        if lines[i].strip().startswith(';'):
            ultima_linha_comentada_index = i
            break

    # Insira a nova linha após a última linha comentada ou no final da seção
    if ultima_linha_comentada_index != -1:
        lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
        lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
    else:
        # Se não houver linhas comentadas, insira no final da seção
        lines.insert(operations_section_end, nova_linha + '\n')

    # Modificar o arquivo para inserir a nova linha
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # exibição da messagebox.showinfo
    exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

# Exibe a mensagem que dará inicio ao Retore dos banco

def exibir_mensagem(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):
    messagebox.showinfo("Sucesso", "Arquivo .ini criado e config.ini modificado com sucesso.\nIniciaremos o Restore do(s) Banco(s)")
    print_status("Arquivo .ini criado e config.ini modificado com sucesso.\nIniciaremos o Restore do(s) Banco(s)")

    ac(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha,root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

#Restore dos Bancos de dados de:

def ac(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    if opcao_combobox.get() == "Total Contador":
        ac_patrio(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Total Contador + AG":
        ac_patrio_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "AC e AG":
        ac_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Apenas AG":
        Apenas_AG(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Apenas PONTO":
        Apenas_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "AC e PONTO":
        ac_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    elif opcao_combobox.get() == "Total Contador + PONTO":
        ac_patrio_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto)

    else:
        # Configurar a barra de progresso
        progress_var.set(0)

        print_status(f'Iniciando restore de AC em MSSQL...')
        progress_var.set(10)  # Atualizar o valor da barra de progresso
        atualizar_status(f'Iniciando restore\nde AC em MSSQL...',root)

        nome_pasta = nome_pasta_entry.get()
        letras_minusculas = letras_entry.get().lower()
        cpf_cnpj = cpf_cnpj_entry.get()
        letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
        senha = senha_cliente.get()

        # Parâmetros de conexão com o SQL Server
        servers = ['FOR-CNT-BD-SQL5',
                   'FORT-CNT-BDSQ-6',
                   'GLAUDSON-DESKTO\\SQLEXPRESS']
        trusted_connection = 'yes'  # Configurando para autenticação do Windows
        driver = '{SQL Server}'


        # Parâmetros específicos para o restore
        banco_de_dados_destino_AC = f'{letras}_AC_{cpf_cnpj}'
        caminho_dados_AC = f'D:\\BDS\\dados\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.mdf'
        caminho_log_AC = f'D:\\BDS\\Logs\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.ldf'

        # Verificar se os bancos de dados já existem
        databases_to_check = [banco_de_dados_destino_AC]

        # Tentar conectar-se a cada servidor da lista
        conectado = False
        conn = None
        for server in servers:
            try:
                # Criar a string de conexão
                conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

                # Tentar estabelecer a conexão
                conn = pyodbc.connect(conn_str)

                # Se a conexão for bem sucedida, sair do loop
                print_status(f"Conectado ao servidor: {server}")
                progress_var.set(30)  # Atualizar o valor da barra de progresso
                atualizar_status(f"Conectado ao servidor: {server}", root)
                conectado = True
                break
            except Exception as e:
                print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
                atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

        if not conectado:
            print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
            atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
            return

        try:
            conn_check = pyodbc.connect(conn_str)
            cursor_check = conn_check.cursor()

            # Verificar se os bancos de dados já existem
            query_check = "SELECT name FROM sys.databases WHERE name IN (?)"
            cursor_check.execute(query_check, databases_to_check)
            existing_databases = [row[0] for row in cursor_check.fetchall()]

            if existing_databases:
                print('O banco de dados já existem. \nNão é possível continuar o restore.')
                print_status("Os bancos de dados já existem.\nNão é possível continuar o restore.")
                atualizar_status("Os bancos de dados já existem.\nNão é possível continuar o restore.", root)
                progress_var.set(30)


                tente_novamente(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
                return

        except Exception as e:
            print(f"Erro ao verificar a existência dos bancos de dados: {e}")
            atualizar_status(f"Erro ao verificar a\nexistência dos bancos de dados: {e}", root)
            return

        finally:
            if conn_check:
                conn_check.close()


        # Comando SQL para realizar o restore usando caminhos fixos para AC
        restore_query_AC = f'RESTORE DATABASE {banco_de_dados_destino_AC} FROM DISK = \'D:\\Bancos Limpos SQL\\AC.bak\' WITH REPLACE, MOVE \'AC_Data\' TO \'{caminho_dados_AC}\', MOVE \'AC_Log\' TO \'{caminho_log_AC}\''

        atualizar_status(f'Restore em andamento...',root)
        progress_var.set(30)  # Atualizar o valor da barra de progresso


        # Executar o comando de restore para AC
        try:
            subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AC], shell=True, check=True)
            print_status(f"Restore do banco {letras}_AC_{cpf_cnpj}\n executado com sucesso.")
            atualizar_status(f"Restore do banco\n{letras}_AC_{cpf_cnpj}\nexecutado com sucesso.", root)
            progress_var.set(50)  # Atualizar o valor da barra de progresso
        except subprocess.CalledProcessError as ex:
            print_status(f"Erro no restore do banco {letras}_AC_{cpf_cnpj}:\n {e}")
            atualizar_status(f"Erro no restore do banco\n{letras}_AC_{cpf_cnpj}:\n {e}", root)

        # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
        if os.path.isfile(caminho_dados_AC) and os.path.isfile(caminho_log_AC):
            # Alterar o modelo de recuperação para "Simples" para AC
            try:
                # Comando SQL para alterar o modelo de recuperação para "Simples" para AC
                alter_recovery_model_AC = f'sqlcmd -S {server} -d {banco_de_dados_destino_AC} -Q "ALTER DATABASE {banco_de_dados_destino_AC} SET RECOVERY SIMPLE;"'
                subprocess.run(alter_recovery_model_AC, check=True, shell=True)
                print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AC_{cpf_cnpj}.")
                atualizar_status(f"Modelo de recuperação\nalterado para 'Simples'\nEm {letras}_AC_{cpf_cnpj}.", root)
            except subprocess.CalledProcessError as ex:
                print_status(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}")
                atualizar_status(f"Erro ao alterar o modelo de\nrecuperação para 'Simples'\nEm {letras}_AC_{cpf_cnpj}: {ex}", root)

            # Executar a query pós restore para AC
            try:
                # Estabelecer a conexão
                conn_AC = pyodbc.connect(conn_str)

                # Criar um objeto cursor
                cursor_AC = conn_AC.cursor()

                # Executar USE master;
                cursor_AC.execute("USE master;")

                # Comando SQL para realizar a query pós restore para AC
                query_pos_restore_AC = f'''
                USE {letras_minusculas}_ac_{cpf_cnpj}
                CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ac_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
                CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
                EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

                INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
                '''

                # Executar a query pós restore para AC
                cursor_AC.execute(query_pos_restore_AC)
                print_status(f"Query pós restore para {letras}_AC_{cpf_cnpj}\n executada com sucesso.")
                atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\nexecutada com sucesso.", root)
                progress_var.set(80)  # Atualizar o valor da barra de progresso

                # Commit para confirmar as alterações
                conn_AC.commit()

                # Fechar o cursor e a conexão
                cursor_AC.close()
                conn_AC.close()

                # Adicione feedback visual ao usuário
                atualizar_status(f"Concluido!\nRestore executado\ncom sucesso.",root)
                progress_var.set(100)
            except pyodbc.Error as ex:
                print(f"Erro ao executar a query pós restore para {letras}_AC_{cpf_cnpj}: {ex}")
                atualizar_status(f"Erro ao executar\na query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}",root)


        else:
            print_status(f"Banco {letras}_AC_{cpf_cnpj} não foi restaurado com sucesso.\n Consulte os logs para mais detalhes.")
            atualizar_status(f"Banco {letras}_AC_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.",root)

        executaondemand_bat(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

def ac_patrio(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    progress_var.set(0)

    print_status(f'Iniciando restore de AC e PATRIO em MSSQL...')
    progress_var.set(10)  # Atualizar o valor da barra de progresso
    atualizar_status(f'Iniciando restore de\nAC e PATRIO em MSSQL...',root)

    nome_pasta = nome_pasta_entry.get()
    letras = letras_entry.get().upper()
    letras_minusculas = letras_entry.get().lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Parâmetros de conexão com o SQL Server
    servers = ['FOR-CNT-BD-SQL5',
               'FORT-CNT-BDSQ-6',
               'GLAUDSON-DESKTO\\SQLEXPRESS']
    trusted_connection = 'yes'  # Configurando para autenticação do Windows
    driver = '{SQL Server}'

    # Parâmetros específicos para o restore
    banco_de_dados_destino_AC = f'{letras}_AC_{cpf_cnpj}'
    caminho_dados_AC = f'D:\\BDS\\dados\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.mdf'
    caminho_log_AC = f'D:\\BDS\\Logs\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.ldf'

    banco_de_dados_destino_Patrio = f'{letras}_PATRIO_{cpf_cnpj}'
    caminho_dados_Patrio = f'D:\\BDS\\dados\\{nome_pasta}\\PATRIO\\{letras}_PATRIO_{cpf_cnpj}.mdf'
    caminho_log_Patrio = f'D:\\BDS\\Logs\\{nome_pasta}\\PATRIO\\{letras}_PATRIO_{cpf_cnpj}.ldf'

    # Verificar se os bancos de dados já existem
    databases_to_check = [banco_de_dados_destino_AC, banco_de_dados_destino_Patrio]

    # Tentar conectar-se a cada servidor da lista
    conectado = False
    conn = None
    for server in servers:
        try:
            # Criar a string de conexão
            conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

            # Tentar estabelecer a conexão
            conn = pyodbc.connect(conn_str)

            # Se a conexão for bem sucedida, sair do loop
            print_status(f"Conectado ao servidor: {server}")
            progress_var.set(30)  # Atualizar o valor da barra de progresso
            atualizar_status(f"Conectado ao servidor: {server}", root)
            conectado = True
            break
        except Exception as e:
            print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
            atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

    if not conectado:
        print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
        atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
        return

    try:
        conn_check = pyodbc.connect(conn_str)
        cursor_check = conn_check.cursor()

        # Verificar se os bancos de dados já existem
        query_check = "SELECT name FROM sys.databases WHERE name IN (?, ?)"
        cursor_check.execute(query_check, databases_to_check)
        existing_databases = [row[0] for row in cursor_check.fetchall()]

        if existing_databases:
            print('Os bancos de dados já existem.\n Não é possível continuar o restore.')
            print_status("Os bancos de dados já existem.\n Não é possível continuar o restore.")
            atualizar_status("Os bancos de dados já existem.\nNão é possível continuar o restore.", root)
            progress_var.set(30)


            tente_novamente_ac_patrio(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
            return

    except Exception as e:
        print(f"Erro ao verificar a existência dos bancos de dados: {e}")
        atualizar_status(f"Erro ao verificar\na existência dos\nbancos de dados: {e}", root)
        return

    finally:
        if conn_check:
            conn_check.close()

    # Comando SQL para realizar o restore usando caminhos fixos para AC
    restore_query_AC = f'RESTORE DATABASE {banco_de_dados_destino_AC} FROM DISK = \'D:\\Bancos Limpos SQL\\AC.bak\' WITH REPLACE, MOVE \'AC_Data\' TO \'{caminho_dados_AC}\', MOVE \'AC_Log\' TO \'{caminho_log_AC}\''

    # Comando SQL para realizar o restore usando caminhos fixos para Patrio
    restore_query_Patrio = f'RESTORE DATABASE {banco_de_dados_destino_Patrio} FROM DISK = \'D:\\Bancos Limpos SQL\\Patrio.bak\' WITH REPLACE, MOVE \'Patrio\' TO \'{caminho_dados_Patrio}\', MOVE \'Patrio_Log\' TO \'{caminho_log_Patrio}\''

    atualizar_status(f'Restore em Andamento...',root)
    progress_var.set(30)  # Atualizar o valor da barra de progresso

    # Executar o comando de restore para AC
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AC], shell=True, check=True)
        print_status(f"Restore do banco {letras}_AC_{cpf_cnpj} executado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro no restore do banco {letras}_AC_{cpf_cnpj}: {e}")

    # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_AC) and os.path.isfile(caminho_log_AC):
        # Alterar o modelo de recuperação para "Simples" para AC
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para AC
            alter_recovery_model_AC = f'sqlcmd -S {server} -d {banco_de_dados_destino_AC} -Q "ALTER DATABASE {banco_de_dados_destino_AC} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_AC, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AC_{cpf_cnpj}.")
            atualizar_status(f"Modelo de recuperação\nalterado para 'Simples'\nEm para {letras}_AC_{cpf_cnpj}.",root)
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao alterar o modelo\nde recuperação para 'Simples'\nEm {letras}_AC_{cpf_cnpj}: {ex}",root)

        # Executar a query pós restore para AC
        try:
            # Estabelecer a conexão
            conn_AC = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_AC = conn_AC.cursor()

            # Executar USE master;
            cursor_AC.execute("USE master;")

            # Comando SQL para realizar a query pós restore para AC
            query_pos_restore_AC = f'''
            USE {letras_minusculas}_ac_{cpf_cnpj}
            CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ac_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
            '''

            # Executar a query pós restore para AC
            cursor_AC.execute(query_pos_restore_AC)
            print_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.",root)

            # Commit para confirmar as alterações
            conn_AC.commit()

            # Fechar o cursor e a conexão
            cursor_AC.close()
            conn_AC.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"{letras}_AC_{cpf_cnpj}\nConcluido!", root)

            progress_var.set(40)  # Atualizar o valor da barra de progresso

        except pyodbc.Error as ex:
            print(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao executar\na query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}",root)

    else:
        print_status(f"Banco {letras}_AC_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    # Executar o comando de restore para PATRIO
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_Patrio], shell=True, check=True)
        print_status(f"Restore do banco\n{letras}_PATRIO_{cpf_cnpj}\nexecutado com sucesso.")
        progress_var.set(60)  # Atualizar o valor da barra de progresso
    except subprocess.CalledProcessError as ex:
        print(f"Erro no restore do banco {letras}_PATRIO_{cpf_cnpj}: {e}")

    # Verificar se o banco PATRIO foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_Patrio) and os.path.isfile(caminho_log_Patrio):
        # Alterar o modelo de recuperação para "Simples" para PATRIO
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para PATRIO
            alter_recovery_model_Patrio = f'sqlcmd -S {server} -d {banco_de_dados_destino_Patrio} -Q "ALTER DATABASE {banco_de_dados_destino_Patrio} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_Patrio, shell=True, check=True)
            print_status(f"Modelo de recuperação alterado para 'Simples' para {letras}_PATRIO_{cpf_cnpj}.")
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples' para {letras}_PATRIO_{cpf_cnpj}: {ex}")

        # Executar a query pós restore para PATRIO
        try:
            # Estabelecer a conexão
            conn_Patrio = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_Patrio = conn_Patrio.cursor()

            # Executar USE master;
            cursor_Patrio.execute("USE master;")

            # Comando SQL para realizar a query pós restore para PATRIO
            query_pos_restore_Patrio = f'''
            USE {letras_minusculas}_patrio_{cpf_cnpj}
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            '''

            # Executar a query pós restore para PATRIO
            cursor_Patrio.execute(query_pos_restore_Patrio)
            print_status(f"Query pós restore para\n{letras}_PATRIO_{cpf_cnpj}\nexecutada com sucesso.")
            atualizar_status(f"{letras}_PATRIO_{cpf_cnpj}\n Concluido!", root)
            progress_var.set(100)

            # Commit para confirmar as alterações
            conn_Patrio.commit()

            # Fechar o cursor e a conexão
            cursor_Patrio.close()
            conn_Patrio.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"Concluido!\nRestore executado com sucesso.",root)
            progress_var.set(100)

        except pyodbc.Error as ex:
            print_status(f"Erro ao executar a query pós restore para {letras}_PATRIO_{cpf_cnpj}: {ex}")

            atualizar_status(f"Erro ao executar\na query pós restore para\n{letras}_PATRIO_{cpf_cnpj}: {ex}", root, fg="red")


    else:
        print_status(f"Banco {letras}_PATRIO_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    executaondemand_bat_ac_patrio(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

def ac_patrio_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    progress_var.set(0)

    print_status(f'Iniciando restore\nde AC, AG e PATRIO em MSSQL...')
    progress_var.set(10)  # Atualizar o valor da barra de progresso
    atualizar_status(f'Iniciando restore\nde AC, AG e PATRIO em MSSQL...',root)

    nome_pasta = nome_pasta_entry.get()
    letras = letras_entry.get().upper()
    letras_minusculas = letras_entry.get().lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Parâmetros de conexão com o SQL Server
    servers = ['FOR-CNT-BD-SQL5', 
               'FORT-CNT-BDSQ-6',
               'GLAUDSON-DESKTO\\SQLEXPRESS']
    trusted_connection = 'yes'  # Configurando para autenticação do Windows
    driver = '{SQL Server}'

    # Parâmetros específicos para o restore
    banco_de_dados_destino_AC = f'{letras}_AC_{cpf_cnpj}'
    caminho_dados_AC = f'D:\\BDS\\dados\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.mdf'
    caminho_log_AC = f'D:\\BDS\\Logs\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.ldf'

    banco_de_dados_destino_Patrio = f'{letras}_PATRIO_{cpf_cnpj}'
    caminho_dados_Patrio = f'D:\\BDS\\dados\\{nome_pasta}\\PATRIO\\{letras}_PATRIO_{cpf_cnpj}.mdf'
    caminho_log_Patrio = f'D:\\BDS\\Logs\\{nome_pasta}\\PATRIO\\{letras}_PATRIO_{cpf_cnpj}.ldf'

    banco_de_dados_destino_Ag = f'{letras}_AG_{cpf_cnpj}'
    caminho_dados_Ag = f'D:\\BDS\\dados\\{nome_pasta}\\AG\\{letras}_AG_{cpf_cnpj}.mdf'
    caminho_log_Ag = f'D:\\BDS\\Logs\\{nome_pasta}\\AG\\{letras}_Ag_{cpf_cnpj}.ldf'

    # Verificar se os bancos de dados já existem
    databases_to_check = [banco_de_dados_destino_AC, banco_de_dados_destino_Patrio, banco_de_dados_destino_Ag]

    # Tentar conectar-se a cada servidor da lista
    conectado = False
    conn = None
    for server in servers:
        try:
            # Criar a string de conexão
            conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

            # Tentar estabelecer a conexão
            conn = pyodbc.connect(conn_str)

            # Se a conexão for bem sucedida, sair do loop
            print_status(f"Conectado ao servidor: {server}")
            progress_var.set(30)  # Atualizar o valor da barra de progresso
            atualizar_status(f"Conectado ao servidor: {server}", root)
            conectado = True
            break
        except Exception as e:
            print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
            atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

    if not conectado:
        print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
        atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
        return

    try:
        conn_check = pyodbc.connect(conn_str)
        cursor_check = conn_check.cursor()

        # Verificar se os bancos de dados já existem
        query_check = "SELECT name FROM sys.databases WHERE name IN (?, ?, ?)"
        cursor_check.execute(query_check, databases_to_check)
        existing_databases = [row[0] for row in cursor_check.fetchall()]

        if existing_databases:
            print('Os bancos de dados já existem.\n Não é possível continuar o restore.')
            print_status("Os bancos de dados já existem.\n Não é possível continuar o restore.")
            atualizar_status("Os bancos de dados já existem.\n Não é possível continuar o restore.", root)
            progress_var.set(30)


            tente_novamente_ac_patrio_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
            return

    except Exception as ex:
        print(f"Erro ao verificar a existência\ndos bancos de dados:\n{ex}")
        atualizar_status(f"Erro ao verificar a existência\ndos bancos de dados:\n{ex}", root)
        return

    finally:
        if conn_check:
            conn_check.close()

    # Comando SQL para realizar o restore usando caminhos fixos para AC
    restore_query_AC = f'RESTORE DATABASE {banco_de_dados_destino_AC} FROM DISK = \'D:\\Bancos Limpos SQL\\AC.bak\' WITH REPLACE, MOVE \'AC_Data\' TO \'{caminho_dados_AC}\', MOVE \'AC_Log\' TO \'{caminho_log_AC}\''

    # Comando SQL para realizar o restore usando caminhos fixos para Patrio
    restore_query_Patrio = f'RESTORE DATABASE {banco_de_dados_destino_Patrio} FROM DISK = \'D:\\Bancos Limpos SQL\\Patrio.bak\' WITH REPLACE, MOVE \'Patrio\' TO \'{caminho_dados_Patrio}\', MOVE \'Patrio_Log\' TO \'{caminho_log_Patrio}\''

    # Comando SQL para realizar o restore usando caminhos fixos para AG
    restore_query_Ag = f'RESTORE DATABASE {banco_de_dados_destino_Ag} FROM DISK = \'D:\\Bancos Limpos SQL\\AG.bak\' WITH REPLACE, MOVE \'AG_Data\' TO \'{caminho_dados_Ag}\', MOVE \'AG_Log\' TO \'{caminho_log_Ag}\''

    atualizar_status(f'Restore em Andamento...',root)
    progress_var.set(30)  # Atualizar o valor da barra de progresso

    # Executar o comando de restore para AC
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AC], shell=True, check=True)
        print_status(f"Restore do banco {letras}_AC_{cpf_cnpj} executado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro no restore do banco {letras}_AC_{cpf_cnpj}: {e}")

    # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_AC) and os.path.isfile(caminho_log_AC):
        # Alterar o modelo de recuperação para "Simples" para AC
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para AC
            alter_recovery_model_AC = f'sqlcmd -S {server} -d {banco_de_dados_destino_AC} -Q "ALTER DATABASE {banco_de_dados_destino_AC} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_AC, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AC_{cpf_cnpj}.")
            atualizar_status(f"Modelo de recuperação\nalterado para 'Simples'\nEm para {letras}_AC_{cpf_cnpj}.",root)
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao alterar o modelo\nde recuperação para 'Simples'\nEm {letras}_AC_{cpf_cnpj}: {ex}",root)

        # Executar a query pós restore para AC
        try:
            # Estabelecer a conexão
            conn_AC = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_AC = conn_AC.cursor()

            # Executar USE master;
            cursor_AC.execute("USE master;")

            # Comando SQL para realizar a query pós restore para AC
            query_pos_restore_AC = f'''
            USE {letras_minusculas}_ac_{cpf_cnpj}
            CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ac_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
            '''

            # Executar a query pós restore para AC
            cursor_AC.execute(query_pos_restore_AC)
            print_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\nexecutada com sucesso.",root)

            # Commit para confirmar as alterações
            conn_AC.commit()

            # Fechar o cursor e a conexão
            cursor_AC.close()
            conn_AC.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.",root)
            atualizar_status(f"{letras}_AC_{cpf_cnpj}\nConcluido!", root)
            progress_var.set(40)  # Atualizar o valor da barra de progresso

        except pyodbc.Error as ex:
            print(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}",root)

    else:
        print_status(f"Banco {letras}_AC_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    # Executar o comando de restore para PATRIO
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_Patrio], shell=True, check=True)
        print_status(f"Restore do banco\n{letras}_PATRIO_{cpf_cnpj}\nexecutado com sucesso.")
        progress_var.set(45)  # Atualizar o valor da barra de progresso
    except subprocess.CalledProcessError as e:
        print(f"Erro no restore do banco {letras}_PATRIO_{cpf_cnpj}: {e}")

    # Verificar se o banco PATRIO foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_Patrio) and os.path.isfile(caminho_log_Patrio):
        # Alterar o modelo de recuperação para "Simples" para PATRIO
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para PATRIO
            alter_recovery_model_Patrio = f'sqlcmd -S {server} -d {banco_de_dados_destino_Patrio} -Q "ALTER DATABASE {banco_de_dados_destino_Patrio} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_Patrio, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples' para {letras}_PATRIO_{cpf_cnpj}.")
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples' para {letras}_PATRIO_{cpf_cnpj}: {ex}")

        # Executar a query pós restore para PATRIO
        try:
            # Estabelecer a conexão
            conn_Patrio = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_Patrio = conn_Patrio.cursor()

            # Executar USE master;
            cursor_Patrio.execute("USE master;")

            # Comando SQL para realizar a query pós restore para PATRIO
            query_pos_restore_Patrio = f'''
            USE {letras_minusculas}_patrio_{cpf_cnpj}
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            '''

            # Executar a query pós restore para PATRIO
            cursor_Patrio.execute(query_pos_restore_Patrio)
            print_status(f"Query pós restore para\n{letras}_PATRIO_{cpf_cnpj}\nexecutada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_PATRIO_{cpf_cnpj}\nexecutada com sucesso.",root)
            progress_var.set(50)

            # Commit para confirmar as alterações
            conn_Patrio.commit()

            # Fechar o cursor e a conexão
            cursor_Patrio.close()
            conn_Patrio.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"{letras}_PATRIO_{cpf_cnpj}\n Concluido!", root)
            progress_var.set(60)

        except pyodbc.Error as ex:
            print_status(f"Erro ao executar a query pós restore para {letras}_PATRIO_{cpf_cnpj}: {ex}")

            atualizar_status(f"Erro ao executar a query pós restore para {letras}_PATRIO_{cpf_cnpj}: {ex}",root, fg="red")

        # Executar o comando de restore para AG
        try:
            subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_Ag], shell=True, check=True)
            print_status(f"Restore do banco\n{letras}_AG_{cpf_cnpj}\nexecutado com sucesso.")
            progress_var.set(70)  # Atualizar o valor da barra de progresso
        except subprocess.CalledProcessError as ex:
            print(f"Erro no restore do banco {letras}_AG_{cpf_cnpj}: {ex}")

        # Verificar se o banco AG foi restaurado com sucesso antes de executar a query pós-restore
        if os.path.isfile(caminho_dados_Ag) and os.path.isfile(caminho_log_Ag):
            # Alterar o modelo de recuperação para "Simples" para AG
            try:
                # Comando SQL para alterar o modelo de recuperação para "Simples" para AG
                alter_recovery_model_Ag = f'sqlcmd -S {server} -d {banco_de_dados_destino_Ag} -Q "ALTER DATABASE {banco_de_dados_destino_Ag} SET RECOVERY SIMPLE;"'
                subprocess.run(alter_recovery_model_Ag, check=True, shell=True)
                print_status(f"Modelo de recuperação alterado para 'Simples' para {letras}_AG_{cpf_cnpj}.")
            except subprocess.CalledProcessError as ex:
                print(f"Erro ao alterar o modelo de recuperação para 'Simples' para {letras}_AG_{cpf_cnpj}: {ex}")

            # Executar a query pós restore para AG
            try:
                # Estabelecer a conexão
                conn_AG = pyodbc.connect(conn_str)

                # Criar um objeto cursor
                cursor_AG = conn_AG.cursor()

                # Executar USE master;
                cursor_AG.execute("USE master;")

                # Comando SQL para realizar a query pós restore para AG
                query_pos_restore_Ag = f'''
                USE {letras_minusculas}_ag_{cpf_cnpj}
                CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
                EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

                '''

                # Executar a query pós restore para AG
                cursor_AG.execute(query_pos_restore_Ag)
                print_status(f"Query pós restore para\n{letras}_AG_{cpf_cnpj}\nexecutada com sucesso.")
                atualizar_status(f"Query pós restore para\n{letras}_AG_{cpf_cnpj}\nexecutada com sucesso.",root)
                progress_var.set(100)

                # Commit para confirmar as alterações
                conn_AG.commit()

                # Fechar o cursor e a conexão
                cursor_AG.close()
                conn_AG.close()

                # Adicione feedback visual ao usuário
                atualizar_status(f"{letras}_AG_{cpf_cnpj}\n Concluido!",root)
                progress_var.set(100)

            except pyodbc.Error as ex:
                print_status(f"Erro ao executar a query pós restore para {letras}_AG_{cpf_cnpj}: {ex}")

                atualizar_status(f"Erro ao executar a query pós restore para {letras}_AG_{cpf_cnpj}: {ex}",root, fg="red")

    else:
        print_status(f"Banco {letras}_AG_{cpf_cnpj} não foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    executaondemand_bat_ac_patrio_ag(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

def ac_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    progress_var.set(0)

    print_status(f'Iniciando restore de AC e AG em MSSQL...')
    progress_var.set(10)  # Atualizar o valor da barra de progresso
    atualizar_status(f'Iniciando restore de AC e AG em MSSQL...',root)

    nome_pasta = nome_pasta_entry.get()
    letras = letras_entry.get().upper()
    letras_minusculas = letras_entry.get().lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Parâmetros de conexão com o SQL Server
    servers = ['FOR-CNT-BD-SQL5', 
               'FORT-CNT-BDSQ-6',
               'GLAUDSON-DESKTO\\SQLEXPRESS']
    trusted_connection = 'yes'  # Configurando para autenticação do Windows
    driver = '{SQL Server}'

    # Parâmetros específicos para o restore
    banco_de_dados_destino_AC = f'{letras}_AC_{cpf_cnpj}'
    caminho_dados_AC = f'D:\\BDS\\dados\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.mdf'
    caminho_log_AC = f'D:\\BDS\\Logs\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.ldf'

    banco_de_dados_destino_AG = f'{letras}_AG_{cpf_cnpj}'
    caminho_dados_AG = f'D:\\BDS\\dados\\{nome_pasta}\\AG\\{letras}_AG_{cpf_cnpj}.mdf'
    caminho_log_AG = f'D:\\BDS\\Logs\\{nome_pasta}\\AG\\{letras}_AG_{cpf_cnpj}.ldf'

    # Verificar se os bancos de dados já existem
    databases_to_check = [banco_de_dados_destino_AC, banco_de_dados_destino_AG]

    # Tentar conectar-se a cada servidor da lista
    conectado = False
    conn = None
    for server in servers:
        try:
            # Criar a string de conexão
            conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

            # Tentar estabelecer a conexão
            conn = pyodbc.connect(conn_str)

            # Se a conexão for bem sucedida, sair do loop
            print_status(f"Conectado ao servidor: {server}")
            progress_var.set(30)  # Atualizar o valor da barra de progresso
            atualizar_status(f"Conectado ao servidor: {server}", root)
            conectado = True
            break
        except Exception as e:
            print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
            atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

    if not conectado:
        print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
        atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
        return

    try:
        conn_check = pyodbc.connect(conn_str)
        cursor_check = conn_check.cursor()

        # Verificar se os bancos de dados já existem
        query_check = "SELECT name FROM sys.databases WHERE name IN (?, ?)"
        cursor_check.execute(query_check, databases_to_check)
        existing_databases = [row[0] for row in cursor_check.fetchall()]

        if existing_databases:
            print('Os bancos de dados já existem.\n Não é possível continuar o restore.')
            print_status("Os bancos de dados já existem.\n Não é possível continuar o restore.")
            atualizar_status("Os bancos de dados já existem.\n Não é possível continuar o restore.", root)
            progress_var.set(30)


            tente_novamente_ac_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
            return

    except Exception as e:
        print(f"Erro ao verificar a existência dos bancos de dados: {e}")
        atualizar_status(f"Erro ao verificar a existência dos bancos de dados: {e}", root)
        return

    finally:
        if conn_check:
            conn_check.close()

    # Comando SQL para realizar o restore usando caminhos fixos para AC
    restore_query_AC = f'RESTORE DATABASE {banco_de_dados_destino_AC} FROM DISK = \'D:\\Bancos Limpos SQL\\AC.bak\' WITH REPLACE, MOVE \'AC_Data\' TO \'{caminho_dados_AC}\', MOVE \'AC_Log\' TO \'{caminho_log_AC}\''

    # Comando SQL para realizar o restore usando caminhos fixos para AG
    restore_query_AG = f'RESTORE DATABASE {banco_de_dados_destino_AG} FROM DISK = \'D:\\Bancos Limpos SQL\\AG.bak\' WITH REPLACE, MOVE \'AG_Data\' TO \'{caminho_dados_AG}\', MOVE \'AG_Log\' TO \'{caminho_log_AG}\''

    atualizar_status(f'Restore em Andamento...',root)
    progress_var.set(30)  # Atualizar o valor da barra de progresso

    # Executar o comando de restore para AC
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AC], shell=True, check=True)
        print_status(f"Restore do banco {letras}_AC_{cpf_cnpj} executado com sucesso.")
    except subprocess.CalledProcessError as ex:
        print(f"Erro no restore do banco {letras}_AC_{cpf_cnpj}: {ex}")

    # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_AC) and os.path.isfile(caminho_log_AC):
        # Alterar o modelo de recuperação para "Simples" para AC
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para AC
            alter_recovery_model_AC = f'sqlcmd -S {server} -d {banco_de_dados_destino_AC} -Q "ALTER DATABASE {banco_de_dados_destino_AC} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_AC, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AC_{cpf_cnpj}.")
            atualizar_status(f"Modelo de recuperação alterado para 'Simples'\n Em para {letras}_AC_{cpf_cnpj}.",root)
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}",root)

        # Executar a query pós restore para AC
        try:
            # Estabelecer a conexão
            conn_AC = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_AC = conn_AC.cursor()

            # Executar USE master;
            cursor_AC.execute("USE master;")

            # Comando SQL para realizar a query pós restore para AC
            query_pos_restore_AC = f'''
            USE {letras_minusculas}_ac_{cpf_cnpj}
            CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ac_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
            '''

            # Executar a query pós restore para AC
            cursor_AC.execute(query_pos_restore_AC)
            print_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.",root)

            # Commit para confirmar as alterações
            conn_AC.commit()

            # Fechar o cursor e a conexão
            cursor_AC.close()
            conn_AC.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"{letras}_AC_{cpf_cnpj}\n Concluido!", root)
            progress_var.set(60)  # Atualizar o valor da barra de progresso

        except pyodbc.Error as ex:
            print(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}",root)

    else:
        print_status(f"Banco {letras}_AC_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    # Executar o comando de restore para PATRIO
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AG], shell=True, check=True)
        print_status(f"Restore do banco\n{letras}_AG_{cpf_cnpj}\nexecutado com sucesso.")
        progress_var.set(70)  # Atualizar o valor da barra de progresso
    except subprocess.CalledProcessError as e:
        print(f"Erro no restore do banco {letras}_AG_{cpf_cnpj}: {e}")

    # Verificar se o banco PATRIO foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_AG) and os.path.isfile(caminho_log_AG):
        # Alterar o modelo de recuperação para "Simples" para PATRIO
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para PATRIO
            alter_recovery_model_AG = f'sqlcmd -S {server} -d {banco_de_dados_destino_AG} -Q "ALTER DATABASE {banco_de_dados_destino_AG} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_AG, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples' para {letras}_AG_{cpf_cnpj}.")
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples' para {letras}_AG_{cpf_cnpj}: {ex}")

        # Executar a query pós restore para AG
        try:
            # Estabelecer a conexão
            conn_AG = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_AG = conn_AG.cursor()

            # Executar USE master;
            cursor_AG.execute("USE master;")

            # Comando SQL para realizar a query pós restore para PATRIO
            query_pos_restore_AG = f'''
            USE {letras_minusculas}_ag_{cpf_cnpj}
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            '''

            # Executar a query pós restore para PATRIO
            cursor_AG.execute(query_pos_restore_AG)
            print_status(f"Query pós restore para\n{letras}_AG_{cpf_cnpj}\nexecutada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_AG_{cpf_cnpj}\nexecutada com sucesso.",root)
            progress_var.set(100)

            # Commit para confirmar as alterações
            conn_AG.commit()

            # Fechar o cursor e a conexão
            cursor_AG.close()
            conn_AG.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"{letras}_AG_{cpf_cnpj}\n Concluido!",root)
            progress_var.set(100)

        except pyodbc.Error as ex:
            print_status(f"Erro ao executar a query pós restore para {letras}_AG_{cpf_cnpj}: {ex}")

            atualizar_status(f"Erro ao executar a query pós restore para {letras}_AG_{cpf_cnpj}: {ex}",root, fg="red")


    else:
        print_status(f"Banco {letras}_AG_{cpf_cnpj} não foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    executaondemand_bat_ac_ag(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

def Apenas_AG(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

        # Configurar a barra de progresso
        progress_var.set(0)

        print_status(f'Iniciando restore de AG em MSSQL...')
        progress_var.set(10)  # Atualizar o valor da barra de progresso
        atualizar_status(f'Iniciando restore de AG em MSSQL...',root)

        nome_pasta = nome_pasta_entry.get()
        letras_minusculas = letras_entry.get().lower()
        cpf_cnpj = cpf_cnpj_entry.get()
        letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
        senha = senha_cliente.get()

        # Parâmetros de conexão com o SQL Server
        servers = ['FOR-CNT-BD-SQL5', 
                   'FORT-CNT-BDSQ-6',
                   'GLAUDSON-DESKTO\\SQLEXPRESS']
        trusted_connection = 'yes'  # Configurando para autenticação do Windows
        driver = '{SQL Server}'

        # Parâmetros específicos para o restore
        banco_de_dados_destino_AG = f'{letras}_AG_{cpf_cnpj}'
        caminho_dados_AG = f'D:\\BDS\\dados\\{nome_pasta}\\AG\\{letras}_AG_{cpf_cnpj}.mdf'
        caminho_log_AG = f'D:\\BDS\\Logs\\{nome_pasta}\\AG\\{letras}_AC_{cpf_cnpj}.ldf'

        # Verificar se os bancos de dados já existem
        databases_to_check = [banco_de_dados_destino_AG]

        # Tentar conectar-se a cada servidor da lista
        conectado = False
        conn = None
        for server in servers:
            try:
                # Criar a string de conexão
                conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

                # Tentar estabelecer a conexão
                conn = pyodbc.connect(conn_str)

                # Se a conexão for bem sucedida, sair do loop
                print_status(f"Conectado ao servidor: {server}")
                progress_var.set(30)  # Atualizar o valor da barra de progresso
                atualizar_status(f"Conectado ao servidor: {server}", root)
                conectado = True
                break
            except Exception as e:
                print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
                atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

        if not conectado:
            print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
            atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
            return

        try:
            conn_check = pyodbc.connect(conn_str)
            cursor_check = conn_check.cursor()

            # Verificar se os bancos de dados já existem
            query_check = "SELECT name FROM sys.databases WHERE name IN (?)"
            cursor_check.execute(query_check, databases_to_check)
            existing_databases = [row[0] for row in cursor_check.fetchall()]

            if existing_databases:
                print('O banco de dados já existem. \nNão é possível continuar o restore.')
                print_status("Os bancos de dados já existem.\nNão é possível continuar o restore.")
                atualizar_status("Os bancos de dados já existem.\n Não é possível continuar o restore.", root)
                progress_var.set(30)


                tente_novamente_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
                return

        except Exception as e:
            print(f"Erro ao verificar a existência dos bancos de dados: {e}")
            atualizar_status(f"Erro ao verificar a existência dos bancos de dados: {e}", root)
            return

        finally:
            if conn_check:
                conn_check.close()


        # Comando SQL para realizar o restore usando caminhos fixos para AG
        restore_query_AG = f'RESTORE DATABASE {banco_de_dados_destino_AG} FROM DISK = \'D:\\Bancos Limpos SQL\\AG.bak\' WITH REPLACE, MOVE \'AG_Data\' TO \'{caminho_dados_AG}\', MOVE \'AG_Log\' TO \'{caminho_log_AG}\''

        atualizar_status(f'Restore em andamento...',root)
        progress_var.set(30)  # Atualizar o valor da barra de progresso


        # Executar o comando de restore para AG
        try:
            subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AG], shell=True, check=True)
            print_status(f"Restore do banco {letras}_AG_{cpf_cnpj}\n executado com sucesso.")
            atualizar_status(f"Restore do banco {letras}_AG_{cpf_cnpj}\n executado com sucesso.", root)
            progress_var.set(50)  # Atualizar o valor da barra de progresso
        except subprocess.CalledProcessError as e:
            print_status(f"Erro no restore do banco {letras}_AG_{cpf_cnpj}:\n {e}")
            atualizar_status(f"Erro no restore do banco {letras}_AG_{cpf_cnpj}:\n {e}", root)

        # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
        if os.path.isfile(caminho_dados_AG) and os.path.isfile(caminho_log_AG):
            # Alterar o modelo de recuperação para "Simples" para AG
            try:
                # Comando SQL para alterar o modelo de recuperação para "Simples" para AG
                alter_recovery_model_AG = f'sqlcmd -S {server} -d {banco_de_dados_destino_AG} -Q "ALTER DATABASE {banco_de_dados_destino_AG} SET RECOVERY SIMPLE;"'
                subprocess.run(alter_recovery_model_AG, check=True, shell=True)
                print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AG_{cpf_cnpj}.")
                atualizar_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AG_{cpf_cnpj}.", root)
            except subprocess.CalledProcessError as ex:
                print_status(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AG_{cpf_cnpj}: {ex}")
                atualizar_status(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AG_{cpf_cnpj}: {ex}", root)

            # Executar a query pós restore para AG
            try:
                # Estabelecer a conexão
                conn_AG = pyodbc.connect(conn_str)

                # Criar um objeto cursor
                cursor_AG = conn_AG.cursor()

                # Executar USE master;
                cursor_AG.execute("USE master;")

                # Comando SQL para realizar a query pós restore para AG
                query_pos_restore_AG = f'''
                USE {letras_minusculas}_ag_{cpf_cnpj}
                CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ag_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
                CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
                EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

                INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
                '''

                # Executar a query pós restore para AG
                cursor_AG.execute(query_pos_restore_AG)
                print_status(f"Query pós restore para {letras}_AG_{cpf_cnpj}\n executada com sucesso.")

                atualizar_status(f"Query pós restore para {letras}_AG_{cpf_cnpj}\n executada com sucesso.", root)
                progress_var.set(80)  # Atualizar o valor da barra de progresso

                # Commit para confirmar as alterações
                conn_AG.commit()

                # Fechar o cursor e a conexão
                cursor_AG.close()
                conn_AG.close()

                # Adicione feedback visual ao usuário
                atualizar_status(f"{letras}_AG_{cpf_cnpj}\n Concluido!", root)
                progress_var.set(100)

            except pyodbc.Error as ex:
                print(f"Erro ao executar a query pós restore para {letras}_AG_{cpf_cnpj}: {ex}")
                atualizar_status(f"Erro ao executar a query pós restore para\n {letras}_AG_{cpf_cnpj}: {ex}",root)


        else:
            print_status(f"Banco {letras}_AG_{cpf_cnpj} não foi restaurado com sucesso.\n Consulte os logs para mais detalhes.")
            atualizar_status(f"Banco {letras}_AG_{cpf_cnpj} não foi restaurado com sucesso.\n Consulte os logs para mais detalhes.",root)

        executaondemand_bat_ag(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

def Apenas_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    # Configurar a barra de progresso
    progress_var.set(0)

    print_status(f'Iniciando restore de PONTO em MSSQL...')
    progress_var.set(10)  # Atualizar o valor da barra de progresso
    atualizar_status(f'Iniciando restore\nde PONTO em MSSQL...',root)

    nome_pasta = nome_pasta_entry.get()
    letras_minusculas = letras_entry.get().lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    letras = letras_entry.get().upper()  # Converte as letras para maiúsculas
    senha = senha_cliente.get()

    # Parâmetros de conexão com o SQL Server
    servers = ['FOR-CNT-BD-SQL5', 
               'FORT-CNT-BDSQ-6',
               'GLAUDSON-DESKTO\\SQLEXPRESS']
    trusted_connection = 'yes'  # Configurando para autenticação do Windows
    driver = '{SQL Server}'


    # Parâmetros específicos para o restore
    banco_de_dados_destino_ponto = f'{letras}_PONTO_{cpf_cnpj}'
    caminho_dados_ponto = f'D:\\BDS\\dados\\{nome_pasta}\\PONTO\\{letras}_PONTO_{cpf_cnpj}.mdf'
    caminho_log_ponto = f'D:\\BDS\\Logs\\{nome_pasta}\\PONTO\\{letras}_PONTO_{cpf_cnpj}.ldf'

    # Verificar se os bancos de dados já existem
    databases_to_check = [banco_de_dados_destino_ponto]

    # Tentar conectar-se a cada servidor da lista
    conectado = False
    conn = None
    for server in servers:
        try:
            # Criar a string de conexão
            conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

            # Tentar estabelecer a conexão
            conn = pyodbc.connect(conn_str)

            # Se a conexão for bem sucedida, sair do loop
            print_status(f"Conectado ao servidor: {server}")
            progress_var.set(30)  # Atualizar o valor da barra de progresso
            atualizar_status(f"Conectado ao servidor: {server}", root)
            conectado = True
            break
        except Exception as e:
            print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
            atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

    if not conectado:
        print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
        atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
        return
        
    try:
        conn_check = pyodbc.connect(conn_str)
        cursor_check = conn_check.cursor()

        # Verificar se os bancos de dados já existem
        query_check = "SELECT name FROM sys.databases WHERE name IN (?)"
        cursor_check.execute(query_check, databases_to_check)
        existing_databases = [row[0] for row in cursor_check.fetchall()]

        if existing_databases:
            print('O banco de dados já existem. \nNão é possível continuar o restore.')
            print_status("Os bancos de dados já existem.\nNão é possível continuar o restore.")
            atualizar_status("Os bancos de dados já existem.\nNão é possível continuar o restore.", root)
            progress_var.set(30)


            tente_novamente_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
            return

    except Exception as e:
        print(f"Erro ao verificar a existência dos bancos de dados: {e}")
        atualizar_status(f"Erro ao verificar a\nexistência dos bancos de dados: {e}", root)
        return

    finally:
        if conn_check:
            conn_check.close()


    # Comando SQL para realizar o restore usando caminhos fixos para PONTO
    restore_query_ponto = f'RESTORE DATABASE {banco_de_dados_destino_ponto} FROM DISK = \'D:\\Bancos Limpos SQL\\PONTO.bak\' WITH REPLACE, MOVE \'PONTO3\' TO \'{caminho_dados_ponto}\', MOVE \'PONTO3_Log\' TO \'{caminho_log_ponto}\''

    atualizar_status(f'Restore em andamento...',root)
    progress_var.set(30)  # Atualizar o valor da barra de progresso


    # Executar o comando de restore para AC
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_ponto], shell=True, check=True)
        print_status(f"Restore do banco {letras}_PONTO_{cpf_cnpj}\n executado com sucesso.")
        atualizar_status(f"Restore do banco\n{letras}_PONTO_{cpf_cnpj}\nexecutado com sucesso.", root)
        progress_var.set(50)  # Atualizar o valor da barra de progresso
    except subprocess.CalledProcessError as ex:
        print_status(f"Erro no restore do banco {letras}_PONTO_{cpf_cnpj}:\n {e}")
        atualizar_status(f"Erro no restore do banco\n{letras}_PONTO_{cpf_cnpj}:\n {e}", root)

    # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_ponto) and os.path.isfile(caminho_log_ponto):
        # Alterar o modelo de recuperação para "Simples" para PONTO
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para AC
            alter_recovery_model_ponto = f'sqlcmd -S {server} -d {banco_de_dados_destino_ponto} -Q "ALTER DATABASE {banco_de_dados_destino_ponto} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_ponto, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_PONTO_{cpf_cnpj}.")
            atualizar_status(f"Modelo de recuperação\nalterado para 'Simples'\nEm {letras}_PONTO_{cpf_cnpj}.", root)
        except subprocess.CalledProcessError as ex:
            print_status(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_PONTO_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao alterar o modelo de\nrecuperação para 'Simples'\nEm {letras}_PONTO_{cpf_cnpj}: {ex}", root)

        # Executar a query pós restore para PONTO
        try:
            # Estabelecer a conexão
            conn_ponto = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_ponto = conn_ponto.cursor()

            # Executar USE master;
            cursor_ponto.execute("USE master;")

            # Comando SQL para realizar a query pós restore para PONTO
            query_pos_restore_ponto = f'''
            USE {letras_minusculas}_ponto_{cpf_cnpj}
            CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ponto_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
            '''

            # Executar a query pós restore para AC
            cursor_ponto.execute(query_pos_restore_ponto)
            print_status(f"Query pós restore para {letras}_PONTO_{cpf_cnpj}\n executada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_PONTO_{cpf_cnpj}\nexecutada com sucesso.", root)
            progress_var.set(80)  # Atualizar o valor da barra de progresso

            # Commit para confirmar as alterações
            conn_ponto.commit()

            # Fechar o cursor e a conexão
            cursor_ponto.close()
            conn_ponto.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"Concluido!\nRestore executado\ncom sucesso.",root)
            progress_var.set(100)
        except pyodbc.Error as ex:
            print(f"Erro ao executar a query pós restore para {letras}_PONTO_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao executar\na query pós restore para\n{letras}_PONTO_{cpf_cnpj}: {ex}",root)


    else:
        print_status(f"Banco {letras}_PONTO_{cpf_cnpj} não foi restaurado com sucesso.\n Consulte os logs para mais detalhes.")
        atualizar_status(f"Banco {letras}_PONTO_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.",root)

    executaondemand_bat_ponto(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

def ac_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    progress_var.set(0)

    print_status(f'Iniciando restore de AC e PONTO em MSSQL...')
    progress_var.set(10)  # Atualizar o valor da barra de progresso
    atualizar_status(f'Iniciando restore de AC e PONTO em MSSQL...',root)

    nome_pasta = nome_pasta_entry.get()
    letras = letras_entry.get().upper()
    letras_minusculas = letras_entry.get().lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Parâmetros de conexão com o SQL Server
    servers = ['FOR-CNT-BD-SQL5', 
               'FORT-CNT-BDSQ-6',
               'GLAUDSON-DESKTO\\SQLEXPRESS']
    trusted_connection = 'yes'  # Configurando para autenticação do Windows
    driver = '{SQL Server}'

    # Parâmetros específicos para o restore
    banco_de_dados_destino_AC = f'{letras}_AC_{cpf_cnpj}'
    caminho_dados_AC = f'D:\\BDS\\dados\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.mdf'
    caminho_log_AC = f'D:\\BDS\\Logs\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.ldf'

    banco_de_dados_destino_ponto = f'{letras}_PONTO_{cpf_cnpj}'
    caminho_dados_ponto = f'D:\\BDS\\dados\\{nome_pasta}\\PONTO\\{letras}_PONTO_{cpf_cnpj}.mdf'
    caminho_log_ponto = f'D:\\BDS\\Logs\\{nome_pasta}\\PONTO\\{letras}_PONTO_{cpf_cnpj}.ldf'

    # Verificar se os bancos de dados já existem
    databases_to_check = [banco_de_dados_destino_AC, banco_de_dados_destino_ponto]

    # Tentar conectar-se a cada servidor da lista
    conectado = False
    conn = None
    for server in servers:
        try:
            # Criar a string de conexão
            conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

            # Tentar estabelecer a conexão
            conn = pyodbc.connect(conn_str)

            # Se a conexão for bem sucedida, sair do loop
            print_status(f"Conectado ao servidor: {server}")
            progress_var.set(30)  # Atualizar o valor da barra de progresso
            atualizar_status(f"Conectado ao servidor: {server}", root)
            conectado = True
            break
        except Exception as e:
            print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
            atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

    if not conectado:
        print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
        atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
        return

    try:
        conn_check = pyodbc.connect(conn_str)
        cursor_check = conn_check.cursor()

        # Verificar se os bancos de dados já existem
        query_check = "SELECT name FROM sys.databases WHERE name IN (?, ?)"
        cursor_check.execute(query_check, databases_to_check)
        existing_databases = [row[0] for row in cursor_check.fetchall()]

        if existing_databases:
            print('Os bancos de dados já existem.\n Não é possível continuar o restore.')
            print_status("Os bancos de dados já existem.\n Não é possível continuar o restore.")
            atualizar_status("Os bancos de dados já existem.\n Não é possível continuar o restore.", root)
            progress_var.set(30)


            tente_novamente_ac_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
            return

    except Exception as e:
        print(f"Erro ao verificar a existência dos bancos de dados: {e}")
        atualizar_status(f"Erro ao verificar a existência dos bancos de dados: {e}", root)
        return

    finally:
        if conn_check:
            conn_check.close()

    # Comando SQL para realizar o restore usando caminhos fixos para AC
    restore_query_AC = f'RESTORE DATABASE {banco_de_dados_destino_AC} FROM DISK = \'D:\\Bancos Limpos SQL\\AC.bak\' WITH REPLACE, MOVE \'AC_Data\' TO \'{caminho_dados_AC}\', MOVE \'AC_Log\' TO \'{caminho_log_AC}\''

    # Comando SQL para realizar o restore usando caminhos fixos para PONTO
    restore_query_ponto = f'RESTORE DATABASE {banco_de_dados_destino_ponto} FROM DISK = \'D:\\Bancos Limpos SQL\\PONTO.bak\' WITH REPLACE, MOVE \'PONTO3\' TO \'{caminho_dados_ponto}\', MOVE \'PONTO3_Log\' TO \'{caminho_log_ponto}\''

    atualizar_status(f'Restore em Andamento...',root)
    progress_var.set(30)  # Atualizar o valor da barra de progresso

    # Executar o comando de restore para AC
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AC], shell=True, check=True)
        print_status(f"Restore do banco {letras}_AC_{cpf_cnpj} executado com sucesso.")
    except subprocess.CalledProcessError as ex:
        print(f"Erro no restore do banco {letras}_AC_{cpf_cnpj}: {ex}")

    # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_AC) and os.path.isfile(caminho_log_AC):
        # Alterar o modelo de recuperação para "Simples" para AC
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para AC
            alter_recovery_model_AC = f'sqlcmd -S {server} -d {banco_de_dados_destino_AC} -Q "ALTER DATABASE {banco_de_dados_destino_AC} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_AC, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AC_{cpf_cnpj}.")
            atualizar_status(f"Modelo de recuperação alterado para 'Simples'\n Em para {letras}_AC_{cpf_cnpj}.",root)
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}",root)

        # Executar a query pós restore para AC
        try:
            # Estabelecer a conexão
            conn_AC = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_AC = conn_AC.cursor()

            # Executar USE master;
            cursor_AC.execute("USE master;")

            # Comando SQL para realizar a query pós restore para AC
            query_pos_restore_AC = f'''
            USE {letras_minusculas}_ac_{cpf_cnpj}
            CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ac_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
            '''

            # Executar a query pós restore para AC
            cursor_AC.execute(query_pos_restore_AC)
            print_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.",root)

            # Commit para confirmar as alterações
            conn_AC.commit()

            # Fechar o cursor e a conexão
            cursor_AC.close()
            conn_AC.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"{letras}_AC_{cpf_cnpj}\n Concluido!", root)
            progress_var.set(60)  # Atualizar o valor da barra de progresso

        except pyodbc.Error as ex:
            print(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}",root)

    else:
        print_status(f"Banco {letras}_AC_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    # Executar o comando de restore para PONTO
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_ponto], shell=True, check=True)
        print_status(f"Restore do banco\n{letras}_PONTO_{cpf_cnpj}\nexecutado com sucesso.")
        progress_var.set(70)  # Atualizar o valor da barra de progresso
    except subprocess.CalledProcessError as e:
        print(f"Erro no restore do banco {letras}_PONTO_{cpf_cnpj}: {e}")

    # Verificar se o banco PONTO foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_ponto) and os.path.isfile(caminho_log_ponto):
        # Alterar o modelo de recuperação para "Simples" para PONTO
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para PONTO
            alter_recovery_model_ponto = f'sqlcmd -S {server} -d {banco_de_dados_destino_ponto} -Q "ALTER DATABASE {banco_de_dados_destino_ponto} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_ponto, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples' para {letras}_PONTO_{cpf_cnpj}.")
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples' para {letras}_PONTO_{cpf_cnpj}: {ex}")

        # Executar a query pós restore para PONTO
        try:
            # Estabelecer a conexão
            conn_ponto = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_ponto = conn_ponto.cursor()

            # Executar USE master;
            cursor_ponto.execute("USE master;")

            # Comando SQL para realizar a query pós restore para PONTO
            query_pos_restore_ponto = f'''
            USE {letras_minusculas}_ponto_{cpf_cnpj}
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            '''

            # Executar a query pós restore para PONTO
            cursor_ponto.execute(query_pos_restore_ponto)
            print_status(f"Query pós restore para\n{letras}_PONTO_{cpf_cnpj}\nexecutada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_PONTO_{cpf_cnpj}\nexecutada com sucesso.",root)
            progress_var.set(100)

            # Commit para confirmar as alterações
            conn_ponto.commit()

            # Fechar o cursor e a conexão
            cursor_ponto.close()
            conn_ponto.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"{letras}_PONTO_{cpf_cnpj}\n Concluido!",root)
            progress_var.set(100)

        except pyodbc.Error as ex:
            print_status(f"Erro ao executar a query pós restore para {letras}_PONTO_{cpf_cnpj}: {ex}")

            atualizar_status(f"Erro ao executar a query pós restore para {letras}_PONTO_{cpf_cnpj}: {ex}",root, fg="red")


    else:
        print_status(f"Banco {letras}_PONTO_{cpf_cnpj} não foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    executaondemand_bat_ac_ponto(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

def ac_patrio_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto):

    progress_var.set(0)

    print_status(f'Iniciando restore\nde AC, PONTO e PATRIO em MSSQL...')
    progress_var.set(10)  # Atualizar o valor da barra de progresso
    atualizar_status(f'Iniciando restore\nde AC, PONTO e PATRIO em MSSQL...',root)

    nome_pasta = nome_pasta_entry.get()
    letras = letras_entry.get().upper()
    letras_minusculas = letras_entry.get().lower()
    cpf_cnpj = cpf_cnpj_entry.get()
    senha = senha_cliente.get()

    # Parâmetros de conexão com o SQL Server
    servers = ['FOR-CNT-BD-SQL5', 
               'FORT-CNT-BDSQ-6',
               'GLAUDSON-DESKTO\\SQLEXPRESS']
    trusted_connection = 'yes'  # Configurando para autenticação do Windows
    driver = '{SQL Server}'

    # Parâmetros específicos para o restore
    banco_de_dados_destino_AC = f'{letras}_AC_{cpf_cnpj}'
    caminho_dados_AC = f'D:\\BDS\\dados\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.mdf'
    caminho_log_AC = f'D:\\BDS\\Logs\\{nome_pasta}\\AC\\{letras}_AC_{cpf_cnpj}.ldf'

    banco_de_dados_destino_Patrio = f'{letras}_PATRIO_{cpf_cnpj}'
    caminho_dados_Patrio = f'D:\\BDS\\dados\\{nome_pasta}\\PATRIO\\{letras}_PATRIO_{cpf_cnpj}.mdf'
    caminho_log_Patrio = f'D:\\BDS\\Logs\\{nome_pasta}\\PATRIO\\{letras}_PATRIO_{cpf_cnpj}.ldf'

    banco_de_dados_destino_ponto = f'{letras}_PONTO_{cpf_cnpj}'
    caminho_dados_ponto = f'D:\\BDS\\dados\\{nome_pasta}\\PONTO\\{letras}_PONTO_{cpf_cnpj}.mdf'
    caminho_log_ponto = f'D:\\BDS\\Logs\\{nome_pasta}\\PONTO\\{letras}_PONTO_{cpf_cnpj}.ldf'

    # Verificar se os bancos de dados já existem
    databases_to_check = [banco_de_dados_destino_AC, banco_de_dados_destino_Patrio, banco_de_dados_destino_ponto]

    # Tentar conectar-se a cada servidor da lista
    conectado = False
    conn = None
    for server in servers:
        try:
            # Criar a string de conexão
            conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

            # Tentar estabelecer a conexão
            conn = pyodbc.connect(conn_str)

            # Se a conexão for bem sucedida, sair do loop
            print_status(f"Conectado ao servidor: {server}")
            progress_var.set(30)  # Atualizar o valor da barra de progresso
            atualizar_status(f"Conectado ao servidor: {server}", root)
            conectado = True
            break
        except Exception as e:
            print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
            atualizar_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}", root)

    if not conectado:
        print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
        atualizar_status("Não foi possível conectar-se a nenhum dos servidores da lista.", root)
        return

    try:
        conn_check = pyodbc.connect(conn_str)
        cursor_check = conn_check.cursor()

        # Verificar se os bancos de dados já existem
        query_check = "SELECT name FROM sys.databases WHERE name IN (?, ?, ?)"
        cursor_check.execute(query_check, databases_to_check)
        existing_databases = [row[0] for row in cursor_check.fetchall()]

        if existing_databases:
            print('Os bancos de dados já existem.\n Não é possível continuar o restore.')
            print_status("Os bancos de dados já existem.\n Não é possível continuar o restore.")
            atualizar_status("Os bancos de dados já existem.\n Não é possível continuar o restore.", root)
            progress_var.set(30)


            tente_novamente_ac_patrio_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar,progress_var)
            return

    except Exception as ex:
        print(f"Erro ao verificar a existência\ndos bancos de dados:\n{ex}")
        atualizar_status(f"Erro ao verificar a existência\ndos bancos de dados:\n{ex}", root)
        return

    finally:
        if conn_check:
            conn_check.close()

    # Comando SQL para realizar o restore usando caminhos fixos para AC
    restore_query_AC = f'RESTORE DATABASE {banco_de_dados_destino_AC} FROM DISK = \'D:\\Bancos Limpos SQL\\AC.bak\' WITH REPLACE, MOVE \'AC_Data\' TO \'{caminho_dados_AC}\', MOVE \'AC_Log\' TO \'{caminho_log_AC}\''

    # Comando SQL para realizar o restore usando caminhos fixos para Patrio
    restore_query_Patrio = f'RESTORE DATABASE {banco_de_dados_destino_Patrio} FROM DISK = \'D:\\Bancos Limpos SQL\\Patrio.bak\' WITH REPLACE, MOVE \'Patrio\' TO \'{caminho_dados_Patrio}\', MOVE \'Patrio_Log\' TO \'{caminho_log_Patrio}\''

    # Comando SQL para realizar o restore usando caminhos fixos para PONTO
    restore_query_ponto = f'RESTORE DATABASE {banco_de_dados_destino_ponto} FROM DISK = \'D:\\Bancos Limpos SQL\\PONTO.bak\' WITH REPLACE, MOVE \'PONTO3\' TO \'{caminho_dados_ponto}\', MOVE \'PONTO3_Log\' TO \'{caminho_log_ponto}\''

    atualizar_status(f'Restore em Andamento...',root)
    progress_var.set(30)  # Atualizar o valor da barra de progresso

    # Executar o comando de restore para AC
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_AC], shell=True, check=True)
        print_status(f"Restore do banco {letras}_AC_{cpf_cnpj} executado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro no restore do banco {letras}_AC_{cpf_cnpj}: {e}")

    # Verificar se o banco AC foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_AC) and os.path.isfile(caminho_log_AC):
        # Alterar o modelo de recuperação para "Simples" para AC
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para AC
            alter_recovery_model_AC = f'sqlcmd -S {server} -d {banco_de_dados_destino_AC} -Q "ALTER DATABASE {banco_de_dados_destino_AC} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_AC, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_AC_{cpf_cnpj}.")
            atualizar_status(f"Modelo de recuperação\nalterado para 'Simples'\nEm para {letras}_AC_{cpf_cnpj}.",root)
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples'\n Em {letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao alterar o modelo\nde recuperação para 'Simples'\nEm {letras}_AC_{cpf_cnpj}: {ex}",root)

        # Executar a query pós restore para AC
        try:
            # Estabelecer a conexão
            conn_AC = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_AC = conn_AC.cursor()

            # Executar USE master;
            cursor_AC.execute("USE master;")

            # Comando SQL para realizar a query pós restore para AC
            query_pos_restore_AC = f'''
            USE {letras_minusculas}_ac_{cpf_cnpj}
            CREATE LOGIN [{letras_minusculas}_{cpf_cnpj}.sql] WITH PASSWORD = '{senha}', DEFAULT_DATABASE = {letras_minusculas}_ac_{cpf_cnpj}, CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF;
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            INSERT INTO cfg (Codigo, Valor) VALUES ('USARAGENTENUVEM', 1);
            '''

            # Executar a query pós restore para AC
            cursor_AC.execute(query_pos_restore_AC)
            print_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\nexecutada com sucesso.",root)

            # Commit para confirmar as alterações
            conn_AC.commit()

            # Fechar o cursor e a conexão
            cursor_AC.close()
            conn_AC.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"Query pós restore para\n{letras}_AC_{cpf_cnpj}\n executada com sucesso.",root)
            atualizar_status(f"{letras}_AC_{cpf_cnpj}\nConcluido!", root)
            progress_var.set(40)  # Atualizar o valor da barra de progresso

        except pyodbc.Error as ex:
            print(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}")
            atualizar_status(f"Erro ao executar a query pós restore para\n{letras}_AC_{cpf_cnpj}: {ex}",root)

    else:
        print_status(f"Banco {letras}_AC_{cpf_cnpj}\nnão foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    # Executar o comando de restore para PATRIO
    try:
        subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_Patrio], shell=True, check=True)
        print_status(f"Restore do banco\n{letras}_PATRIO_{cpf_cnpj}\nexecutado com sucesso.")
        progress_var.set(45)  # Atualizar o valor da barra de progresso
    except subprocess.CalledProcessError as e:
        print(f"Erro no restore do banco {letras}_PATRIO_{cpf_cnpj}: {e}")

    # Verificar se o banco PATRIO foi restaurado com sucesso antes de executar a query pós-restore
    if os.path.isfile(caminho_dados_Patrio) and os.path.isfile(caminho_log_Patrio):
        # Alterar o modelo de recuperação para "Simples" para PATRIO
        try:
            # Comando SQL para alterar o modelo de recuperação para "Simples" para PATRIO
            alter_recovery_model_Patrio = f'sqlcmd -S {server} -d {banco_de_dados_destino_Patrio} -Q "ALTER DATABASE {banco_de_dados_destino_Patrio} SET RECOVERY SIMPLE;"'
            subprocess.run(alter_recovery_model_Patrio, check=True, shell=True)
            print_status(f"Modelo de recuperação alterado para 'Simples' para {letras}_PATRIO_{cpf_cnpj}.")
        except subprocess.CalledProcessError as ex:
            print(f"Erro ao alterar o modelo de recuperação para 'Simples' para {letras}_PATRIO_{cpf_cnpj}: {ex}")

        # Executar a query pós restore para PATRIO
        try:
            # Estabelecer a conexão
            conn_Patrio = pyodbc.connect(conn_str)

            # Criar um objeto cursor
            cursor_Patrio = conn_Patrio.cursor()

            # Executar USE master;
            cursor_Patrio.execute("USE master;")

            # Comando SQL para realizar a query pós restore para PATRIO
            query_pos_restore_Patrio = f'''
            USE {letras_minusculas}_patrio_{cpf_cnpj}
            CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
            EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
            EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

            '''

            # Executar a query pós restore para PATRIO
            cursor_Patrio.execute(query_pos_restore_Patrio)
            print_status(f"Query pós restore para\n{letras}_PATRIO_{cpf_cnpj}\nexecutada com sucesso.")
            atualizar_status(f"Query pós restore para\n{letras}_PATRIO_{cpf_cnpj}\nexecutada com sucesso.",root)
            progress_var.set(50)

            # Commit para confirmar as alterações
            conn_Patrio.commit()

            # Fechar o cursor e a conexão
            cursor_Patrio.close()
            conn_Patrio.close()

            # Adicione feedback visual ao usuário
            atualizar_status(f"{letras}_PATRIO_{cpf_cnpj}\n Concluido!", root)
            progress_var.set(60)

        except pyodbc.Error as ex:
            print_status(f"Erro ao executar a query pós restore para {letras}_PATRIO_{cpf_cnpj}: {ex}")

            atualizar_status(f"Erro ao executar a query pós restore para {letras}_PATRIO_{cpf_cnpj}: {ex}",root, fg="red")

        # Executar o comando de restore para PONTO
        try:
            subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query_ponto], shell=True, check=True)
            print_status(f"Restore do banco\n{letras}_PONTO_{cpf_cnpj}\nexecutado com sucesso.")
            progress_var.set(70)  # Atualizar o valor da barra de progresso
        except subprocess.CalledProcessError as ex:
            print(f"Erro no restore do banco {letras}_PONTO_{cpf_cnpj}: {ex}")

        # Verificar se o banco PONTO foi restaurado com sucesso antes de executar a query pós-restore
        if os.path.isfile(caminho_dados_ponto) and os.path.isfile(caminho_log_ponto):
            # Alterar o modelo de recuperação para "Simples" para PATRIO
            try:
                # Comando SQL para alterar o modelo de recuperação para "Simples" para PONTO
                alter_recovery_model_ponto = f'sqlcmd -S {server} -d {banco_de_dados_destino_ponto} -Q "ALTER DATABASE {banco_de_dados_destino_ponto} SET RECOVERY SIMPLE;"'
                subprocess.run(alter_recovery_model_ponto, check=True, shell=True)
                print_status(f"Modelo de recuperação alterado para 'Simples' para {letras}_AG_{cpf_cnpj}.")
            except subprocess.CalledProcessError as ex:
                print(f"Erro ao alterar o modelo de recuperação para 'Simples' para {letras}_AG_{cpf_cnpj}: {ex}")

            # Executar a query pós restore para PONTO
            try:
                # Estabelecer a conexão
                conn_ponto = pyodbc.connect(conn_str)

                # Criar um objeto cursor
                cursor_ponto = conn_ponto.cursor()

                # Executar USE master;
                cursor_ponto.execute("USE master;")

                # Comando SQL para realizar a query pós restore para PONTO
                query_pos_restore_ponto = f'''
                USE {letras_minusculas}_ponto_{cpf_cnpj}
                CREATE USER [{letras_minusculas}_{cpf_cnpj}.sql] FOR LOGIN [{letras_minusculas}_{cpf_cnpj}.sql]
                EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{cpf_cnpj}.sql';
                EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{cpf_cnpj}.sql';

                '''

                # Executar a query pós restore para PONTO
                cursor_ponto.execute(query_pos_restore_ponto)
                print_status(f"Query pós restore para\n{letras}_AG_{cpf_cnpj}\nexecutada com sucesso.")
                atualizar_status(f"Query pós restore para\n{letras}_AG_{cpf_cnpj}\nexecutada com sucesso.",root)
                progress_var.set(100)

                # Commit para confirmar as alterações
                conn_ponto.commit()

                # Fechar o cursor e a conexão
                cursor_ponto.close()
                conn_ponto.close()

                # Adicione feedback visual ao usuário
                atualizar_status(f"{letras}_PONTO_{cpf_cnpj}\n Concluido!",root)
                progress_var.set(100)

            except pyodbc.Error as ex:
                print_status(f"Erro ao executar a query pós restore para {letras}_PONTO_{cpf_cnpj}: {ex}")

                atualizar_status(f"Erro ao executar a query pós restore para {letras}_PONTO_{cpf_cnpj}: {ex}",root, fg="red")

    else:
        print_status(f"Banco {letras}_PONTO_{cpf_cnpj} não foi restaurado com sucesso.\nConsulte os logs para mais detalhes.")

    executaondemand_bat_ac_patrio_ponto(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root)

# Chamando o Atualizador

def executaondemand_bat(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo)
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco ao final
        verificar_atualizacao_banco(root)

def executaondemand_bat_ac_patrio(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo)
        os.system(caminho_arquivo_patrio)
        
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco_ac_ag ao final
        verificar_atualizacao_banco_ac_patrio(root)

def executaondemand_bat_ac_patrio_ag(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo)
        os.system(caminho_arquivo_patrio)
        os.system(caminho_arquivo_ag)
        
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco_ac_ag ao final
        verificar_atualizacao_banco_ac_patrio_ag(root)

def executaondemand_bat_ac_ag(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo)
        os.system(caminho_arquivo_ag)
        
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco_ac_ag ao final
        verificar_atualizacao_banco_ac_ag(root)

def executaondemand_bat_ag(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo_ag)
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco_ag ao final
        verificar_atualizacao_banco_ag(root)

def executaondemand_bat_ponto(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo_ponto)
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco_ponto ao final
        verificar_atualizacao_banco_ponto(root)

def executaondemand_bat_ac_ponto(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo)
        os.system(caminho_arquivo_ponto)
        
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco_ac_ponto ao final
        verificar_atualizacao_banco_ac_ponto(root)

def executaondemand_bat_ac_patrio_ponto(caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto, root):
    try:
        os.system(caminho_arquivo)
        os.system(caminho_arquivo_patrio)
        os.system(caminho_arquivo_ponto)
        
        # Mensagem de sucesso ao usuário
        print("Atualização no Banco executado com sucesso!")
        print_status("Atualização no Banco executado com sucesso!")
        atualizar_status("Atualização no Banco executado com sucesso!", root)
    except Exception as e:
        # Mensagem de erro ao usuário
        print(f"Erro ao executar o arquivo: {e}")
        print_status(f"Erro ao executar o arquivo: {e}")
        atualizar_status(f"Erro ao executar o arquivo: {e}", root)

        return
    finally:
        # Chama a função verificar_atualizacao_banco_ac_patrio_ponto ao final
        verificar_atualizacao_banco_ac_patrio_ponto(root)


# Verificando se tudo aconteceu conforme o esperado (a chamada do Atualizador e Restores)

def verificar_atualizacao_banco(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations()
    elif resposta == 'no':
         verificar_atualizacao_banco(root)

def verificar_atualizacao_banco_ac_patrio(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations_ac_patrio()
    elif resposta == 'no':
         verificar_atualizacao_banco_ac_patrio(root)

def verificar_atualizacao_banco_ac_patrio_ag(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations_ac_patrio_ag()
    elif resposta == 'no':
         verificar_atualizacao_banco_ac_patrio_ag(root)

def verificar_atualizacao_banco_ac_ag(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations_ac_ag()
    elif resposta == 'no':
         verificar_atualizacao_banco_ac_ag(root)

def verificar_atualizacao_banco_ag(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations_ag()
    elif resposta == 'no':
         verificar_atualizacao_banco_ag(root)

def verificar_atualizacao_banco_ponto(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations_ponto()
    elif resposta == 'no':
         verificar_atualizacao_banco_ponto(root)

def verificar_atualizacao_banco_ac_ponto(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations_ac_ponto()
    elif resposta == 'no':
         verificar_atualizacao_banco_ac_ponto(root)

def verificar_atualizacao_banco_ac_patrio_ponto(root):

    # Adicione mensagens de log
    print_status("Restore concluido com sucesso!")
    atualizar_status("Verificação no banco de dados realizada.", root)

    # Adicione feedback visual ao usuário
    print_status("Verificação de atualização no banco realizada.")
    atualizar_status("Restore concluido com sucesso!", root)

    resposta = messagebox.askquestion("Verificando Atualização no Banco", "O Restore e atualização\nAconteceram conforme o esperado?")
    if resposta == 'yes':
        descomentar_linhas_operations_ac_patrio_ponto()
    elif resposta == 'no':
         verificar_atualizacao_banco_ac_patrio_ponto(root)


# Caso o Banco de dados já exista no servidor, o usuario terá a chance de tentar
                # Novamente utlizando outros parametros

def tente_novamente(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):

    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_AC_{cpf_cnpj} ou {letras}_PATRIO_{cpf_cnpj} já existe. Deseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC com sucesso.', root)

        except Exception as ex:
            print(f'Erro ao remover a pasta de CloudUpCmd\\AC: {ex}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Uso da função para remover a última linha
            config_ini_path = 'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\config.ini'
            remover_ultima_linha(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())

def tente_novamente_ac_patrio(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):
    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_AC_{cpf_cnpj}, {letras}_AG_{cpf_cnpj}\nou {letras}_PATRIO_{cpf_cnpj} já existe. Deseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')
        atualizar_status(f'Tentando novamente...', root)

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\{nome_pasta}')
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC e CloudUpCmd\\PATRIO com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC e CloudUpCmd\\PATRIO com sucesso.', root)
        except Exception as e:
            print(f'Erro ao remover a pasta de CloudUpCmd\\AC e CloudUpCmd\\PATRIO: {e}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Lista de caminhos dos arquivos de configuração
            config_ini_paths = [
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\config.ini',
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO\\config.ini'
            ]

            # Iterar sobre cada caminho de arquivo e remover a última linha
            for config_ini_path in config_ini_paths:
                remover_ultima_linha_ac_patrio(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha_ac_patrio(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())

def tente_novamente_ac_patrio_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):
    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_AC_{cpf_cnpj}, {letras}_PATRIO_{cpf_cnpj}\nou {letras}_AG_{cpf_cnpj} já existe. Deseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')
        atualizar_status(f'Tentando novamente...', root)

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\{nome_pasta}')
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO\\{nome_pasta}')
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC, CloudUpCmd\\PATRIO\ne CloudUpCmd\\AG com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC, CloudUpCmd\\PATRIO\ne CloudUpCmd\\AG com sucesso.', root)
        except Exception as e:
            print(f'Erro ao remover a Pasta {nome_pasta} de CloudUpCmd\\AC, CloudUpCmd\\PATRIO\ne CloudUpCmd\\AG: {e}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Lista de caminhos dos arquivos de configuração
            config_ini_paths = [
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\config.ini',
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO\\config.ini',
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG\\config.ini'
            ]

            # Iterar sobre cada caminho de arquivo e remover a última linha
            for config_ini_path in config_ini_paths:
                remover_ultima_linha_ac_patrio_ag(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha_ac_patrio_ag(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())

def tente_novamente_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):


    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_AG_{cpf_cnpj} já existe.\nDeseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AG com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AG com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a pasta de CloudUpCmd\\AG: {e}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Uso da função para remover a última linha
            config_ini_path = 'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG\\config.ini'
            remover_ultima_linha_ag(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha_ag(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())

def tente_novamente_ac_ag(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):
    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_AC_{cpf_cnpj}, {letras}_AG_{cpf_cnpj}\nou {letras}_PATRIO_{cpf_cnpj} já existe. Deseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\{nome_pasta}')
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC e CloudUpCmd\\AG com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC e CloudUpCmd\\AG com sucesso.', root)
        except Exception as e:
            print(f'Erro ao remover a pasta de CloudUpCmd\\AC e CloudUpCmd\\AG: {e}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Lista de caminhos dos arquivos de configuração
            config_ini_paths = [
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\config.ini',
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG\\config.ini'
            ]

            # Iterar sobre cada caminho de arquivo e remover a última linha
            for config_ini_path in config_ini_paths:
                remover_ultima_linha_ac_ag(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha_ac_ag(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())

def tente_novamente_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):

    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_PONTO_{cpf_cnpj} já existe.\nDeseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\PONTO com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\PONTO com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a pasta de CloudUpCmd\\PONTO: {e}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Uso da função para remover a última linha
            config_ini_path = 'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO\\config.ini'
            remover_ultima_linha_ponto(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha_ponto(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())

def tente_novamente_ac_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):
    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_AC_{cpf_cnpj}, {letras}_AG_{cpf_cnpj}\nou {letras}_PONTO_{cpf_cnpj} já existe. Deseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\{nome_pasta}')
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC e CloudUpCmd\\PONTO com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC e CloudUpCmd\\PONTO com sucesso.', root)
        except Exception as e:
            print(f'Erro ao remover a pasta de CloudUpCmd\\AC e CloudUpCmd\\PONTO: {e}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Lista de caminhos dos arquivos de configuração
            config_ini_paths = [
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\config.ini',
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO\\config.ini'
            ]

            # Iterar sobre cada caminho de arquivo e remover a última linha
            for config_ini_path in config_ini_paths:
                remover_ultima_linha_ac_ponto(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha_ac_ponto(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())

def tente_novamente_ac_patrio_ponto(nome_pasta, letras, letras_minusculas, cpf_cnpj, senha, root, progress_bar, progress_var):
    progress_var.set(0)

    mensagem = f'O banco de dados {letras}_AC_{cpf_cnpj}, {letras}_PATRIO_{cpf_cnpj}\nou {letras}_PONTO_{cpf_cnpj} já existe. Deseja tentar novamente?'

    # Exibir a messagebox
    resposta = messagebox.askquestion("Banco de Dados Existente", mensagem)

    if resposta == 'yes':
        print_status(f'Tentando novamente...')
        atualizar_status(f'Tentando novamente...', root)

        # Limpar os campos
        letras_entry.delete(0, 'end')
        cpf_cnpj_entry.delete(0, 'end')
        senha_cliente.delete(0, 'end')

        # Remover a pasta específica
        try:
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\{nome_pasta}')
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO\\{nome_pasta}')
            shutil.rmtree(f'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO\\{nome_pasta}')
            print_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC, CloudUpCmd\\PATRIO\ne CloudUpCmd\\PONTO com sucesso.')
            atualizar_status(f'Pasta {nome_pasta} removida de CloudUpCmd\\AC, CloudUpCmd\\PATRIO\ne CloudUpCmd\\PONTO com sucesso.', root)
        except Exception as e:
            print(f'Erro ao remover a Pasta {nome_pasta} de CloudUpCmd\\AC, CloudUpCmd\\PATRIO\ne CloudUpCmd\\PONTO: {e}')

        # Remover a última linha no arquivo config.ini dentro da seção [Operations]
        try:
            # Lista de caminhos dos arquivos de configuração
            config_ini_paths = [
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\config.ini',
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO\\config.ini',
                'C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO\\config.ini'
            ]

            # Iterar sobre cada caminho de arquivo e remover a última linha
            for config_ini_path in config_ini_paths:
                remover_ultima_linha_ac_patrio_ponto(config_ini_path)

            print_status(f'Última linha removida em config.ini com sucesso.')
            atualizar_status(f'Última linha removida em config.ini com sucesso.', root)

        except Exception as e:
            print(f'Erro ao remover a última linha em config.ini: {e}')

        # Agora você pode chamar sua função criar_ini novamente
        habilitar_entradas()
        atualizar_status("Tente novamente!", root)

def remover_ultima_linha_ac_patrio_ponto(file_path):
    with open(file_path, 'r') as config_file:
        lines = config_file.readlines()

    # Remover a quebra de linha da penúltima linha se houver alguma
    if len(lines) > 1:
        lines[-2] = lines[-2].rstrip('\n')

    # Remover a última linha se houver alguma
    if lines:
        lines.pop()

    with open(file_path, 'w') as config_file:
        # Escrever as linhas, evitando escrever uma linha em branco no final
        config_file.writelines(line for line in lines if line.strip())


# Seção para descomentar linhas ao final de todos os processos

def descomentar_linhas_operations():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()

def descomentar_linhas_operations_ac_patrio():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Agora descomentar AG
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em AG
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)


    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()

def descomentar_linhas_operations_ac_patrio_ag():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Agora descomentar PATRIO
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em PATRIO
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Agora descomentar AG
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em AG
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()

def descomentar_linhas_operations_ac_ag():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Agora descomentar AG
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em AG
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)


    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()

def descomentar_linhas_operations_ag():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em AG
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)


    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()

def descomentar_linhas_operations_ponto():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()

def descomentar_linhas_operations_ac_ponto():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Agora descomentar AG
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em AG
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)


    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()

def descomentar_linhas_operations_ac_patrio_ponto():
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Agora descomentar PATRIO
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em PATRIO
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    #Agora descomentar AG
    config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", "config.ini")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    # Encontre a seção [Operations] no arquivo congig.ini em PONTO
    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    # Encontre o final da seção [Operations]
    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    # Descomente todas as linhas dentro da seção [Operations]
    for i in range(operations_section_start + 1, operations_section_end):
        lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


    # Modificar o arquivo para descomentar as linhas
    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # Exiba uma mensagem informativa
    messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
    print("Linhas descomentadas com sucesso.")
    print_status("Linhas descomentadas com sucesso.")

    # Encerrar Aplicação
    root.quit()


# Configurações para centralização da janela, e caminho relativo para a imagem no PyInstaller

def centralizar_janela(root, width, height):
    # Atualizar a janela para garantir que todos os elementos sejam renderizados
    root.update_idletasks()

    # Obter a largura e a altura da tela
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calcular as coordenadas x e y para a janela
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Definir a geometria da janela
    root.geometry('{}x{}+{}+{}'.format(width, height, x, y))

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



# Criação da janelas e Logica por traz do Menu Adicionar Sistema

def atualizar_status_adicionar_sistema(label, mensagem):
    """ Atualiza o Label com a mensagem fornecida """
    label.config(text=mensagem)
    label.update_idletasks()


def status_restore_info(label, mensagem):

    label.config(text=mensagem)
    label.update_idletasks()


def Adicionar_sistema_config():
    global clientes, sistemas_disponiveis, status_label_adicionar_sistema, listbox_clientes, combobox_caminhos, frame_opcoes
    clientes = {}

    def buscar_cliente():
        global clientes
        cliente_nome = entry_cliente.get().lower()
        caminho_selecionado = combobox_caminhos.get()

        if not cliente_nome:
            messagebox.showerror("Erro", "Por favor, insira o nome do cliente")
            return

        atualizar_status_adicionar_sistema(status_label_adicionar_sistema, "Buscando clientes...")
        clientes = buscar_dados_clientes(caminho_selecionado)
        clientes_filtrados = {nome: sistemas for nome, sistemas in clientes.items() if cliente_nome in nome.lower()}

        if clientes_filtrados:
            atualizar_opcoes_clientes(list(clientes_filtrados.keys()))
            atualizar_status_adicionar_sistema(status_label_adicionar_sistema, f"{len(clientes_filtrados)} cliente(s) encontrado(s).\nSelecione um cliente.")
        else:
            messagebox.showerror("Erro", "Cliente não encontrado")
            atualizar_status_adicionar_sistema(status_label_adicionar_sistema, "Cliente não encontrado.")
    
    def atualizar_opcoes_clientes(opcoes):
        listbox_clientes.delete(0, tk.END)
        for cliente in opcoes:
            listbox_clientes.insert(tk.END, cliente)

    def selecionar_cliente(event):
        global cliente_selecionado
        cliente_selecionado = listbox_clientes.get(listbox_clientes.curselection())
        sistemas_cliente = clientes[cliente_selecionado]
        sistemas_para_adicionar = [sistema for sistema in sistemas_disponiveis if sistema not in sistemas_cliente]
        atualizar_opcoes_sistemas(sistemas_para_adicionar, cliente_selecionado, '', '')
        atualizar_status_adicionar_sistema(status_label_adicionar_sistema, f"Este cliente já possui os sistemas\n{', '.join(sistemas_cliente)}.\n\nEstas são as opções de adição para\neste cliente:\n↓")

    def atualizar_opcoes_sistemas(opcoes, cliente_nome, caminho_log, caminho_dados):
        for widget in frame_opcoes.winfo_children():
            widget.destroy()
        for idx, sistema in enumerate(opcoes):
            tk.Button(frame_opcoes, text=sistema, command=lambda s=sistema: adicionar_sistema(cliente_nome, s, caminho_log, caminho_dados)).grid(row=0, column=idx, padx=5, pady=5)

    def adicionar_sistema(cliente_nome, sistema, caminho_log, caminho_dados):
        global cliente_selecionado
        if cliente_selecionado in clientes:
            try:
                caminho_cliente_logs = os.path.join("D:\\BDS\\Logs", cliente_selecionado, sistema)
                caminho_cliente_dados = os.path.join("D:\\BDS\\Dados", cliente_selecionado, sistema)

                if not os.path.exists(caminho_cliente_logs):
                    os.makedirs(caminho_cliente_logs)
                if not os.path.exists(caminho_cliente_dados):
                    os.makedirs(caminho_cliente_dados)

                messagebox.showinfo("Confirmação", f"As pastas do sistema {sistema} foram criadas para o cliente {cliente_selecionado}.")
                atualizar_status_adicionar_sistema(status_label_adicionar_sistema, f"As pastas do sistema {sistema}\nforam criadas para o cliente {cliente_selecionado}.\nPor favor, insira os dígitos do CPF/CNPJ,\nas letras e a senha.")

                # Solicitar informações adicionais
                solicitar_informacoes(cliente_nome, sistema, caminho_log, caminho_dados, progress_var)

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao adicionar sistema: {e}")
                atualizar_status_adicionar_sistema(status_label_adicionar_sistema, f"Erro ao adicionar sistema: {e}")
        else:
            messagebox.showerror("Erro", "Cliente não encontrado")
            atualizar_status_adicionar_sistema(status_label_adicionar_sistema, "Cliente não encontrado.")

    def buscar_dados_clientes(caminho):
        clientes = {}
        try:
            for root, dirs, files in os.walk(caminho):
                for dir_name in dirs:
                    sistemas_cliente = os.listdir(os.path.join(root, dir_name))
                    # Verifica quais sistemas disponíveis o cliente possui
                    sistemas_encontrados = [sistema for sistema in sistemas_disponiveis if sistema in sistemas_cliente]
                    if sistemas_encontrados:
                        clientes[dir_name] = sistemas_encontrados
        except Exception as e:
            print(f"Erro ao buscar dados dos clientes: {e}")
        return clientes
    
    
    def confirmar_restore(cliente_nome, sistema, caminho_log, caminho_dados, progress_var, status_label_restore):

        try:
        
            status_restore_info(status_label_restore, f"Iniciando restore para o Banco {sistema}.")

            letras_minusculas = entry_letras.get().strip().lower()
            digitos = entry_digitos.get().strip()
            letras = entry_letras.get().strip().upper()  # Converte as letras para maiúsculas
            senha = entry_senha.get().strip()

            if not digitos or not letras or not senha:
                messagebox.showerror("Erro", "Todos os campos são obrigatórios")
                return
            
            status_restore_info(status_label_restore, f"Iniciando restore para o Banco {sistema}..")
            progress_var.set(10)
            # Parâmetros de conexão com o SQL Server
            servers = ['FOR-CNT-BD-SQL5', 
                       'FORT-CNT-BDSQ-6',
                       'GLAUDSON-DESKTO\\SQLEXPRESS']
            trusted_connection = 'yes'  # Configurando para autenticação do Windows
            driver = '{SQL Server}'

            # Tentar conectar-se a cada servidor da lista
            conectado = False
            conn = None
            for server in servers:
                try:
                    # Criar a string de conexão
                    conn_str = f'DRIVER={driver};SERVER={server};Trusted_Connection={trusted_connection}'

                    # Tentar estabelecer a conexão
                    conn = pyodbc.connect(conn_str)

                    # Se a conexão for bem sucedida, sair do loop
                    print_status(f"Conectado ao servidor: {server}")
                    progress_var.set(30)  # Atualizar o valor da barra de progresso
                    # status_restore_info(status_label,f"Conectado ao servidor: {server}")
                    conectado = True
                    break
                except Exception as e:
                    print_status(f"Erro ao conectar-se ao servidor {server}: {str(e)}")
                    # status_restore_info(status_label,f"Erro ao conectar-se\nao servidor {server}:\n{str(e)}", is_error=True)

            if not conectado:
                print_status("Não foi possível conectar-se a nenhum dos servidores da lista.")
                status_restore_info(f"Não foi possível conectar-se\na nenhum dos servidores da lista.")
                return

            status_restore_info(status_label_restore, f"Iniciando restore para o Banco {sistema}...")

            # Parâmetros específicos para o restore
            banco_de_dados_destino = f'{letras}_{sistema}_{digitos}'
            caminho_dados = f'D:\\BDS\\dados\\{cliente_nome}\\{sistema}\\{letras}_{sistema}_{digitos}.mdf'
            caminho_log = f'D:\\BDS\\Logs\\{cliente_nome}\\{sistema}\\{letras}_{sistema}_{digitos}.ldf'

            progress_var.set(40)
            status_restore_info(status_label_restore, f"Verificando se {letras}_{sistema}_{digitos}\njá existe...")


            # Verificar se os bancos de dados já existem
            databases_to_check = [banco_de_dados_destino]

            try:
                conn_check = pyodbc.connect(conn_str)
                cursor_check = conn_check.cursor()

                # Verificar se os bancos de dados já existem
                query_check = "SELECT name FROM sys.databases WHERE name IN (?)"
                cursor_check.execute(query_check, databases_to_check)
                existing_databases = [row[0] for row in cursor_check.fetchall()]

                if existing_databases:
                    print('Os bancos de dados já existem.\n Não é possível continuar o restore.')
                    print_status("Os bancos de dados já existem.\n Não é possível continuar o restore.")
                    status_restore_info(status_label_restore, f"Os bancos de dados já existem.\nNão é possível continuar o restore.")
                    return

            except Exception as ex:
                print(f"Erro ao verificar a existência\ndos bancos de dados:\n{ex}")
                status_restore_info(status_label_restore, f"Erro ao verificar a existência\ndos bancos de dados:\n{ex}")
                return

            finally:
                if conn_check:
                    conn_check.close()

            status_restore_info(status_label_restore, f"Restaurando\n{letras}_{sistema}_{digitos}...")
            progress_var.set(60)

            # Comando SQL para realizar o restore usando caminhos fixos para cada Sistema
            # Determinar o tipo de sistema para construir a query de restore apropriada
            if sistema == 'AC':
                restore_query = f'RESTORE DATABASE {banco_de_dados_destino} FROM DISK = \'D:\\Bancos Limpos SQL\\{sistema}.bak\' WITH REPLACE, MOVE \'{sistema}_Data\' TO \'{caminho_dados}\', MOVE \'{sistema}_Log\' TO \'{caminho_log}\''
            elif sistema == 'AG':
                restore_query = f'RESTORE DATABASE {banco_de_dados_destino} FROM DISK = \'D:\\Bancos Limpos SQL\\{sistema}.bak\' WITH REPLACE, MOVE \'{sistema}_Data\' TO \'{caminho_dados}\', MOVE \'{sistema}_Log\' TO \'{caminho_log}\''
            elif sistema == 'PATRIO':
                restore_query = f'RESTORE DATABASE {banco_de_dados_destino} FROM DISK = \'D:\\Bancos Limpos SQL\\{sistema}.bak\' WITH REPLACE, MOVE \'{sistema}\' TO \'{caminho_dados}\', MOVE \'{sistema}_Log\' TO \'{caminho_log}\''
            elif sistema == 'PONTO':
                restore_query = f'RESTORE DATABASE {banco_de_dados_destino} FROM DISK = \'D:\\Bancos Limpos SQL\\{sistema}.bak\' WITH REPLACE, MOVE \'{sistema}3\' TO \'{caminho_dados}\', MOVE \'{sistema}3_Log\' TO \'{caminho_log}\''
            else:
                # Lógica para lidar com sistemas desconhecidos ou erro
                status_restore_info(status_label_restore, f"Sistema '{sistema}' não suportado para restore.")
                print(f"Sistema '{sistema}' não suportado para restore.")
                return

            progress_var.set(70)
            # Executar o comando de restore para o Sistema escolhido
            try:
                subprocess.run(['sqlcmd', '-S', server, '-d', 'master', '-Q', restore_query], shell=True, check=True)
                print_status(f"Restore do banco {letras}_{sistema}_{digitos} executado com sucesso.")
            except subprocess.CalledProcessError as e:
                print(f"Erro no restore do banco {letras}_{sistema}_{digitos}: {e}")

            # Verificar se o banco foi restaurado com sucesso antes de executar a query pós-restore
            if os.path.isfile(caminho_dados) and os.path.isfile(caminho_log):
                # Alterar o modelo de recuperação para "Simples" para o Sistema
                try:
                    # Comando SQL para alterar o modelo de recuperação para "Simples" para o Sistema
                    alter_recovery_model = f'sqlcmd -S {server} -d {banco_de_dados_destino} -Q "ALTER DATABASE {banco_de_dados_destino} SET RECOVERY SIMPLE;"'
                    subprocess.run(alter_recovery_model, check=True, shell=True)
                    print_status(f"Modelo de recuperação alterado para 'Simples'\n Em {letras}_{sistema}_{digitos}.")
                    status_restore_info(status_label_restore, f"Modelo de recuperação\nalterado para 'Simples'\nEm {letras}_{sistema}_{digitos}.")
                except subprocess.CalledProcessError as e:
                    print(f"Erro ao alterar o modelo de recuperação para 'Simples'\n para {letras}_{sistema}_{digitos}: {e}")
                    status_restore_info(status_label_restore, f"Erro ao alterar o modelo de recuperação para 'Simples'\n para {letras}_{sistema}_{digitos}: {e}")
                    return

                try:
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    # Comando SQL para realizar a query pós restore para o Sistema
                    query_pos_restore = f'''
                    USE {letras}_{sistema}_{digitos}
                    CREATE USER [{letras_minusculas}_{digitos}.sql] FOR LOGIN [{letras_minusculas}_{digitos}.sql]
                    EXEC sp_addrolemember 'DB_DATAREADER', '{letras_minusculas}_{digitos}.sql';
                    EXEC sp_addrolemember 'DB_DATAWRITER', '{letras_minusculas}_{digitos}.sql';
                    EXEC sp_addrolemember 'DB_DDLADMIN', '{letras_minusculas}_{digitos}.sql';
                    EXEC sp_addrolemember 'DB_OWNER', '{letras_minusculas}_{digitos}.sql';

                    '''
                    # Executar a query para AC
                    cursor.execute(query_pos_restore)
                    conn.commit()
                    print_status(f"Query pós-restore para {sistema} executada com sucesso.")
                    status_restore_info(status_label_restore, f"Query pós-restore para\n{sistema}\nexecutada com sucesso.")

                    # Liberar recursos
                    cursor.close()
                    conn.close()

                    print_status(f"Procedimento concluído para {cliente_nome} com sucesso.")
                    # Após a conclusão do processo
                    status_restore_info(status_label_restore, "Restore Concluído\ncom sucesso!")
                    

                except pyodbc.Error as e:
                    print_status(f"Erro ao executar\nquery pós-restore para\n{sistema}:\n{e}")
                    status_restore_info(status_label_restore, f"Erro ao executar\nquery pós-restore para\n{sistema}:\n{e}")
                    return
            else:
                print_status(f"Erro no restore do banco {letras}_{sistema}_{digitos}: Arquivos de dados ou log não encontrados.")
                status_restore_info(status_label_restore, f"Erro no restore do banco {letras}_{sistema}_{digitos}:\nArquivos de dados ou log não encontrados.")

            status_restore_info(status_label_restore, f"Criando '.ini'...")

            # Crie o arquivo .ini no local desejado
            cliente_nome = cliente_selecionado
            # Usando o mesmo nome para a pasta do .ini
            caminho_ag_ini = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd", sistema, cliente_nome)
            os.makedirs(caminho_ag_ini, exist_ok=True)  # Crie a pasta se ela ainda não existir

            ini_path = os.path.join(caminho_ag_ini, f'{cliente_nome}.ini')
            with open(ini_path, 'w') as ini_file:
                ini_file.write(f'[Startup]\nProgramFolder = C:\\Atualiza\\Exe\\{sistema}\nDataFolder = D:\\BDS\\Dados\\{cliente_nome}\\{sistema}\nDatabaseFile = 127.0.0.1:{letras}_{sistema}_{digitos}\nDriverName = MSSQL\nUserName = {letras_minusculas}_{digitos}.sql\nPassword = {senha}\n\n[Settings]\nContinueUpdateAfterCrashRecovery = True\nAllowOldBackupsRestoration = False\nSkipWarnings = True\nByPassInUseCheck = True\nRetryOnDatabaseInUse = False\n\n[Backup]\nSkipBackupDatabase = True')


            status_restore_info(status_label_restore, f"Inserindo linha\nno 'Config.ini'...")
            # Modificar_config_ini para o novo Sistema
            config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd", sistema, "config.ini")

            with open(config_ini_path, 'r') as config_file:
                lines = config_file.readlines()

            # Encontre a seção [Operations] no arquivo
            operations_section_start = -1
            for i, line in enumerate(lines):
                if line.strip() == '[Operations]':
                    operations_section_start = i
                    break

            # Encontre o final da seção [Operations]
            operations_section_end = len(lines)
            for i in range(operations_section_start + 1, len(lines)):
                if lines[i].strip().startswith('['):
                    operations_section_end = i
                    break

            # Verifique se há linhas não comentadas dentro da seção [Operations]
            linhas_nao_comentadas = False
            for i in range(operations_section_start + 1, operations_section_end):
                if not lines[i].strip().startswith(';'):
                    linhas_nao_comentadas = True
                    break

            # Comente todas as linhas não comentadas dentro da seção [Operations]
            if linhas_nao_comentadas:
                for i in range(operations_section_start + 1, operations_section_end):
                    if not lines[i].strip().startswith(';'):
                        lines[i] = ';' + lines[i]

            # Construa a nova linha com os valores especificados
            nova_linha = f"Customer={cliente_nome},ExeName={sistema},ExeDirName={sistema},AppSubdir={cliente_nome},DbInstance=,DbName={letras}_{sistema}_{digitos},DbUser={letras_minusculas}_{digitos}.sql,DbPass={senha}"

            # Encontre a última linha comentada dentro da seção [Operations]
            ultima_linha_comentada_index = -1
            for i in range(operations_section_end - 1, operations_section_start, -1):
                if lines[i].strip().startswith(';'):
                    ultima_linha_comentada_index = i
                    break

            # Insira a nova linha após a última linha comentada ou no final da seção
            if ultima_linha_comentada_index != -1:
                lines.insert(ultima_linha_comentada_index + 1, '\n')  # Adicione uma nova linha
                lines.insert(ultima_linha_comentada_index + 2, nova_linha)  # Inserir nova linha
            else:
                # Se não houver linhas comentadas, insira no final da seção
                lines.insert(operations_section_end, nova_linha + '\n')

            # Modificar o arquivo para inserir a nova linha
            with open(config_ini_path, 'w') as config_file:
                config_file.writelines(lines)

            status_restore_info(status_label_restore, f"Concluído!")
            progress_var.set(100)

            messagebox.showinfo("Sucesso", f"Adição de Restore Concluído para {cliente_nome}.\n\n-> Criação '.Ini' Realizada.\n-> Linha adicionada em 'Config.ini'\n\nChamaremos o atualizador agora, ok?")

            if sistema == 'AC':
                try:
                    os.system(caminho_arquivo)
                    # Mensagem de sucesso ao usuário
                    print("Atualização no Banco executado com sucesso!")
                    print_status("Atualização no Banco executado com sucesso!")
                    status_restore_info(status_label_restore, f"Atualização no Banco\nexecutado com sucesso!")
                except Exception as e:
                    # Mensagem de erro ao usuário
                    print(f"Erro ao executar o arquivo: {e}")
                    print_status(f"Erro ao executar o arquivo: {e}")
                    status_restore_info(status_label_restore, f"Erro ao executar o arquivo:\n{e}")

                    return
                finally:
                    # Chama a função para descomentar as linhas ao final
                    
                        config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC", "config.ini")

                        with open(config_ini_path, 'r') as config_file:
                            lines = config_file.readlines()

                        # Encontre a seção [Operations] no arquivo
                        operations_section_start = -1
                        for i, line in enumerate(lines):
                            if line.strip() == '[Operations]':
                                operations_section_start = i
                                break

                        # Encontre o final da seção [Operations]
                        operations_section_end = len(lines)
                        for i in range(operations_section_start + 1, len(lines)):
                            if lines[i].strip().startswith('['):
                                operations_section_end = i
                                break

                        # Descomente todas as linhas dentro da seção [Operations]
                        for i in range(operations_section_start + 1, operations_section_end):
                            lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


                        # Modificar o arquivo para descomentar as linhas
                        with open(config_ini_path, 'w') as config_file:
                            config_file.writelines(lines)

                        # Exiba uma mensagem informativa
                        messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
                        print("Linhas descomentadas com sucesso.")
                        print_status("Linhas descomentadas com sucesso.")
                        status_restore_info(status_label_restore, f"Linhas descomentadas com sucesso.")

                        # Encerrar janela
                        menu_bar.quit()

            elif sistema == 'PATRIO':
                try:
                    os.system(caminho_arquivo_patrio)
                    # Mensagem de sucesso ao usuário
                    print("Atualização no Banco executado com sucesso!")
                    print_status("Atualização no Banco executado com sucesso!")
                    status_restore_info(status_label_restore, f"Atualização no Banco\nexecutado com sucesso!")
                except Exception as e:
                    # Mensagem de erro ao usuário
                    print(f"Erro ao executar o arquivo: {e}")
                    print_status(f"Erro ao executar o arquivo: {e}")
                    status_restore_info(status_label_restore, f"Erro ao executar o arquivo:\n{e}")

                    return
                finally:
                    # Chama a função verificar_atualizacao_banco ao final
                    
                        config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO", "config.ini")

                        with open(config_ini_path, 'r') as config_file:
                            lines = config_file.readlines()

                        # Encontre a seção [Operations] no arquivo
                        operations_section_start = -1
                        for i, line in enumerate(lines):
                            if line.strip() == '[Operations]':
                                operations_section_start = i
                                break

                        # Encontre o final da seção [Operations]
                        operations_section_end = len(lines)
                        for i in range(operations_section_start + 1, len(lines)):
                            if lines[i].strip().startswith('['):
                                operations_section_end = i
                                break

                        # Descomente todas as linhas dentro da seção [Operations]
                        for i in range(operations_section_start + 1, operations_section_end):
                            lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


                        # Modificar o arquivo para descomentar as linhas
                        with open(config_ini_path, 'w') as config_file:
                            config_file.writelines(lines)

                        # Exiba uma mensagem informativa
                        messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
                        print("Linhas descomentadas com sucesso.")
                        print_status("Linhas descomentadas com sucesso.")
                        status_restore_info(status_label_restore, f"Linhas descomentadas com sucesso.")

                        # Encerrar janela
                        menu_bar.quit()

            elif sistema == 'AG':
                try:
                    os.system(caminho_arquivo_ag)
                    # Mensagem de sucesso ao usuário
                    print("Atualização no Banco executado com sucesso!")
                    print_status("Atualização no Banco executado com sucesso!")
                    status_restore_info(status_label_restore, f"Atualização no Banco\nexecutado com sucesso!")
                except Exception as e:
                    # Mensagem de erro ao usuário
                    print(f"Erro ao executar o arquivo: {e}")
                    print_status(f"Erro ao executar o arquivo: {e}")
                    status_restore_info(status_label_restore, f"Erro ao executar o arquivo:\n{e}")

                    return
                finally:
                    # Chama a função para descomentar as linhas ao final
                    
                        config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG", "config.ini")

                        with open(config_ini_path, 'r') as config_file:
                            lines = config_file.readlines()

                        # Encontre a seção [Operations] no arquivo
                        operations_section_start = -1
                        for i, line in enumerate(lines):
                            if line.strip() == '[Operations]':
                                operations_section_start = i
                                break

                        # Encontre o final da seção [Operations]
                        operations_section_end = len(lines)
                        for i in range(operations_section_start + 1, len(lines)):
                            if lines[i].strip().startswith('['):
                                operations_section_end = i
                                break

                        # Descomente todas as linhas dentro da seção [Operations]
                        for i in range(operations_section_start + 1, operations_section_end):
                            lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


                        # Modificar o arquivo para descomentar as linhas
                        with open(config_ini_path, 'w') as config_file:
                            config_file.writelines(lines)

                        # Exiba uma mensagem informativa
                        messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
                        print("Linhas descomentadas com sucesso.")
                        print_status("Linhas descomentadas com sucesso.")
                        status_restore_info(status_label_restore, f"Linhas descomentadas com sucesso.")

                        # Encerrar janela
                        menu_bar.quit()

            elif sistema == 'PONTO':
                try:
                    os.system(caminho_arquivo_ponto)
                    # Mensagem de sucesso ao usuário
                    print("Atualização no Banco executado com sucesso!")
                    print_status("Atualização no Banco executado com sucesso!")
                    status_restore_info(status_label_restore, f"Atualização no Banco\nexecutado com sucesso!")
                except Exception as e:
                    # Mensagem de erro ao usuário
                    print(f"Erro ao executar o arquivo: {e}")
                    print_status(f"Erro ao executar o arquivo: {e}")
                    status_restore_info(status_label_restore, f"Erro ao executar o arquivo:\n{e}")

                    return
                finally:
                    # Chama a função para descomentar as linhas ao final
                    
                        config_ini_path = os.path.join("C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO", "config.ini")

                        with open(config_ini_path, 'r') as config_file:
                            lines = config_file.readlines()

                        # Encontre a seção [Operations] no arquivo
                        operations_section_start = -1
                        for i, line in enumerate(lines):
                            if line.strip() == '[Operations]':
                                operations_section_start = i
                                break

                        # Encontre o final da seção [Operations]
                        operations_section_end = len(lines)
                        for i in range(operations_section_start + 1, len(lines)):
                            if lines[i].strip().startswith('['):
                                operations_section_end = i
                                break

                        # Descomente todas as linhas dentro da seção [Operations]
                        for i in range(operations_section_start + 1, operations_section_end):
                            lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')


                        # Modificar o arquivo para descomentar as linhas
                        with open(config_ini_path, 'w') as config_file:
                            config_file.writelines(lines)

                        # Exiba uma mensagem informativa
                        messagebox.showinfo("Descomentar Linhas", "Linhas descomentadas com sucesso.")
                        print("Linhas descomentadas com sucesso.")
                        print_status("Linhas descomentadas com sucesso.")
                        status_restore_info(status_label_restore,f"Linhas descomentadas com sucesso.")

                        # Encerrar janela
                        menu_bar.quit()

        except Exception as e:
            status_restore_info(status_label_restore, f"Erro durante a restauração:\n{e}")
        

    def solicitar_informacoes(cliente_nome, sistema, caminho_log, caminho_dados, progress_var):
        # Criar janela de entrada para solicitar os dados
        info_window = tk.Toplevel(root)
        info_window.title("Informações para o Banco Sistema")
        largura_janela = 300
        altura_janela = 353

        largura_tela = menu_bar.winfo_screenwidth()
        altura_tela = menu_bar.winfo_screenheight()
        
        x = (largura_tela - largura_janela) // 2
        y = (altura_tela - altura_janela) // 2

        info_window.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

        # Widgets para entrada de dados
        tk.Label(info_window, text="Insira as letras:").place(x=109, y=15, anchor='nw')
        global entry_letras
        entry_letras = tk.Entry(info_window)
        entry_letras.place(x=75, y=45, width=150, anchor='nw')

        tk.Label(info_window, text="Insira os dígitos do CPF/CNPJ:").place(x=70, y=75, anchor='nw')
        global entry_digitos
        entry_digitos = tk.Entry(info_window)
        entry_digitos.place(x=75, y=105, width=150, anchor='nw')

        tk.Label(info_window, text="Senha:").place(x=127, y=135, anchor='nw')
        global entry_senha
        entry_senha = tk.Entry(info_window)
        entry_senha.place(x=75, y=165, width=150, anchor='nw')

        # Barra de progresso
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(info_window, variable=progress_var, length=150, maximum=100)
        progress_bar.place(x=75, y=300, anchor='nw')

        # Frame para feedback visual
        status_label_restore = tk.Label(info_window, text="", font=("Roboto", 8), fg="blue")
        status_label_restore.place(x=60, y=250)


        # Passar todos os parâmetros necessários para a função de confirmação
        confirmar_restore_parcial = partial(confirmar_restore, cliente_nome, sistema, caminho_log, caminho_dados, progress_var, status_label_restore)
        tk.Button(info_window, text="Iniciar Restore", command=confirmar_restore_parcial).place(x=109, y=195, anchor='nw')




    

    # Interface da janela de adicionar sistema
    menu_bar = tk.Toplevel(root)
    menu_bar.title("Adicionar Sistema")
    largura_janela = 500
    altura_janela = 375

    caminho_icon = resource_path("loki_s_helmet.ico")

    if os.path.exists(caminho_icon):
        icon = ImageTk.PhotoImage(Image.open(caminho_icon))
        menu_bar.iconphoto(True, icon)

    largura_tela = menu_bar.winfo_screenwidth()
    altura_tela = menu_bar.winfo_screenheight()
    
    x = (largura_tela - largura_janela) // 2
    y = (altura_tela - altura_janela) // 2
    
    menu_bar.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

    sistemas_disponiveis = ["AC", "PATRIO", "AG", "PONTO"]
    caminhos_predefinidos = [
        
        r"E:\BDS\Dados",
        r"D:\BDS\Dados",
        r"C:\BDS\Dados"
    ]

    frame_principal = tk.Frame(menu_bar)
    frame_principal.grid(row=0, column=0, padx=20, pady=20)

    tk.Label(frame_principal, text="Buscar Cliente").grid(row=3, column=0, sticky="w")
    entry_cliente = tk.Entry(frame_principal)
    entry_cliente.grid(row=4, column=0, pady=5)

    tk.Button(frame_principal, text="Buscar", command=buscar_cliente).grid(row=5, column=0, pady=5)

    tk.Label(frame_principal, text="Onde buscar?").grid(row=0, column=0, sticky="w")
    combobox_caminhos = ttk.Combobox(frame_principal, values=caminhos_predefinidos, state='readonly', width=35)
    combobox_caminhos.grid(row=1, column=0, pady=5)
    combobox_caminhos.current(0)

    frame_resultados = tk.Frame(menu_bar)
    frame_resultados.grid(row=0, column=2, padx=20, pady=20)

    tk.Label(frame_resultados, text="Clientes Encontrados:").grid(row=0, column=0, sticky="w")
    listbox_clientes = tk.Listbox(frame_resultados)
    listbox_clientes.grid(row=1, column=0, pady=5)
    listbox_clientes.bind('<<ListboxSelect>>', selecionar_cliente)

    frame_opcoes = tk.Frame(menu_bar)
    frame_opcoes.grid(row=1, column=0, padx=20, pady=20)

    status_label_adicionar_sistema = tk.Label(frame_principal, text="", fg="blue")
    status_label_adicionar_sistema.grid(row=6, column=0, padx=20, pady=10)


# Logica por traz do Menu Atualização

def get_server_version():
    try:
        response = urllib.request.urlopen(VERSION_URL)
        version = response.read().decode('utf-8').strip()
        print(f"Versão do servidor: {version}")  # Print de depuração
        return version
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao obter a versão do servidor: {e}")

def update_local_version(new_version):
    try:
        with open("version.txt", "w") as f:
            f.write(new_version)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar a versão local: {e}")

def download_and_extract_update(progress_bar, progress_var_update):
    try:
        # Simula obtenção da versão do servidor
        server_version = get_server_version()

        print(f"Versão do servidor: {server_version}")
        progress_var_update.set(10)
        progress_bar.update()

        local_filename = f"arquivo_{server_version}_update.rar"
        temp_filename = "Loki.rar"

        # Baixa o arquivo usando gdown
        print("Iniciando download...")
        gdown.download(UPDATE_URL, temp_filename, quiet=False)
        print(f"Download concluído: {temp_filename}")

        progress_var_update.set(35)
        progress_bar.update()

        # Cria o diretório de destino se ele não existir
        destination_dir = f"./versao_{server_version}"
        os.makedirs(destination_dir, exist_ok=True)
        print(f"Diretório de destino criado: {destination_dir}")

        time.sleep(1)

        # Move e renomeia o arquivo baixado para o diretório de destino
        destination_file = os.path.join(destination_dir, local_filename)
        shutil.move(temp_filename, destination_file)
        print(f"Arquivo movido para: {destination_file}")

        progress_var_update.set(50)
        progress_bar.update()

        time.sleep(1)

        progress_var_update.set(60)
        progress_bar.update()

        time.sleep(1)

        # Função para extrair o arquivo com WinRAR
        def extract_file_with_winrar(rar_file, extract_to):
            winrar_path = "C:\\Program Files\\WinRAR\\WinRAR.exe"  # Verifique se este é o caminho correto no seu sistema
            comando = [winrar_path, 'x', rar_file, extract_to]
            try:
                result = subprocess.run(comando, check=True, capture_output=True, text=True)
                print(f"Arquivo extraído com sucesso para {extract_to}")
                print(result.stdout)
            except subprocess.CalledProcessError as e:
                print(f"Erro durante a extração do arquivo: {e}")
                print(e.stdout)
                print(e.stderr)

        # Cria o diretório de destino se ele não existir
        extract_to = os.path.join(os.getcwd(), f"versao_{server_version}")
        os.makedirs(extract_to, exist_ok=True)
        print(f"Extraindo arquivos para: {extract_to}")

        extract_file_with_winrar(destination_file, extract_to)

        progress_var_update.set(100)
        progress_bar.update()

        messagebox.showinfo("Atualização", f"Download da nova versão concluído com sucesso!\n\nNavegue até a pasta atual\ne acesse a pasta 'versão_{server_version}'\n\nEncontrará a nova versão disponibilizada lá!")

        root.destroy()

    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante a verificação/atualização: {e}")

def check_for_updates(manual=False):
    try:
        server_version = get_server_version()

        # Comparando versões
        if CURRENT_VERSION != server_version:
            if manual:
                response = messagebox.askyesno("Atualização Disponível", f"Nova versão disponível: {server_version}.\nDeseja atualizar agora?")
                if response:
                    show_progress_bar()
            else:
                messagebox.showinfo("Atualização Disponível", f"Nova versão disponível: {server_version}.\n\n->Vá até o menu\n->Busque 'Verificar atualizações'\n\n Para atualizar a ferramenta a versão mais recente.")
        elif manual:
            messagebox.showinfo("Atualização", "Você está utilizando a versão mais recente.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante a verificação/atualização: {e}")

def show_progress_bar():
    
    menu_bar = tk.Toplevel(root)
    menu_bar.title("Atualização")

    largura_janela = 230
    altura_janela = 145

    caminho_icon = resource_path("loki_s_helmet.ico")

    if os.path.exists(caminho_icon):
        icon = ImageTk.PhotoImage(Image.open(caminho_icon))
        menu_bar.iconphoto(True, icon)

    menu_bar.geometry("230x145")

    largura_tela = menu_bar.winfo_screenwidth()
    altura_tela = menu_bar.winfo_screenheight()
    
    x = (largura_tela - largura_janela) // 2
    y = (altura_tela - altura_janela) // 2
    
    menu_bar.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

    progress_label = tk.Label(menu_bar, text="Baixando e descompactando\na nova versão...")
    progress_label.place(x=34, y=15,  anchor='nw')

    progress_var_update = tk.DoubleVar()
    progress_bar = ttk.Progressbar(menu_bar, variable=progress_var_update, length=120, maximum=100)
    progress_bar.place(x=54, y=80, anchor='nw')
    
    

    root.after(100, lambda: download_and_extract_update(progress_bar, progress_var_update))
    root.mainloop()

# Configuração do logger
logging.basicConfig(level=logging.INFO)


# Menu Atualização

def menu_sobre():

    messagebox.showinfo("Sobre a Ferramenta", "Foi idealizada para automatizar a realização de\nRestores do processo de Setup\n\nNome: Loki \nVersão: 8.2.3")


# Criação e configurações da janela principal

root = tk.Tk()
root.title("Loki")

# Obtém o caminho absoluto dos arquivos usando a função resource_path
caminho_icon = resource_path("loki_s_helmet.ico")
caminho_imagem = resource_path("alterar-a-senha.png")

# Configurar o ícone da janela
if os.path.exists(caminho_icon):
    icon = ImageTk.PhotoImage(Image.open(caminho_icon))
    root.tk.call("wm", "iconphoto", root._w, icon)
else:
    print(f"O arquivo {caminho_icon} não foi encontrado.")

largura_janela = 210
altura_janela = 520

largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

x = (largura_tela - largura_janela) // 2
y = (altura_tela - altura_janela) // 2

root.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

# Chame a função para centralizar a janela
centralizar_janela(root, largura_janela, altura_janela)

# Definir tamanho mínimo da janela
root.minsize(210, 520)

# Carregue a imagem original
imagem_original = Image.open(caminho_imagem)

#Checagem de update ao iniciar
check_for_updates()

# Criando o menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

config_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Menu", menu=config_menu)
config_menu.add_command(label="Adicionar Sistema", command=Adicionar_sistema_config)
config_menu.add_command(label="Verificar Atualizações", command=lambda: check_for_updates(manual=True))
config_menu.add_command(label="Sobre", command=menu_sobre)

# Redimensione a imagem para o tamanho desejado
largura_botao = 35  # Substitua pelo tamanho desejado
altura_botao = 20   # Substitua pelo tamanho desejado
imagem_redimensionada = imagem_original.resize((largura_botao, altura_botao), Image.Resampling.LANCZOS)

# Caminho do ExecutaOnDemand.bat
caminho_arquivo =('C:\\Atualiza\\CloudUp\\CloudUpCmd\\AC\\ExecutaOnDemand.bat')
caminho_arquivo_ag =('C:\\Atualiza\\CloudUp\\CloudUpCmd\\AG\\ExecutaOnDemand.bat')
caminho_arquivo_patrio =('C:\\Atualiza\\CloudUp\\CloudUpCmd\\PATRIO\\ExecutaOnDemand.bat')
caminho_arquivo_ponto =('C:\\Atualiza\\CloudUp\\CloudUpCmd\\PONTO\\ExecutaOnDemand.bat')

# Crie um objeto PhotoImage com a imagem redimensionada
imagem_botao = ImageTk.PhotoImage(imagem_redimensionada)

# Adicione uma Label para mensagens de status
status_label = tk.Label(root, text="", font=("Roboto", 8), fg="blue")
status_label.place(x=20, y=394, width=170, anchor='nw')

# Labels e entradas para as pastas
nome_pasta_label = tk.Label(root, text="Nome da Pasta:")
nome_pasta_label.place(x=62, y=8, anchor='nw')

nome_pasta_entry = tk.Entry(root)
nome_pasta_entry.place(x=30, y=34, width=150, height=20, anchor='nw')

opcao_label = tk.Label(root, text="Opção:")
opcao_label.place(x=82, y=58, anchor='nw')

opcoes = [ "Total Contador Fit", "Total Contador", "Apenas AG", "Total Contador + AG", "AC e AG", "Apenas PONTO", "Total Contador + PONTO", "AC e PONTO"]

opcao_combobox = ttk.Combobox(root, values=opcoes, state="Total Contador")
opcao_combobox.current(0)  # Defina a opção padrão
opcao_combobox.place(x=30, y=88, width=150, height=20, anchor='nw')

criar_pastas_button = tk.Button(root, text="Criar Pastas", command=criar_pastas)
criar_pastas_button.place(x=68, y=119, anchor='nw')

# Configurar a barra de progresso
progress_frame = tk.Frame(root)
progress_frame.place(x=20, y=435)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, length=150, maximum=100)
progress_bar.grid(padx=10, pady=5)

# Labels e entradas para letras, dígitos e senha (inicialmente desabilitados)
letras_label = tk.Label(root, text="Insira as letras:")
letras_label.place(x=62, y=160, anchor='nw')
letras_label.config(state=tk.DISABLED)

letras_entry = tk.Entry(root)
letras_entry.place(x=30, y=190, width=150, height=20, anchor='nw')
letras_entry.config(state=tk.DISABLED)

cpf_cnpj_label = tk.Label(root, text="Insira os dígitos do CPF/CNPJ:")
cpf_cnpj_label.place(x=22, y=220, anchor='nw')
cpf_cnpj_label.config(state=tk.DISABLED)

cpf_cnpj_entry = tk.Entry(root)
cpf_cnpj_entry.place(x=30, y=254, width=150, height=20, anchor='nw')
cpf_cnpj_entry.config(state=tk.DISABLED)

# Crie a label para o campo de senha
senha_label = tk.Label(root, text="Senha para o Banco:")
senha_label.place(x=45, y=280, anchor='nw')
senha_label.config(state=tk.DISABLED)

# Crie o campo de entrada da senha
senha_cliente = tk.Entry(root, state=tk.DISABLED)
senha_cliente.place(x=30, y=314, anchor='nw')

# Crie o botão de gerar senha (ao lado do campo de senha)
gerar_senha_button = tk.Button(root, image=imagem_botao, command=gerar_senha, state=tk.DISABLED)
gerar_senha_button.place(x=160, y=310)

# Botão "Criar Arquivo .ini" (abaixo do campo de senha e botão "Gerar Senha")
criar_button = tk.Button(root, text="Registrar", command=lambda: criar_ini(progress_bar, progress_var, caminho_arquivo, caminho_arquivo_ag, caminho_arquivo_patrio, caminho_arquivo_ponto))
criar_button.place(x=73, y=354, anchor='nw')
criar_button.config(state=tk.DISABLED)

root.mainloop()