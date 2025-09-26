import os
import sys
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import logging
import pyodbc



# Função para configurar o logging
def setup_logging():
    log_filename = 'Optimus.log'

    # Cria um logger personalizado
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Remove handlers antigos, se existirem, para evitar duplicação de logs
    if logger.hasHandlers():
        logger.handlers.clear()

    # Cria um handler para escrever o log em um arquivo com codificação UTF-8
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)

    # Define o formato do log
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)

    # Adiciona o handler ao logger
    logger.addHandler(file_handler)

    # Adiciona o cabeçalho a cada execução
    log_header = (f"\n{'=' * 48} INÍCIO DO AGENDAMENTO: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} {'=' * 48}\n")

    # Escreve o cabeçalho no log
    with open(log_filename, 'a', encoding='utf-8') as log_file:
        log_file.write(log_header)

    return logger, log_filename

# Função para registrar o fim do processo
def log_fim_processo(log_filename):
    log_footer = (
        f"\n{'=' * 48} FIM DO AGENDAMENTO: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')} {'=' * 48}\n"
    )

    # Escreve o rodapé no log
    with open(log_filename, 'a', encoding='utf-8') as log_file:
        log_file.write(log_footer)


# Função para centralizar a janela
def centralizar_janela(root, largura, altura):
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2
    root.geometry(f"{largura}x{altura}+{x}+{y}")


# Função para obter o caminho absoluto do ícone
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# Função para atualizar o progresso na barra de progresso
def atualizar_feedback(label, progress_bar, mensagem, progresso):
    label.config(text=mensagem)
    label.update_idletasks()
    progress_bar['value'] = progresso
    progress_bar.update_idletasks()


# Função para ler o arquivo config.ini e verificar qual SGBD está sendo utilizado
def verificar_sgbd(config_file):
    # Abrir e ler o arquivo manualmente
    try:
        with open(config_file, 'r', encoding='utf-8-sig') as file:  # 'utf-8-sig' remove o BOM se presente
            linhas = file.readlines()
    except Exception as e:
        logging.error(f"Erro ao abrir o arquivo de configuração: {e}")
        raise

    sgbd = None
    em_settings = False

    # Log para verificação do conteúdo do arquivo
    # logging.debug(f"Conteúdo do arquivo config.ini:\n{''.join(linhas)}")

    # Percorre as linhas para encontrar a seção [Settings] e o valor de Sgbd
    for linha in linhas:
        linha_original = linha  # Manter a linha original para debugging
        linha = linha.strip()  # Remove espaços extras no início e no fim

        # Detecta o início da seção [Settings]
        if linha.lower() == '[settings]':
            em_settings = True
            logging.debug(f"Seção [Settings] encontrada na linha: {linha_original}")
        elif linha.startswith('[') and em_settings:  # Sai da seção [Settings] se encontrar outra seção
            logging.debug(f"Saindo da seção [Settings] ao encontrar: {linha_original}")
            break

        # Verifica se a chave Sgbd está dentro da seção [Settings]
        if em_settings and linha.lower().startswith('sgbd='):
            sgbd = linha.split('=', 1)[1].strip()  # Captura o valor de Sgbd
            logging.debug(f"Valor de Sgbd encontrado: {sgbd}")
            break

    # Verificar o valor de Sgbd encontrado
    if sgbd:
        sgbd = sgbd.upper()
        logging.debug(f"SGBD processado: {sgbd}")
        if sgbd == 'MSSQL':
            return 'MSSQL'
        elif sgbd == 'FIREBIRD':
            return 'FIREBIRD'
        else:
            logging.error(f"Valor inválido para 'Sgbd': {sgbd}")
            raise ValueError(f"Valor inválido para 'Sgbd': {sgbd}")
    else:
        messagebox.showwarning("Sgbd não encontrado na seção [Settings].")
        logging.error("Sgbd não encontrado na seção [Settings].")
        raise ValueError("Sgbd não encontrado na seção [Settings].")

# Função que realizará a verificação/comparação entre banco MSSQL e config.ini
# e fará a remoção da linha em config.ini conforme o banco esteja desativado
def verificar_todos_bancos_sql(config_file, conexao_sqlserver, subpasta):
    # Abre o arquivo config.ini e carrega todas as linhas
    with open(config_file, 'r', encoding='utf-8-sig') as file:
        config_linhas = file.readlines()

    # Filtrar bancos de dados ativos no SQL Server (não contém "_DESAT", "_desat", "_DESATIVADO" e etc.)
    cursor = conexao_sqlserver.cursor()
    cursor.execute("""
        SELECT name FROM sys.databases 
        WHERE name NOT LIKE '%_DESAT%' 
        AND name NOT LIKE '%_desat%' 
        AND name NOT LIKE '%_DESATIVADO%' 
        AND name NOT LIKE '% DESAT%'
        AND name NOT LIKE '%DESAT%'
        AND name NOT LIKE '%desativado%'""")
    bancos_ativos = [row[0].strip().upper() for row in cursor.fetchall()]
    logging.info(f"Bancos ativos encontrados: {bancos_ativos}")

    linhas_removidas = []
    linhas_atualizadas = []

    # Verificar cada linha do config.ini
    for linha in config_linhas:
        if 'DbName=' in linha:
            # Extrai o nome do banco de dados da linha (parte após 'DbName=')
            db_name = linha.split('DbName=')[1].split(',')[0].strip().upper()

            # Verifica se o nome do banco inclui a subpasta correta (por exemplo, "TES_PONTO_")
            if subpasta.upper() in db_name:
                logging.debug(f"Verificando banco do Sistema {subpasta}: {db_name}")

                # Verifica se o banco de dados está inativo (não está nos bancos ativos)
                if db_name not in bancos_ativos:
                    # Verifica se no banco SQL existe alguma versão do banco com sufixo "_DESAT", etc.
                    cursor.execute(f"""
                        SELECT name FROM sys.databases 
                        WHERE name LIKE '{db_name}_%' 
                        OR name LIKE '{db_name} DESAT%'  -- Adiciona a busca por " DESAT"
                        """)
                    banco_com_sufixo = cursor.fetchone()

                    if banco_com_sufixo:
                        banco_inativo = banco_com_sufixo[0].strip().upper()  # Captura o nome completo com o sufixo
                        logging.info(f"Banco inativo encontrado: {banco_inativo}")
                    else:
                        banco_inativo = db_name
                        logging.info(f"Banco inativo encontrado: {banco_inativo}")

                    linhas_removidas.append(linha)
                else:
                    logging.debug(f"Banco ativo: {db_name}")
                    linhas_atualizadas.append(linha)
            else:
                # Mantém as linhas de bancos que não pertencem à subpasta atual
                linhas_atualizadas.append(linha)
        else:
            # Linha que não contém DbName, mantém no arquivo
            linhas_atualizadas.append(linha)

    # Reescreve o arquivo config.ini com as linhas atualizadas (sem os bancos inativos)
    with open(config_file, 'w', encoding='utf-8-sig') as file:
        file.writelines(linhas_atualizadas)

    # Log das linhas removidas
    if linhas_removidas:
        logging.info(f"Linhas removidas do arquivo config.ini: {len(linhas_removidas)}")
        for linha_removida in linhas_removidas:
            logging.info(f"Linha removida: {linha_removida.strip()}")
    else:
        logging.info("Nenhuma linha irregular foi encontrada no config.ini.")
        print("Nenhuma linha irregular foi encontrada no config.ini.")

    return linhas_removidas


# Função que realizará a verificação/comparação entre banco FIREBIRD e config.ini
# e fará a remoção da linha em config.ini conforme o banco esteja desativado
def verificar_bancos_firebird(config_file, subpasta, base_dirs=('D:\\BDS', 'C:\\BDS')):
    linhas_removidas = []
    linhas_atualizadas = []
    sufixos_inativos = ['_DESAT.FDB', '_desat.fdb', '_DESAT.FDB', '_desat.FDB', '_DESAT.fdb']

    # Lê o arquivo config.ini preservando sua estrutura original
    with open(config_file, 'r', encoding='utf-8-sig') as file:
        linhas = file.readlines()

    # Flag para indicar se estamos na seção correta
    dentro_da_secao_operations = False

    for linha in linhas:
        # Detecta a seção [Operations] e verifica se estamos dentro dela
        if linha.strip() == "[Operations]":
            dentro_da_secao_operations = True
            linhas_atualizadas.append(linha)  # Mantemos a linha da seção
            continue
        elif linha.startswith('['):
            dentro_da_secao_operations = False  # Sair da seção se outra seção começar
            linhas_atualizadas.append(linha)
            continue

        # Apenas processa linhas dentro da seção [Operations]
        if dentro_da_secao_operations:
            if 'DbName=' in linha:
                # Extraímos o caminho do banco de dados da linha no config.ini
                inicio = linha.find('DbName=') + len('DbName=')
                fim = linha.find(',', inicio)
                caminho_banco_original = linha[inicio:fim] if fim != -1 else linha[inicio:].strip()

                # Verifica se o caminho inclui a subpasta correta
                if subpasta.upper() in caminho_banco_original.upper():
                    # Extraímos o diretório e o nome do banco
                    dir_banco = os.path.dirname(caminho_banco_original)
                    nome_banco = os.path.basename(caminho_banco_original)

                    # Verifica se existe o arquivo com o sufixo "_DESAT"
                    logging.debug(f"Verificando banco: {caminho_banco_original}")
                    banco_inativo = False
                    for sufixo in sufixos_inativos:
                        caminho_banco_desat = os.path.join(dir_banco, nome_banco.replace('.FDB', sufixo).replace('.fdb', sufixo))

                        if os.path.exists(caminho_banco_desat):
                            banco_inativo = True
                            logging.info(f"Banco inativo encontrado: {nome_banco} com sufixo {sufixo}")
                            break

                    # Se o banco está inativo, não adicionar a linha à lista de linhas atualizadas
                    if banco_inativo:
                        linhas_removidas.append(linha)
                    else:
                        linhas_atualizadas.append(linha)
                else:
                    # Mantém as linhas de bancos que não pertencem à subpasta atual
                    linhas_atualizadas.append(linha)
            elif linha.strip():  # Adiciona a linha apenas se não estiver vazia (incluindo espaços e \n)
                linhas_atualizadas.append(linha)
        else:
            # Linhas fora da seção ou que não sejam relacionadas a DbName, preserva todas
            linhas_atualizadas.append(linha)

    # Reescreve o arquivo config.ini com as linhas atualizadas
    with open(config_file, 'w', encoding='utf-8-sig') as file:
        file.writelines(linhas_atualizadas)

    # Log das linhas removidas
    if linhas_removidas:
        logging.info(f"Linhas removidas do arquivo config.ini: {len(linhas_removidas)}")
        print(f"Linhas removidas do arquivo config.ini: {len(linhas_removidas)}")
        for linha_removida in linhas_removidas:
            logging.info(f"Linha removida: {linha_removida.strip()}")
            print(f"Linha removida: {linha_removida.strip()}")
    else:
        logging.info("Nenhuma linha irregular foi encontrada no config.ini.")
        print("Nenhuma linha irregular foi encontrada no config.ini.")

    return linhas_removidas


# Função responsavel por realizar o direcionamemto entre as funções para MSSQL e FIREBIRD
def atualizar_configuracoes(subpasta):
    logging.info("Iniciando verificação de bancos e arquivo 'config.ini'...")
    print("Iniciando verificação de bancos e arquivo 'config.ini'...")
    config_file = os.path.join(f"D:\\BDS\\Atualiza\\CloudUp\\CloudUpCmd\\{subpasta}", "config.ini")
    # config_file = os.path.join(f"C:\\Atualiza\\CloudUp\\CloudUpCmd\\{subpasta}", "config.ini")

    # Verifica se o arquivo config.txt existe
    if not os.path.exists(config_file):
        messagebox.showwarning("Configuração não encontrada", f"O arquivo config.ini não foi encontrado em:\n{config_file}")
        logging.error(f"Arquivo config.ini não encontrado em:\n{config_file}")
        return

    # Verifica qual SGBD está sendo usado (MSSQL ou Firebird)
    sgbd = verificar_sgbd(config_file)
    

    if sgbd == 'MSSQL':
        servidores = [
            'GLAUDSON-DESKTO\\SQLEXPRESS',  # Servidor 1
            'FORTES-BD-SQL-3',              # Servidor 2
            'CONT-BD-SQL-01',               # Servidor 3
            'CONT-BD-SQL-02',
            'FORTES-BD-SQL-2',
            'CONT-BD-SQL-03',
            'FORT-POD2-SQL01',
            'FORT-CNT-BD-SQL\\CONTABILPOD3',
            'FORT-CON3-SQ-2',
            'FOR-CNT-BD-SQL-',
            'FOR-CNT-SQL-1\\FORCNTBDSQL1',
            'FOR-CNT-BD-SQL5',
            'FORT-CNT-BDSQ-6'
        ]

        def testar_conexao_sqlserver(servidor):
            try:
                logging.info(f"Tentando conectar ao servidor: {servidor}")
                conexao = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={servidor};Trusted_Connection=yes;', timeout=5)
                logging.info(f"Conectado com sucesso ao servidor: {servidor}")
                return conexao, servidor
            except pyodbc.Error as e:
                logging.error(f"Falha ao conectar ao servidor: {servidor}. Erro: {e}")
                return None, servidor

        def conectar_servidores(servidores):
            conexao_sqlserver = None
            servidor_conectado = None

            # Usar ThreadPoolExecutor para testar conexões simultaneamente
            with ThreadPoolExecutor(max_workers=len(servidores)) as executor:
                # Submeter as conexões para todos os servidores
                futuros = {executor.submit(testar_conexao_sqlserver, servidor): servidor for servidor in servidores}

                # Processar as conexões assim que uma delas for bem-sucedida
                for futuro in as_completed(futuros):
                    conexao, servidor = futuro.result()
                    if conexao:
                        conexao_sqlserver = conexao
                        servidor_conectado = servidor
                        break  # Parar assim que encontrar um servidor válido

            if conexao_sqlserver:
                logging.info(f"Conexão estabelecida com o servidor: {servidor_conectado}")
                print(f"Conexão estabelecida com o servidor: {servidor_conectado}")
            else:
                logging.error("Falha ao conectar a qualquer servidor SQL.")
                print("Falha ao conectar a qualquer servidor SQL.")

            return conexao_sqlserver

        # Chama a função que tenta conectar aos servidores simultaneamente
        conexao_sqlserver = conectar_servidores(servidores)

        # Verifica se a conexão foi estabelecida com sucesso
        if conexao_sqlserver:
            try:
                # Chama a função que verifica os bancos no servidor conectado
                linhas_removidas_sql = verificar_todos_bancos_sql(config_file, conexao_sqlserver, subpasta)
                conexao_sqlserver.close()
                quantidade_removida_sql = len(linhas_removidas_sql)

                logging.info("Verificação concluída")
                print("Verificação concluída")

                # Log da listagem completa
                if linhas_removidas_sql:
                    print(f"Linhas removidas do arquivo config.ini: {quantidade_removida_sql}\n" + ''.join(linhas_removidas_sql))

                # Exibindo a quantidade de linhas removidas
                if quantidade_removida_sql > 0:
                    messagebox.showinfo("Resultado MSSQL", f"Foram removidas {quantidade_removida_sql} linha(s) do arquivo config.ini.\n\n->Confira o 'Log' para verificar quais linhas foram removidas")
                    return True
                else:
                    messagebox.showinfo("Resultado MSSQL", "Nenhuma linha irregular foi encontrada no config.ini.")
                    return True
                    
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar os bancos de dados: {e}")
                logging.error(f"Erro ao processar os bancos de dados: {e}")
                print(f"Erro ao processar os bancos de dados: {e}")
                return False
        else:
            messagebox.showerror("Erro", "Falha ao conectar a qualquer servidor SQL.")
            logging.error("Falha ao conectar a qualquer servidor SQL.")
            return False


    elif sgbd == 'FIREBIRD':
        try:
            linhas_removidas_firebird = verificar_bancos_firebird(config_file, subpasta)
            quantidade_removida_fb = len(linhas_removidas_firebird)

            logging.info("Verificação concluída")
            print("Verificação concluída")

            # Log da listagem completa
            if linhas_removidas_firebird:
                print(f"Linhas removidas do arquivo config.ini: {quantidade_removida_fb}\n" + ''.join(linhas_removidas_firebird))

            # Exibindo a quantidade de linhas removidas
            
            if quantidade_removida_fb > 0:
                messagebox.showinfo("Resultado Firebird", f"Foram removidas {quantidade_removida_fb} linha(s) do arquivo config.ini.\n\n->Confira o 'Log' para verificar quais linhas foram removidas")
                return True

            else:
                messagebox.showinfo("Resultado Firebird", "Nenhuma linha irregular foi encontrada no config.ini.")
                logging.info("Nenhuma linha irregular foi encontrada no config.ini.")
                print("Nenhuma linha irregular foi encontrada no config.ini.")
                return True

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o Firebird: {e}")
            logging.error(f"Erro ao processar o Firebird: {e}")
            print(f"Erro ao processar o Firebird: {e}")
            return False
            
    else:
        messagebox.showwarning("Aviso", "SGBD não identificado ou não suportado.")
        logging.error("SGBD não identificado ou não suportado.")
        print("SGBD não identificado ou não suportado.")
        return False

    
# Função auxiliar que fará a leitura do arquivo 'config.txt' para a observação das Subpastas
def carregar_configuracao(config_path):
    config_dict = {}
    try:
        with open(config_path, 'r') as config_file:
            for line in config_file:
                line = line.strip()  # Remove espaços em branco no início e no fim

                # Ignorar linhas vazias
                if not line:
                    continue

                # Ignorar títulos ou seções identificados com '[]'
                if line.startswith('[') and line.endswith(']'):
                    logging.info(f"Lendo Arquivo: Config.txt")
                    logging.info(f"Verificando seção: {line}")
                    continue

                # Processar linhas com ':' (chave: valor)
                if ':' in line:
                    try:
                        chave, valor = line.split(':', 1)
                        chave = chave.strip("' ")
                        # Em vez de converter para int, agora vamos usar uma lista de subpastas
                        subpastas = [v.strip() for v in valor.split(',')]
                        config_dict[chave] = subpastas
                    except ValueError as e:
                        logging.error(f"Erro ao processar linha: {line}. Erro: {e}")
                else:
                    logging.warning(f"Linha ignorada (não contém ':'): {line}")
    except FileNotFoundError as e:
        logging.error(f"Arquivo de configuração não encontrado: {config_path}. Erro: {e}")
    except Exception as e:
        logging.error(f"Erro ao carregar o arquivo de configuração: {e}")
    return config_dict


# Função para verificar e criar subpastas, e remover pastas excedentes de acordo com o arquivo config.txt
def verificar_e_criar_ou_remover_subpastas(base_path, subpastas_selecionadas, config_path):
    configuracao_atualizadores = carregar_configuracao(config_path)

    subpastas_selecionadas = subpasta_var.get()

    # Certificação de que subpastas_selecionadas é uma lista de strings
    if isinstance(subpastas_selecionadas, str):
        subpastas_selecionadas = [subpastas_selecionadas]  # Converte string única para lista

    logging.info(f"Subpasta selecionada: {subpastas_selecionadas}")
    print(f"Subpasta selecionada: {subpastas_selecionadas}")

    for subpasta in subpastas_selecionadas:
        if subpasta in configuracao_atualizadores:
            subpastas_esperadas = configuracao_atualizadores[subpasta]

            # Verificar quais subpastas existem na pasta
            atualizadores_dir = os.path.join(base_path, subpasta, 'Atualizadores')
            subpastas_existentes = sorted([d for d in os.listdir(atualizadores_dir) if os.path.isdir(os.path.join(atualizadores_dir, d))])

            # Criar as subpastas que estão faltando
            for subpasta_nome in subpastas_esperadas:
                subpasta_path = os.path.join(atualizadores_dir, subpasta_nome)
                if not os.path.exists(subpasta_path):
                    os.makedirs(subpasta_path)
                    logging.info(f"Subpasta criada: {subpasta_path}")
                    print(f"Subpasta criada: {subpasta_path}")
                else:
                    logging.info(f"Subpasta já existe: {subpasta_path}")
                    print(f"Subpasta já existe: {subpasta_path}")

            # Remover subpastas que não estão na lista esperada
            for subpasta_existente in subpastas_existentes:
                if subpasta_existente not in subpastas_esperadas:
                    subpasta_path_excedente = os.path.join(atualizadores_dir, subpasta_existente)
                    shutil.rmtree(subpasta_path_excedente)
                    logging.info(f"Subpasta irregular removida: {subpasta_path_excedente}")
                    print(f"Subpasta irregular removida: {subpasta_path_excedente}")

        else:
            logging.info(f"A subpasta {subpasta} não está configurada no arquivo.")
            print(f"A subpasta {subpasta} não está configurada no arquivo.")


# Função para copiar arquivos com progresso na barra
def copytree_with_progress(source, destination, update_progress):
    # Calcula o tamanho total de todos os arquivos na pasta de origem
    total_size = sum(
        os.path.getsize(os.path.join(dirpath, filename))
        for dirpath, _, filenames in os.walk(source)
        for filename in filenames
    )
    
    progress = [0]  # Progresso compartilhado como lista para ser mutável

    def copy_file(src, dst):
        try:
            shutil.copy2(src, dst)
            file_size = os.path.getsize(src)
            progress[0] += file_size
            # Atualiza o progresso usando valores numéricos
            update_progress(float(progress[0]), float(total_size))
        except Exception as e:
            logging.error(f"Erro ao copiar o arquivo {src} para {dst}: {e}")
            raise

    def recursive_copy(src_dir, dst_dir):
        if not os.path.exists(dst_dir):
            os.makedirs(dst_dir)
        for item in os.listdir(src_dir):
            src_path = os.path.join(src_dir, item)
            dst_path = os.path.join(dst_dir, item)
            if os.path.isdir(src_path):
                recursive_copy(src_path, dst_path)
            else:
                copy_file(src_path, dst_path)

    # Inicia a cópia recursiva
    try:
        recursive_copy(source, destination)
    except Exception as e:
        logging.error(f"Erro durante a cópia da árvore de diretórios de {source} para {destination}: {e}")
        raise


#Função que realizará a exclusão dos arquivos em .txt
def delete_files_and_folders(path, extensions=None):

    if extensions is None:
        extensions = []

    if not os.path.exists(path):
        print(f"O caminho especificado não existe:\n{path}")
        logging.error(f"O caminho especificado não existe:\n{path}")
        return False

    for name in os.listdir(path):
        file_path = os.path.join(path, name)
        try:
            if os.path.isfile(file_path) and (any(name.endswith(ext) for ext in extensions) or not extensions):
                os.remove(file_path)
                print(f"Arquivo {file_path} excluído.")
                logging.info(f"Arquivo {file_path} excluído.")
        except Exception as e:
            print(f"Erro ao excluir {file_path}: {e}")
            logging.error(f"Erro ao excluir {file_path}: {e}")

    logging.info(f"Exclusão de arquivos '.txt' concluido.")
    print(f"Exclusão de arquivos '.txt' concluido.")
    return True


# Função para deletar e copiar subpastas simultaneamente
def delete_and_copy_for_subfolders(base_path, subpasta, source_path, update_progress):

    config_path = os.path.join(base_path, subpasta, 'Atualizadores', 'config.txt')

    # Verifica se o arquivo config.txt existe
    if not os.path.exists(config_path):
        messagebox.showwarning("Configuração não encontrada", f"O arquivo config.txt não foi encontrado em:\n{config_path}")
        logging.error(f"Arquivo config.txt não encontrado em:\n{config_path}")
        return False

    # Criação e verificação das subpastas conforme config.txt
    verificar_e_criar_ou_remover_subpastas(base_path, [subpasta], config_path)

    # Atualizadores Path para as subpastas permitidas
    atualizadores_dir = os.path.join(base_path, subpasta, 'Atualizadores')
    subpastas_existentes = sorted([d for d in os.listdir(atualizadores_dir) if os.path.isdir(os.path.join(atualizadores_dir, d))])

    # Adiciona todas as subpastas existentes para serem copiadas
    subpasta_paths = [os.path.join(atualizadores_dir, subpasta_existente) for subpasta_existente in subpastas_existentes]

    # Caso não haja subpastas para o sistema especificado, apresenta um aviso e retorna
    if not subpasta_paths:
        messagebox.showwarning("Caminho não encontrado", f"O caminho especificado não foi encontrado:\n{atualizadores_dir}")
        logging.error(f"O caminho especificado não foi encontrado:\n{atualizadores_dir}")
        return False

    # Verificar se a pasta de origem está vazia
    if not os.path.exists(source_path) or not os.listdir(source_path):
        messagebox.showwarning("Pasta de origem vazia", f"A pasta de origem está vazia:\n\n{source_path}\n\n-> Informe ao Administrador do agendamento para\nsolucionar o ocorrido")
        logging.error(f"A pasta de origem está vazia:\n{source_path}\n-> Informe ao Administrador do agendamento para\nsolucionar o ocorrido")
        print(f"A pasta de origem está vazia ou não foi encontrada:\n{source_path}")
        return False

    def process_subfolder(subpasta_path):
        try:
            # Remove arquivos e diretórios específicos
            for item in os.listdir(subpasta_path):
                item_path = os.path.join(subpasta_path, item)
                if os.path.isfile(item_path):
                    os.remove(item_path)
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)

            dirs_to_remove = ['AutenticaNFeWeb', 'locales', 'Logs']
            for dir_name in dirs_to_remove:
                dir_path = os.path.join(subpasta_path, dir_name)
                if os.path.exists(dir_path):
                    shutil.rmtree(dir_path)

            # Copiar arquivos para a subpasta
            copytree_with_progress(source_path, subpasta_path, update_progress)
            print("Realizando copia dos arquivos...")

        except Exception as e:
            logging.error(f"Erro ao processar a subpasta {subpasta_path}: {e}")
            raise

    # Executa a cópia das subpastas em paralelo com limitação no número de threads
    max_workers = min(5, len(subpasta_paths))  # Limita a 5 threads ou ao número de subpastas, o que for menor
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_subfolder, path) for path in subpasta_paths]
        for future in futures:
            try:
                future.result()  # Aguarda todas as threads terminarem
            except Exception as e:
                logging.error(f"Erro ao processar uma das subpastas: {e}")
                return False

    logging.info("Copia de arquivos concluída.")
    print("Copia de arquivos concluída.")
    return True


# Função para remover um diretório e todo o seu conteúdo
def remove_directory(path):
    if os.path.exists(path):
        shutil.rmtree(path)

    logging.info("Remoção de diretórios e arquivos concluida.")
    print("Remoção de diretórios e arquivos concluida.")
    return True


# Função para descomentar linhas na seção [Operations] do arquivo config.ini
def descomentar_linhas_operations(subpasta, root):
    # config_ini_path = os.path.join(f"C:\\Atualiza\\CloudUp\\CloudUpCmd\\{subpasta}", "config.ini")
    config_ini_path = os.path.join(f"D:\\BDS\\Atualiza\\CloudUp\\CloudUpCmd\\{subpasta}", "config.ini")

    # Verifica se o arquivo config.ini existe
    if not os.path.exists(config_ini_path):
        messagebox.showwarning("Configuração não encontrada", f"O arquivo config.ini não foi encontrado em:\n{config_ini_path}")
        logging.error(f"Arquivo config.ini não encontrado em:\n{config_ini_path}")
        return False

    logging.info("Descomentando linhas no arquivo config.ini...")
    print("Descomentando linhas no arquivo config.ini...")

    with open(config_ini_path, 'r') as config_file:
        lines = config_file.readlines()

    operations_section_start = -1
    for i, line in enumerate(lines):
        if line.strip() == '[Operations]':
            operations_section_start = i
            break

    operations_section_end = len(lines)
    for i in range(operations_section_start + 1, len(lines)):
        if lines[i].strip().startswith('['):
            operations_section_end = i
            break

    lines_count = 0
    for i in range(operations_section_start + 1, operations_section_end):
        if operations_section_start < i < operations_section_end:
            lines[i] = lines[i].lstrip(';').lstrip(' ; ').rstrip(';')
            lines_count += 1

    with open(config_ini_path, 'w') as config_file:
        config_file.writelines(lines)

    # Copia a quantidade na area de transferencia do usuario
    root.clipboard_clear()
    root.clipboard_append(str(lines_count))
    root.update()

    messagebox.showinfo("Descomentando Linhas", f"Linhas descomentadas com sucesso.\nUm total de {lines_count} Bancos foram encontrados dentro de config.ini.\n\nO total foi copiado para sua área de transferência.")
    print(f"Linhas descomentadas com sucesso.\nUm total de {lines_count} clientes foram encontrados dentro de config.ini.")
    logging.info(f"Linhas descomentadas com sucesso. => Um total de {lines_count} clientes foram encontrados dentro de config.ini.")
    return True


# Função para formatar a data no formato DD/MM/YYYY
def format_data_dd_mm_yyyy(dia):
    return dia.strftime("%d/%m/%Y")


# Função para formatar a data no formato MM/DD/YYYY
def format_data_mm_dd_yyyy(dia):
    return dia.strftime("%m/%d/%Y")


# Função para exibir um alerta de erro
def exibir_alerta_erro(titulo, mensagem):
    messagebox.showwarning(titulo, mensagem)

# Função para tentar alterar a data da tarefa com dois formatos
def alterar_data_tarefa(nome_tarefa, usuario, senha, status_label, progress_bar):
    logging.info(f"Alterar a data da tarefa:{nome_tarefa}")
    print(f"Alterar a data da tarefa:{nome_tarefa}")

    data_disparo = datetime.datetime.now() + datetime.timedelta(days=1)

    if data_disparo.weekday() in [5, 6]:
        dias_para_segunda = (7 - data_disparo.weekday())
        data_disparo += datetime.timedelta(days=dias_para_segunda)

    formato_dd_mm = format_data_dd_mm_yyyy(data_disparo)
    formato_mm_dd = format_data_mm_dd_yyyy(data_disparo)

    comando_dd_mm = [
        'schtasks', '/Change', '/TN', nome_tarefa, '/SD', formato_dd_mm, '/RU', usuario, '/RP', senha
    ]
    comando_mm_dd = [
        'schtasks', '/Change', '/TN', nome_tarefa, '/SD', formato_mm_dd, '/RU', usuario, '/RP', senha
    ]

    # Tentar alterar a tarefa usando o formato dd/mm/yyyy
    try:
        resultado = subprocess.run(comando_dd_mm, check=True, shell=True, capture_output=True, text=True)
        if resultado.returncode == 0:
            atualizar_feedback(status_label, progress_bar, f"Tarefa alterada com sucesso\n para {formato_dd_mm}.", 90)
            logging.info(f"Data da tarefa alterada com sucesso usando o formato {formato_dd_mm}.")
            print(f"Data da tarefa alterada com sucesso usando o formato {formato_dd_mm}.")
            return True
    except subprocess.CalledProcessError as e:
        exibir_alerta_erro("Erro ao alterar a data da tarefa", f"Erro ao alterar a data da tarefa para {formato_dd_mm}\n\n{e.stderr}\n\nA tarefa precisa ser alterada manualmente.")
        atualizar_feedback(status_label, progress_bar, "Processo concluído com erro.\nA tarefa precisa ser alterada manualmente.", 90)
        logging.info("Processo concluído com erro. A tarefa precisa ser alterada manualmente.")
        print(f"Erro ao alterar a data da tarefa para {formato_dd_mm}\n\n{e.stderr}")
        return False

    # Tentar alterar a tarefa usando o formato mm/dd/yyyy
    try:
        resultado = subprocess.run(comando_mm_dd, check=True, shell=True, capture_output=True, text=True)
        if resultado.returncode == 0:
            atualizar_feedback(status_label, progress_bar, f"Data da tarefa alterada com sucesso usando o formato {formato_mm_dd}.", 90)
            logging.info(f"Data da tarefa alterada com sucesso usando o formato {formato_mm_dd}.")
            print(f"Data da tarefa alterada com sucesso usando o formato {formato_mm_dd}.")
            return True
    except subprocess.CalledProcessError as e:
        exibir_alerta_erro("Erro ao alterar a data da tarefa", f"Erro ao alterar a data da tarefa para {formato_mm_dd}\n\n{e.stderr}\n\nA tarefa precisa ser alterada manualmente.")
        atualizar_feedback(status_label, progress_bar, "Processo concluído com erro.\nA tarefa precisa ser alterada manualmente.", 90)
        logging.info("Processo concluído com erros. A tarefa precisa ser alterada manualmente.")
        print(f"Erro ao alterar a data da tarefa para {formato_mm_dd}\n{e.stderr}")
        return False


# Função principal para processar as funções e atualizar o feedback
def processar():

    subpasta = subpasta_var.get()  # Obtém o sistema selecionado
    radio_selecionado = radio_value.get()  # Obtém o valor da tarefa selecionada
    versao = versao_var.get()  # Obtém a versão selecionada
    nome_tarefa = f"{subpasta} {radio_value.get()}"  # Obtém a tarefa a ser alterada
    usuario = r"Fortes\Parceiro"
    senha = "f1ltHlBA3Hl7"

    # Verifica se o source_path existe
    base_path = r'D:\BDS\Atualiza\CloudUp\CloudUpCmd'
    # base_path = r'C:\Atualiza\CloudUp\CloudUpCmd'

    # Definir o caminho de origem conforme a seleção
    if subpasta == 'AC':
        source_path = os.path.join(r'\\GLAUDSON-NOTE\Rede_ParaTestes\Shere\utilitarios\AtualizaNuvem_Op', subpasta, versao)
        # source_path = os.path.join(r'\\FORTES-RPS-02\Share\utilitarios\AtualizaNuvem_Op', subpasta, versao)

    else:
        source_path = os.path.join(r'\\GLAUDSON-NOTE\Rede_ParaTestes\Shere\utilitarios\AtualizaNuvem_Op', subpasta)
        # source_path = os.path.join(r'\\FORTES-RPS-02\Share\utilitarios\AtualizaNuvem_Op', subpasta)


    subpasta_path = os.path.join(base_path, subpasta)

    def executar_processos():
        try:
            # Chama a função setup_logging no início do script
            logger, log_filename = setup_logging()

            # Atualiza prints de depuração
            print(f"Subpasta selecionada: {subpasta}")
            print(f"Valor do rádio selecionado: {radio_selecionado if radio_selecionado != 'None' else 'Nenhum selecionado'}")
            print(f"Versão selecionada: {versao}")

            # Verifica se a versão foi selecionada
            if versao == "":
                messagebox.showwarning("Seleção Inválida", "Por favor, selecione uma versão.")
                logging.warning("Versão não selecionada.")
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)            
                return False

            # Verifica se subpasta ou radio_value não estão preenchidos corretamente
            if subpasta == "" or radio_selecionado == "None":
                messagebox.showwarning("Seleção Inválida", "Por favor, selecione uma subpasta e uma tarefa.")
                logging.warning("Seleção inválida. Subpasta ou tarefa não selecionada.")
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False

            # Verifica se o caminho da subpasta existe
            if not os.path.exists(subpasta_path):
                messagebox.showwarning("Caminho não encontrado", f"O caminho especificado não foi encontrado:\n{subpasta_path}")
                logging.error(f"O caminho especificado não foi encontrado: {subpasta_path}")
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False

            # Verifica se o caminho de origem existe
            if not os.path.exists(source_path):
                messagebox.showwarning("Caminho não encontrado", f"O caminho de origem não foi encontrado:\n{source_path}")
                logging.error(f"O caminho de origem não foi encontrado: {source_path}")
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False
            
            if not os.path.exists(source_path) or not os.listdir(source_path):
                messagebox.showwarning("Pasta de origem vazia", f"A pasta de origem está vazia:\n\n{source_path}\n\n-> Informe ao Administrador do agendamento para\nsolucionar o ocorrido")
                logging.error(f"A pasta de origem está vazia: {source_path}")
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return

            # Atualizar feedback no início do processo
            atualizar_feedback(status_label, progress_bar, "Iniciando o processo...", 0)
            logging.info("Iniciando o processo...")

            # Verificar e atualizar configurações
            atualizar_feedback(status_label, progress_bar, "Iniciando verificação de\nbancos e arquivo 'config.ini'...", 5)
            sucesso_atualizacao = atualizar_configuracoes(subpasta)
            if not sucesso_atualizacao:
                logging.error("Erro ao verificar configurações. Interrompendo o processo.")
                atualizar_feedback(status_label, progress_bar, "Erro ao verificar configurações\nInterrompendo o processo.", 0)
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False
            atualizar_feedback(status_label, progress_bar, "Verificação Concluída", 10)

            # Deletar arquivos .txt
            atualizar_feedback(status_label, progress_bar, "Deletando arquivos .txt...", 15)
            logging.info("Deletando arquivos .txt...")
            sucesso_delecao = delete_files_and_folders(subpasta_path, extensions=['.txt'])
            if not sucesso_delecao:
                logging.error("Erro ao deletar arquivos .txt. Interrompendo o processo.")
                atualizar_feedback(status_label, progress_bar, "Erro ao deletar arquivos .txt.\nInterrompendo o processo.", 0)
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False

            # Remover diretórios de logs
            atualizar_feedback(status_label, progress_bar, "Removendo diretórios de logs...", 20)
            logging.info("Removendo diretórios e arquivos...")
            sucesso_remocao_logs = remove_directory(os.path.join(subpasta_path, 'Atualizadores', subpasta, 'Logs'))
            if not sucesso_remocao_logs:
                logging.error("Erro ao remover diretórios de logs. Interrompendo o processo.")
                atualizar_feedback(status_label, progress_bar, "Erro ao remover diretórios de logs.\nInterrompendo o processo.", 0)
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False

            # Remover diretórios AutenticaNFeWeb e locales
            atualizar_feedback(status_label, progress_bar, "Removendo diretórios\nAutenticaNFeWeb e locales...", 30)
            sucesso_remocao_autentica = remove_directory(os.path.join(subpasta_path, 'Atualizadores', subpasta, 'AutenticaNFeWeb'))
            sucesso_remocao_locales = remove_directory(os.path.join(subpasta_path, 'Atualizadores', subpasta, 'locales'))

            if not (sucesso_remocao_autentica and sucesso_remocao_locales):
                logging.error("Erro ao remover diretórios AutenticaNFeWeb ou locales. Interrompendo o processo.")
                atualizar_feedback(status_label, progress_bar, "Erro ao remover diretórios\nAutenticaNFeWeb ou locales.\nInterrompendo o processo.", 0)
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False    
            

            # Verificar arquivos
            atualizar_feedback(status_label, progress_bar, "Verificando arquivos...", 35)
            logging.info("Verificando arquivos...")
            print("Verificando arquivos...")

            # Transferir arquivos
            logging.info("Copiando arquivos...")
            sucesso_copia = delete_and_copy_for_subfolders(base_path, subpasta, source_path, lambda p, t: atualizar_feedback(status_label, progress_bar, "Copiando arquivos...\nAguarde", 35 + (50 * int(p) / int(t))))
            if not sucesso_copia:
                logging.error("Erro ao copiar arquivos. Interrompendo o processo.")
                atualizar_feedback(status_label, progress_bar, "Erro ao copiar arquivos.\nInterrompendo o processo.", 0)
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False

            # Descomentar linhas no arquivo config.ini
            atualizar_feedback(status_label, progress_bar, "Descomentando linhas\nno arquivo config.ini...", 85)
            sucesso_descomentar = descomentar_linhas_operations(subpasta, root)
            if not sucesso_descomentar:
                logging.error("Erro ao descomentar linhas no arquivo config.ini. Interrompendo o processo.")
                atualizar_feedback(status_label, progress_bar, "Erro ao descomentar linhas no arquivo config.ini. Interrompendo o processo.", 0)
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False
            atualizar_feedback(status_label, progress_bar, "Linhas descomentadas com sucesso.", 85)

            # Alterar a data da tarefa
            sucesso_alterar_tarefa = alterar_data_tarefa(nome_tarefa, usuario, senha, status_label, progress_bar)
            if not sucesso_alterar_tarefa:
                atualizar_feedback(status_label, progress_bar, "Processo concluído com erro.\nA tarefa precisa ser alterada manualmente.", 90)
                logging.error("Erro ao alterar a data da tarefa. Processo concluído com erro.")
                print("Processo concluído com erro.\nA tarefa precisa ser alterada manualmente.")
                # Adiciona o rodapé ao fim do processo, independentemente de erros
                log_fim_processo(log_filename)
                return False

            # Finaliza o processo com sucesso
            atualizar_feedback(status_label, progress_bar, "Processo concluído!\nAtualização agendada com sucesso", 100)
            logging.info("Processo concluído.\nAtualização agendada com sucesso")
            print("Processo concluído.\nAtualização agendada com sucesso")
            # Adiciona o rodapé ao fim do processo, independentemente de erros
            log_fim_processo(log_filename)
            return True

        except Exception as e:
            logging.error(f"Ocorreu um erro inesperado: {e}")
            atualizar_feedback(status_label, progress_bar, f"Erro inesperado: {e}", 0)
            messagebox.showerror("Erro", f"Ocorreu um erro inesperado: {e}")
            # Adiciona o rodapé ao fim do processo, independentemente de erros
            log_fim_processo(log_filename)
            return False

    thread = threading.Thread(target=executar_processos)
    thread.start()


# Configuração da interface gráfica
root = tk.Tk()
root.title("Optimus")
root.resizable(False, False)
centralizar_janela(root, 250, 350)

# Obter o caminho absoluto do arquivo (usando a função resource_path)
caminho_icon = resource_path("optimus-face.ico")

# Verificar se o arquivo existe
if os.path.exists(caminho_icon):
    imagem_original = Image.open(caminho_icon)
    largura_botao = 55
    altura_botao = 40

    imagem_redimensionada = imagem_original.resize((largura_botao, altura_botao), Image.Resampling.LANCZOS)
    imagem_converted = ImageTk.PhotoImage(imagem_redimensionada)
    root.iconphoto(True, imagem_converted)
else:
    print(f"O arquivo {caminho_icon} não foi encontrado.")

# Subpastas e valores para botões de rádio
subpasta_var = tk.StringVar(value="")
radio_value = tk.StringVar(value="None")
versao_var = tk.StringVar(value="versao7")

# Frame principal para organizar os widgets centralmente
frame_principal = tk.Frame(root)
frame_principal.pack(padx=10, pady=10)

# Label de seleção de subpasta
subpasta_label = tk.Label(frame_principal, text="Selecione o Sistema:")
subpasta_label.grid(row=0, column=0, columnspan=2, pady=(0, 2))

# Combobox para seleção de subpasta
subpasta_combobox = ttk.Combobox(frame_principal, textvariable=subpasta_var, width=10)
subpasta_combobox['values'] = ('AC', 'AG', 'PATRIO', 'PONTO')
subpasta_combobox.grid(row=1, column=0, columnspan=2, pady=(0, 10))

# Frame para a seleção da versão (inicialmente oculto)
frame_versao = tk.Frame(frame_principal)

# Label de seleção de versão
versao_label = tk.Label(frame_versao, text="Selecione a Versão:")
versao_label.grid(row=0, column=0, columnspan=2, pady=(0, 2))

# Combobox para seleção de versão
versao_combobox = ttk.Combobox(frame_versao, textvariable=versao_var, width=10)
versao_combobox['values'] = ('versao7', 'versao8')
versao_combobox.grid(row=1, column=0, columnspan=2, pady=(0, 10))

# Frame para os botões de rádio
radio_frame = tk.Frame(frame_principal)
radio_frame.grid(row=3, column=0, columnspan=2, pady=(10, 10))

# Botões de rádio
radio_exe = tk.Radiobutton(radio_frame, text="EXE", variable=radio_value, value="EXE")
radio_full = tk.Radiobutton(radio_frame, text="FULL", variable=radio_value, value="FULL")
radio_exe.grid(row=0, column=0, padx=10)
radio_full.grid(row=0, column=1, padx=10)

# Botão para iniciar o processo
botao_iniciar = tk.Button(frame_principal, image=imagem_converted, command=processar, borderwidth=0)
botao_iniciar.grid(row=4, column=0, columnspan=2, pady=(10, 0))

# Label para identificação do botão "Iniciar"
label_iniciar = tk.Label(frame_principal, text="Iniciar Processo")
label_iniciar.grid(row=5, column=0, columnspan=2, pady=(2, 10))

# Label de status e barra de progresso
status_label = tk.Label(frame_principal, text="", anchor="center", fg="blue")
status_label.grid(row=6, column=0, columnspan=2, pady=(10, 5))

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(frame_principal, orient="horizontal", length=180, mode="determinate")
progress_bar.grid(row=7, column=0, columnspan=2, pady=(0, 10))



# Atualize a função que mostra/oculta a combobox da versão
def atualizar_visibilidade_versao(event):
    subpasta_selecionada = subpasta_var.get()
    if subpasta_selecionada == 'AC':
        frame_versao.grid(row=2, column=0, columnspan=2, pady=(5, 10))  # Mostra o frame da versão
    else:
        frame_versao.grid_remove()  # Oculta o frame da versão

# Associe a função ao evento de seleção do combobox da subpasta
subpasta_combobox.bind("<<ComboboxSelected>>", atualizar_visibilidade_versao)

# Iniciar o loop da interface gráfica
root.mainloop()
