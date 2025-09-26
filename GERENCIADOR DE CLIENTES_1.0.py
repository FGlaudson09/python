
# GERENCIADOR DE CLIENTES

import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk, messagebox
from validate_docbr import CPF, CNPJ
import threading
import re
import fdb
import time



def listar_clientes():
    global lista_clientes
    if lista_clientes is None:
        return
    lista_clientes.delete(0, tk.END)
    try:
        con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
        cur = con.cursor()


        # Imprima a consulta SQL para depuração
        query = 'SELECT AG, NOME, CPF_CNPJ FROM CLIENTES'
        print("Consulta SQL:", query)

        cur.execute(query)
        rows = cur.fetchall()

        #utltimo cliente adicionado no topo da lista
        rows.reverse()

        for row in rows:
            print("Row:", row)
            lista_clientes.insert(tk.END, f'AG: {row[0]}, CLIENTE: {row[1]}, CPF/CNPJ: {row[2]}')

    except Exception as e:
        print(f"Erro ao listar clientes: {e}")

    finally:
        if cur:
            cur.close()
        if con:
            con.close()

def salvar_cliente(entry_ag,entry_nome,entry_nome_pasta,entry_pod,entry_cpf_cnpj,entry_liberacao,entry_ambiente,entry_responsavel,entry_acessos,tipo_cliente_var,entry_contato,entry_email,entry_produto,situacao_var,atualizacao_var,sgdb_var):

    try:
        con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
        cur = con.cursor()

        ag = entry_ag.get()
        nome = entry_nome.get()
        nome_pasta = entry_nome_pasta.get()
        pod = entry_pod.get()
        cpf_cnpj = entry_cpf_cnpj.get()
        liberacao = entry_liberacao.get()
        liberacao = formatar_data(liberacao)
        ambiente = entry_ambiente.get()
        responsavel = entry_responsavel.get()
        acessos = entry_acessos.get()
        tipo_cliente = tipo_cliente_var.get()
        contato = entry_contato.get()
        email = entry_email.get()
        produto = entry_produto.get()
        situacao = situacao_var.get()
        atualizacao = atualizacao_var.get()
        sgdb = sgdb_var.get()

        liberacao = formatar_data(liberacao)

        if ag and nome and nome_pasta and pod and tipo_cliente and cpf_cnpj and liberacao and ambiente and responsavel and acessos and tipo_cliente_var and contato and email and produto and situacao_var and atualizacao_var and sgdb_var:

            # Verificar se o ag do cliente ja existe
            cur.execute('SELECT COUNT(*) FROM CLIENTES WHERE AG = ?', (ag,))
            count = cur.fetchone()[0]

            # Verificar se o CPF/CNPJ já existe
            cur.execute('SELECT COUNT(*) FROM CLIENTES WHERE CPF_CNPJ = ?', (cpf_cnpj,))
            cpf_cnpj_count = cur.fetchone()[0]


            if count > 0:
                cur.execute(f'UPDATE CLIENTES SET NOME = ?, NOME_DA_PASTA = ?, POD = ?, CPF_CNPJ = ?, DATA_DE_LIBERACAO = ?, AMBIENTE = ?, RESPONSAVEL = ?, ACESSOS = ?, TIPO_DE_CLIENTE = ?, CONTATO = ?, EMAIL = ?, PRODUTO = ?, SITUACAO = ?, ATUALIZACAO_APLICADA = ?, SGDB = ? WHERE AG = ?',
                    (nome, nome_pasta, pod, cpf_cnpj, liberacao, ambiente, responsavel, acessos, tipo_cliente, contato, email, produto, situacao, atualizacao, sgdb, ag))

            elif cpf_cnpj_count > 0:
                # CPF/CNPJ já existe, exiba um alerta
                messagebox.showerror("Erro", "CPF/CNPJ já cadastrado na base de dados.")
                return

            else:
                cur.execute(f'INSERT INTO CLIENTES (AG, NOME, NOME_DA_PASTA, POD, CPF_CNPJ, DATA_DE_LIBERACAO, AMBIENTE, RESPONSAVEL, ACESSOS, TIPO_DE_CLIENTE, CONTATO, EMAIL, PRODUTO, SITUACAO, ATUALIZACAO_APLICADA, SGDB) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                    (ag, nome, nome_pasta, pod, cpf_cnpj, liberacao, ambiente, responsavel, acessos, tipo_cliente, contato, email, produto, situacao, atualizacao, sgdb))

            cur.close()
            con.commit()
            con.close()

            entry_ag.delete(0, tk.END)
            entry_nome.delete(0, tk.END)
            entry_nome_pasta.delete(0, tk.END)
            entry_pod.delete(0, tk.END)
            entry_cpf_cnpj.delete(0, tk.END)
            entry_liberacao.delete(0, tk.END)
            entry_ambiente.delete(0, tk.END)
            entry_responsavel.delete(0, tk.END)
            entry_acessos.delete(0, tk.END)
            tipo_cliente_var.set("")
            entry_contato.delete(0, tk.END)
            entry_email.delete(0, tk.END)
            entry_produto.delete(0, tk.END)
            situacao_var.set("")
            atualizacao_var.set("")
            sgdb_var.set("")

            messagebox.showinfo("Sucesso", "Cliente Adicionado/Atualizado com sucesso!")

        else:
            # Exiba uma mensagem de erro se algum campo estiver em branco
            messagebox.showerror("Campos Incompletos!", "Por favor, preencha todos os campos antes de salvar ou Atualizar o cliente.")


    except Exception as e:
        print(f"Erro ao salvar clientes: {e}")
            # Exiba uma mensagem de erro se ocorrer algum problema
        messagebox.showerror("Erro ao Salvar Cliente", f"Ocorreu um erro ao salvar o cliente: {str(e)}")


    finally:
            cur.close()
            con.close()


def editar_cliente(lista_clientes, entry_ag, entry_nome, entry_pasta, entry_pod, entry_cpf_cnpj, entry_liberacao, entry_ambiente, entry_responsavel, entry_acessos, tipo_cliente_var, entry_contato, entry_email, entry_produto, situacao_var, atualizacao_var, sgdb_var):
    try:


        selected_item = lista_clientes.curselection()

        print("selected_item:", selected_item)

        if not selected_item or len(selected_item) == 0:
            return

        selected_item_text = lista_clientes.get(selected_item[0])
        ag = selected_item_text.split(',')[0][4:]

        con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
        cur = con.cursor()

        cur.execute(f'SELECT NOME, NOME_DA_PASTA, POD, CPF_CNPJ, DATA_DE_LIBERACAO, AMBIENTE, RESPONSAVEL, ACESSOS, TIPO_DE_CLIENTE, contato, EMAIL, PRODUTO, SITUACAO, ATUALIZACAO_APLICADA, SGDB FROM CLIENTES WHERE AG = ?', (ag,))
        row = cur.fetchone()


        entry_ag.delete(0, tk.END)
        entry_ag.insert(0, ag)
        entry_nome.delete(0, tk.END)
        entry_nome.insert(0, row[0] if row[0] else "")
        entry_pasta.delete(0, tk.END)
        entry_pasta.insert(0, row[1] if row[1] else "")
        entry_pod.delete(0, tk.END)
        entry_pod.insert(0, row[2] if row[2] else "")
        entry_cpf_cnpj.delete(0, tk.END)
        entry_cpf_cnpj.insert(0, row[3] if row[3] else "")
        entry_liberacao.delete(0, tk.END)
        entry_liberacao.insert(0, row[4] if row[4] else "")
        entry_ambiente.delete(0, tk.END)
        entry_ambiente.insert(0, row[5] if row[5] else "")
        entry_responsavel.delete(0, tk.END)
        entry_responsavel.insert(0, row[6] if row[6] else "")
        entry_acessos.delete(0, tk.END)
        entry_acessos.insert(0, row[7] if row[7] else "")
        tipo_cliente_var.set(row[8])
        entry_contato.delete(0, tk.END)
        entry_contato.insert(0, row[9] if row[9] else "")
        entry_email.delete(0, tk.END)
        entry_email.insert(0, row[10] if row[10] else "")
        entry_produto.delete(0, tk.END)
        entry_produto.insert(0, row[11] if row[11] else "")
        situacao_var.set(row[12])
        atualizacao_var.set(row[13])
        sgdb_var.set(row[14])

        # Verificar valores das entradas

        ag = entry_ag.get()
        nome = entry_nome.get()
        nome_pasta = entry_pasta.get()
        pod = entry_pod.get()
        cpf_cnpj = entry_cpf_cnpj.get()
        liberacao = entry_liberacao.get()
        liberacao = formatar_data(liberacao)
        ambiente = entry_ambiente.get()
        responsavel = entry_responsavel.get()
        acessos = entry_acessos.get()
        tipo_cliente = tipo_cliente_var.get()
        contato = entry_contato.get()
        email = entry_email.get()
        produto = entry_produto.get()
        situacao = situacao_var.get()
        atualizacao = atualizacao_var.get()
        sgdb = sgdb_var.get()



        # Atualizar registro
        cur.execute(f'UPDATE CLIENTES SET NOME = ?, NOME_DA_PASTA = ?, POD = ?, CPF_CNPJ = ?, DATA_DE_LIBERACAO = ?, AMBIENTE = ?, RESPONSAVEL = ?, ACESSOS = ?, TIPO_DE_CLIENTE = ?, CONTATO = ?, EMAIL = ?, PRODUTO = ?, SITUACAO = ?, ATUALIZACAO_APLICADA = ?, SGDB = ? WHERE AG = ?',
                    (nome, nome_pasta, pod, cpf_cnpj, liberacao, ambiente, responsavel, acessos, tipo_cliente, contato, email, produto, situacao, atualizacao, sgdb, ag))


        print(f'Valores passados para cur.execute: {entry_ag=}, {entry_nome=}, {entry_pasta=}, {entry_pod=},{cpf_cnpj=},{entry_liberacao=},{entry_ambiente=},{entry_responsavel=},{entry_acessos=},{tipo_cliente_var=},{entry_contato=},{entry_email=},{entry_produto=},{situacao_var=},{atualizacao_var=},{sgdb_var=}')
        cur.execute(f'UPDATE CLIENTES SET NOME = ?, NOME_DA_PASTA = ?, POD = ?, DATA_DE_LIBERACAO = ?, AMBIENTE = ?, RESPONSAVEL = ?, ACESSOS = ?, TIPO_DE_CLIENTE = ?, CONTATO = ?, EMAIL = ?, PRODUTO = ?, SITUACAO = ?, ATUALIZACAO_APLICADA = ?, SGDB = ? WHERE AG = ?',
                    (nome, nome_pasta, pod, liberacao, ambiente, responsavel, acessos, tipo_cliente, contato, email, produto, situacao, atualizacao, sgdb, ag))


        con.commit()

        cur.close()
        con.close()


    except Exception as e:
        print(f"Erro ao editar: {e}")


def excluir_cliente(lista_clientes, entry_ag, entry_nome, entry_pasta, entry_pod, entry_cpf_cnpj, entry_liberacao, entry_ambiente, entry_responsavel, entry_acessos, tipo_cliente_var, entry_contato, entry_email, entry_produto, situacao_var, atualizacao_var, sgdb_var):
    try:
        selected_item = lista_clientes.curselection()
        if not selected_item or len(selected_item) == 0:
            return

        selected_item_text = lista_clientes.get(selected_item[0])
        ag = selected_item_text.split(',')[0][4:]

        resultado = messagebox.askquestion("Confirmar Exclusão", "Tem certeza que deseja excluir este cliente?")

        if resultado == 'yes':
            con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
            cur = con.cursor()

            cur.execute("DELETE FROM clientes WHERE AG=?", (ag,))
            con.commit()

            cur.close()
            con.close()

            entry_ag.delete(0, tk.END)
            entry_nome.delete(0, tk.END)
            entry_pasta.delete(0, tk.END)
            entry_pod.delete(0, tk.END)
            entry_cpf_cnpj.delete(0, tk.END)
            entry_liberacao.delete(0, tk.END)
            entry_ambiente.delete(0, tk.END)
            entry_responsavel.delete(0, tk.END)
            entry_acessos.delete(0, tk.END)
            tipo_cliente_var.set("")
            entry_contato.delete(0, tk.END)
            entry_email.delete(0, tk.END)
            entry_produto.delete(0, tk.END)
            situacao_var.set("")
            atualizacao_var.set("")
            sgdb_var.set("")

            # Exibe uma mensagem de sucesso
            messagebox.showinfo("Sucesso", "Cliente excluído com sucesso!")

            # Atualize a treeview para refletir a exclusão
            atualizar_lista_clientes(lista_clientes)

    except fdb.Error as e:
        messagebox.showerror("Erro", "Erro ao excluir o cliente: " + str(e))

def atualizar_lista_clientes(lista_clientes):
    lista_clientes.delete(0, tk.END)

def formatar_data(data):
    # usar biblioteca tkcalendar para formatar a data
    # Por exemplo, se a entrada do usuário for no formato 'dd/mm/yyyy'
    # pode converter para 'yyyy-mm-dd' para armazenar no banco de dados
    return data.replace('/', '-')


def validar_cpf(cpf):
    return CPF().validate(cpf)

def validar_cnpj(cnpj):
    return CNPJ().validate(cnpj)


def buscar_por_ag(ag):
    try:
        con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
        cur = con.cursor()

        query = 'SELECT AG, NOME, CPF_CNPJ FROM CLIENTES WHERE AG = ?'
        print("Consulta SQL:", query)

        cur.execute(query, (ag,))
        rows = cur.fetchall()

        lista_clientes.delete(0, tk.END)
        if rows:
            for row in rows:
                lista_clientes.insert(tk.END, f'AG: {row[0]}, CLIENTE: {row[1]}, CPF/CNPJ: {row[2]}')
        else:
            lista_clientes.insert(tk.END, 'Nenhum cliente encontrado com o critério de busca.')

    except Exception as e:
        print(f"Erro ao buscar clientes: {e}")

    finally:
        if cur:
            cur.close()
        if con:
            con.close()


# Função para formatar data enquanto o usuário digita
def formatar_data_em_tempo_real(event):
    entrada = event.widget.get()
    entrada = re.sub(r'\D', '', entrada)  # Remover caracteres não numéricos

    if len(entrada) >= 8:
        # Formatar como YYYY-MM-DD
        entrada = f"{entrada[:4]}-{entrada[4:6]}-{entrada[6:8]}"
    elif len(entrada) >= 6:
        # Formatar como YYYY-MM-
        entrada = f"{entrada[:4]}-{entrada[4:6]}-{entrada[6:8]}"

    event.widget.delete(0, tk.END)
    event.widget.insert(0, entrada)


def buscar_por_cpf(cpf):
    cpf = cpf.replace(".", "").replace("-", "")
    try:
        con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
        cur = con.cursor()

        query = 'SELECT AG, NOME, CPF_CNPJ FROM CLIENTES WHERE CPF_CNPJ = ?'
        print("Consulta SQL:", query)

        cur.execute(query, (cpf,))
        rows = cur.fetchall()

        lista_clientes.delete(0, tk.END)
        for row in rows:
            lista_clientes.insert(tk.END, f'AG: {row[0]}, Nome: {row[1]}, CPF/CNPJ: {row[2]}')

        if not rows:
            messagebox.showinfo("Resultado da busca", "Nenhum cliente encontrado com o CPF informado.")
    except fdb.Error as e:
        print(f"Erro na busca por CPF: {e}")
    finally:
        if cur:
            cur.close()
        if con:
            con.close()


def buscar_por_cnpj(cnpj):
    cnpj = cnpj.replace(".", "").replace("/", "").replace("-", "")
    try:
        con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
        cur = con.cursor()

        cur.execute("SELECT AG, NOME, CPF_CNPJ FROM CLIENTES WHERE CPF_CNPJ = ?", (cnpj,))
        rows = cur.fetchall()

        lista_clientes.delete(0, tk.END)
        for row in rows:
            lista_clientes.insert(tk.END, f'AG: {row[0]}, Nome: {row[1]}, CPF/CNPJ: {row[2]}')

        if not rows:
            messagebox.showinfo("Resultado da busca", "Nenhum cliente encontrado com o CNPJ informado.")
    except fdb.Error as e:
        print(f"Erro na busca por CNPJ: {e}")
    finally:
        if cur:
            cur.close()
        if con:
            con.close()


def exibir_resultados(clientes):
    lista_clientes.delete(0, tk.END)
    for cliente in clientes:
        lista_clientes.insert(tk.END, f'AG: {cliente[0]}, CLIENTE: {cliente[1]}, CPF/CNPJ: {cliente[2]}')



def buscar_clientes(entry_busca):
    busca = entry_busca.get().strip()

    if len(busca) > 0:
        busca_formatada = busca.replace(".", "").replace("/", "").replace("-", "")

        if busca.isdigit() and (len(busca_formatada) <= 8 or len(busca) <= 8):
            buscar_por_ag(busca)
        elif len(busca_formatada) == 11 and validar_cpf(busca_formatada):
            buscar_por_cpf(busca_formatada)
        elif len(busca_formatada) == 14 and validar_cnpj(busca_formatada):
            buscar_por_cnpj(busca_formatada)
        else:
            messagebox.showerror("Erro de busca", "Digite um AG válido ou um CPF/CNPJ válido.")
    else:
        messagebox.showerror("Erro de busca", "Digite um AG ou CPF/CNPJ para buscar.")



# Função para atualizar os dados no banco de dados Firebird
def atualizar_dados_firebird(progress_var, mensagem_exibida, lista_clientes):

    conn = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB',
        user='SYSDBA',
        password='masterkey')
    cursor = conn.cursor()

    try:
        # Simula a atualização dos dados
        for i in range(10):
            time.sleep(0.5)  # Simula o tempo necessário para atualizar
            progress_var.set((i + 1) * 10)  # Atualiza a barra de progresso

            update_query = f"UPDATE clientes SET nome = 'Novo Nome' WHERE ag = {i}"
            cursor.execute(update_query)
            conn.commit()

        if not mensagem_exibida[0]:
            mensagem_exibida[0] = True
            root.update_idletasks()  # Força a atualização imediata da GUI
            messagebox.showinfo("Atualização Concluída", "A atualização dos dados foi concluída com sucesso!")

    except Exception as e:
        progress_var.set(0)
        messagebox.showerror("Erro na Atualização", f"Ocorreu um erro durante a atualização: {str(e)}")
    finally:
        cursor.close()
        conn.close()
        progress_var.set(100)


# Função para simular a atualização dos dados
def atualizar_progresso(progress_window, progress_bar, progress_var, mensagem_exibida):
    value = progress_var.get()
    progress_bar["value"] = value
    if value < 100:
        root.after(1000, atualizar_progresso, progress_window, progress_bar, progress_var, mensagem_exibida)
    else:
        progress_window.destroy()


# Função para iniciar a atualização em uma thread separada
def iniciar_atualizacao():

    lista_clientes.delete(0, tk.END)
    progress_window = tk.Toplevel(root)
    progress_window.title("Atualização em Progresso")

    progress_var = tk.IntVar()
    mensagem_exibida = [False]

    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=200, mode="determinate", variable=progress_var)
    progress_bar.pack(padx=20, pady=20)

    thread = threading.Thread(target=atualizar_dados_firebird, args=(progress_var, mensagem_exibida, lista_clientes))
    thread.start()

    # Inicializa a função de atualização de progresso
    atualizar_progresso(progress_window, progress_bar, progress_var, mensagem_exibida)


# Função para limpar os campos
def limpar_campos(entry_ag, entry_nome, entry_pasta, entry_pod, entry_cpf_cnpj, entry_liberacao, entry_ambiente, entry_responsavel, entry_acessos, tipo_cliente_var, entry_contato, entry_email, entry_produto, situacao_var, atualizacao_var, sgdb_var):
    entry_ag.delete(0, tk.END)
    entry_nome.delete(0, tk.END)
    entry_pasta.delete(0, tk.END)
    entry_pod.delete(0, tk.END)
    entry_cpf_cnpj.delete(0, tk.END)
    entry_liberacao.delete(0, tk.END)
    entry_ambiente.delete(0, tk.END)
    entry_responsavel.delete(0, tk.END)
    entry_acessos.delete(0, tk.END)
    tipo_cliente_var.set("")
    entry_contato.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    entry_produto.delete(0, tk.END)
    situacao_var.set("")
    atualizacao_var.set("")
    sgdb_var.set("")


def criar_interface():

    global lista_clientes, con, cur, root
    con = fdb.connect(dsn='D:\Banco_de_Dados\CADCLI\BDS\BASE_DE_DADOS.FDB', user='SYSDBA', password='masterkey')
    cur = con.cursor()

    root = tk.Tk()
    root.title("Gerenciador de Clientes")

    #Variaveis combobox

    tipo_cliente_var = tk.StringVar()
    situacao_var = tk.StringVar()
    atualizacao_var = tk.StringVar()
    sgdb_var = tk.StringVar()


    lista_clientes = tk.Listbox(root, height=10, width=75)

    label_ag = tk.Label(root, text="AG:")
    entry_ag = tk.Entry(root, width=10)

    label_nome = tk.Label(root, text="Razão Social:")
    entry_nome = tk.Entry(root, width=30)

    label_pasta = tk.Label(root, text="Nome da Pasta:")
    entry_pasta = tk.Entry(root, width=30)

    label_pod = tk.Label(root, text="POD:")
    entry_pod = tk.Entry(root, width=15)

    label_cpf_cnpj = tk.Label(root, text="CPF / CNPJ:")
    entry_cpf_cnpj = tk.Entry(root, width=15)

    label_liberacao = tk.Label(root, text="Liberado em:")
    entry_liberacao = tk.Entry(root, width=15)

    label_ambiente = tk.Label(root, text="Ambiente:")
    entry_ambiente = tk.Entry(root, width=15)

    label_responsavel = tk.Label(root, text="Responsável:")
    entry_responsavel = tk.Entry(root, width=15)

    label_acessos = tk.Label(root, text="Acessos:")
    entry_acessos = tk.Entry(root, width=10)

    label_tipo_cliente = tk.Label(root, text="Tipo de Cliente:")
    tipo_cliente_var = tk.Entry(root, width=15)

    label_contato = tk.Label(root, text="Contato:")
    entry_contato = tk.Entry(root, width=20)

    label_email = tk.Label(root, text="Email:")
    entry_email = tk.Entry(root, width=30)

    label_produto = tk.Label(root, text="Produto:")
    entry_produto = tk.Entry(root, width=25)

    label_situacao = tk.Label(root, text="Situação:")
    situacao_var = tk.Entry(root, width=10)

    label_atualizacao = tk.Label(root, text="Atualização Aplicada:")
    atualizacao_var = tk.Entry(root, width=10)

    label_sgdb = tk.Label(root, text="SGDB:")
    sgdb_var = tk.Entry(root, width=10)

    lista_clientes = tk.Listbox(root, height=15, width=110)
    entry_ag = tk.Entry(root, width=30)
    entry_nome = tk.Entry(root, width=30)
    entry_cpf_cnpj = tk.Entry(root, width=30)

    # Botão para listar clientes
    btn_listar = ttk.Button(root, text="Listar Clientes", command=listar_clientes)
    btn_listar.grid(row=18, column=1, columnspan=1, padx=10, pady=5)


    #Botaõ para Salvar

    btn_salvar = ttk.Button(root, text="Atualizar/Adicionar Cliente", command=lambda: salvar_cliente(entry_ag,entry_nome,entry_pasta,entry_pod,entry_cpf_cnpj,entry_liberacao,entry_ambiente,entry_responsavel,entry_acessos,tipo_cliente_var,entry_contato,entry_email,entry_produto,situacao_var,atualizacao_var,sgdb_var))
    btn_salvar.grid(row=18, column=0, padx=10, pady=5)

    # Botão par Editar

    btn_editar = ttk.Button(root, text="Editar", command=lambda: editar_cliente(lista_clientes,entry_ag,entry_nome,entry_pasta,entry_pod,entry_cpf_cnpj,entry_liberacao,entry_ambiente,entry_responsavel,entry_acessos,tipo_cliente_var,entry_contato,entry_email,entry_produto,situacao_var,atualizacao_var,sgdb_var))
    btn_editar.grid(row=19, column=2, columnspan=1, padx=10, pady=5)

    label_ag.grid(row=0, column=0, padx=10, pady=5)
    entry_ag.grid(row=0, column=1, padx=10, pady=5)

    label_nome.grid(row=1, column=0, padx=10, pady=5)
    entry_nome.grid(row=1, column=1, padx=10, pady=5)

    label_pasta.grid(row=2, column=0, padx=10, pady=5)
    entry_pasta.grid(row=2, column=1, padx=10, pady=5)

    label_pod.grid(row=3, column=0, padx=10, pady=5)
    entry_pod.grid(row=3, column=1, padx=10, pady=5)

    label_cpf_cnpj.grid(row=4, column=0, padx=10, pady=5)
    entry_cpf_cnpj.grid(row=4, column=1, padx=10, pady=5)

    label_liberacao.grid(row=5, column=0, padx=10, pady=5)
    entry_liberacao.grid(row=5, column=1, padx=10, pady=5)
    entry_liberacao.bind("<KeyRelease>", formatar_data_em_tempo_real)

    label_ambiente.grid(row=6, column=0, padx=10, pady=5)
    entry_ambiente.grid(row=6, column=1, padx=10, pady=5)

    label_responsavel.grid(row=7, column=0, padx=10, pady=5)
    entry_responsavel.grid(row=7, column=1, padx=10, pady=5)

    label_acessos.grid(row=8, column=0, padx=10, pady=5)
    entry_acessos.grid(row=8, column=1, padx=10, pady=5)

    label_tipo_cliente.grid(row=9, column=0, padx=10, pady=5)
    tipo_cliente_var = tk.StringVar()
    tipo_cliente_var.set("")
    tipo_cliente_combobox = ttk.Combobox(root, values=["Novo", "Base", "NovoEsocial"],textvariable=tipo_cliente_var)
    tipo_cliente_combobox.set("")
    tipo_cliente_combobox.grid(row=9, column=1, padx=10, pady=5)

    label_contato.grid(row=10, column=0, padx=10, pady=5)
    entry_contato.grid(row=10, column=1, padx=10, pady=5)

    label_email.grid(row=11, column=0, padx=10, pady=5)
    entry_email.grid(row=11, column=1, padx=10, pady=5)

    label_produto.grid(row=12, column=0, padx=10, pady=5)
    entry_produto.grid(row=12, column=1, padx=10, pady=5)

    label_situacao.grid(row=13, column=0, padx=10, pady=5)
    situacao_var = tk.StringVar()
    situacao_var.set("")
    situacao_combobox = ttk.Combobox(root, values=["Ativo", "Ativando","Desativado"], textvariable=situacao_var)
    situacao_combobox.set("")
    situacao_combobox.grid(row=13, column=1, padx=10, pady=5)

    label_atualizacao.grid(row=14, column=0, padx=10, pady=5)
    atualizacao_var = tk.StringVar()
    atualizacao_var.set("")
    atualizacao_combobox = ttk.Combobox(root, values=["Sim", "Não"], textvariable=atualizacao_var)
    atualizacao_combobox.set("")
    atualizacao_combobox.grid(row=14, column=1, padx=10, pady=5)

    label_sgdb.grid(row=15, column=0, padx=10, pady=5)
    sgdb_var = tk.StringVar()
    sgdb_var.set("")
    sgdb_combobox = ttk.Combobox(root, values=["SQL", "FB"], textvariable=sgdb_var)
    sgdb_combobox.set("")
    sgdb_combobox.grid(row=15, column=1, padx=10, pady=5)


    # Adicionando um buscador

    label_busca = tk.Label(root, text="Buscar por AG ou CPF/CNPJ:")
    entry_busca = tk.Entry(root, width=30)

    btn_buscar = ttk.Button(root, text="Buscar", command=lambda: buscar_clientes(entry_busca))

    label_busca.grid(row=19, column=0, padx=10, pady=5)
    entry_busca.grid(row=19, column=0, columnspan=3, padx=12, pady=5)
    btn_buscar.grid(row=19, column=1, columnspan=4, padx=7, pady=5)

    lista_clientes.grid(row=20, column=0, columnspan=3, padx=10, pady=5)

    # Botão para atualizar
    btn_atualizar = ttk.Button(root, text="Atualizar", command=iniciar_atualizacao)
    btn_atualizar.grid(row=18, column=2, columnspan=1, padx=10, pady=5)


    # Botão para limpar campos
    btn_limpar = ttk.Button(root, text="Limpar", command=lambda: limpar_campos(entry_ag, entry_nome, entry_pasta, entry_pod, entry_cpf_cnpj, entry_liberacao, entry_ambiente, entry_responsavel, entry_acessos, tipo_cliente_var, entry_contato, entry_email, entry_produto, situacao_var, atualizacao_var, sgdb_var))
    btn_limpar.grid(row=18, column=0, columnspan=2, padx=10, pady=5)

    #Botão de Excluir
    btn_excluir = ttk.Button(root, text="Excluir Cliente", command=lambda: excluir_cliente(lista_clientes, entry_ag, entry_nome, entry_pasta, entry_pod, entry_cpf_cnpj, entry_liberacao, entry_ambiente, entry_responsavel, entry_acessos, tipo_cliente_var, entry_contato, entry_email, entry_produto, situacao_var, atualizacao_var, sgdb_var))
    btn_excluir.grid(row=1, column=2, columnspan=3, padx=10, pady=5)


    root.protocol("WM_DELETE_WINDOW", fechar_conexao)


def fechar_conexao():
    global con, cur
    if cur:
        cur.close()
    if con:
        con.close()
    root.destroy()


criar_interface()
root.mainloop()