import os
import mysql.connector
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from PIL import Image, ImageTk

# Variáveis globais também para os botões "eye_open" e "eye_closed"
show_password_button_open = None
show_password_button_closed = None
edit_password_id = None
show_passwords = None
show_passwords_senha = False
show_passwords = False
add_mode = True
password_entries = []

def generate_password():
    import random
    import string
    length = 13
    characters = string.ascii_letters + string.digits + string.punctuation
    password = ''.join(random.choice(characters) for _ in range(length))
    return password

def generate_and_display_password():
    generated_password = generate_password()
    senha_entry.delete(0, tk.END)
    senha_entry.insert(0, generated_password)


# Defina uma variável para controlar o modo de operação (adicionar ou atualizar senha)
add_mode = True

def add_password():
    onde = onde_entry.get()
    tipo = tipo_combobox.get()
    id_login = id_login_entry.get()
    senha = senha_entry.get()

    if not senha:
        messagebox.showerror("Erro", "A senha não pode estar em branco.")
        return

    try:
        conn = mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='',  # Insira sua senha do MySQL, ou deixe em branco se não houver senha
            database='Armazem'
        )
        cursor = conn.cursor()

        insert_query = "INSERT INTO passwords (Onde, Tipo, IDdeLogin, senha) VALUES (%s, %s, %s, %s)"
        data = (onde, tipo, id_login, senha)  # Salvar a senha real, não o hash
        cursor.execute(insert_query, data)

        conn.commit()
        conn.close()

        onde_entry.delete(0, tk.END)
        tipo_combobox.set("")
        id_login_entry.delete(0, tk.END)
        senha_entry.delete(0, tk.END)
        refresh_password_list()

    except mysql.connector.Error as err:
        messagebox.showerror("Erro no MySQL", f"Erro ao salvar a senha: {err}")



def save_or_update_password():
    global add_mode
    if add_mode:
        add_password()
    else:
        update_password()
    add_mode = True



def search_password():
    search_term = search_entry.get()

    try:
        conn = mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='',  # Insira sua senha do MySQL, ou deixe em branco se não houver senha
            database='Armazem'
        )
        cursor = conn.cursor()

        search_query = "SELECT Onde, Tipo, IDdeLogin, senha FROM passwords WHERE Onde LIKE %s"
        data = ('%' + search_term + '%',)
        cursor.execute(search_query, data)
        results = cursor.fetchall()

        conn.close()

        if results:
            password_list.delete(0, tk.END)
            password_entries.clear()
            for result in results:
                onde, tipo, id_login, senha = result
                password = senha if show_passwords else "*****"
                password_entries.append(result)

                password_list.insert(tk.END, f"{onde} | {tipo} | {id_login}: {password}")

        else:
            messagebox.showinfo("Lista de Senhas", "Nenhuma senha encontrada.")

    except mysql.connector.Error as err:
        messagebox.showerror("Erro no MySQL", f"Erro ao buscar senhas: {err}")


show_passwords = False  # Variável global para rastrear o estado do botão

def toggle_show_password_senha():
    if senha_entry.cget("show") == "":
        senha_entry.config(show="*")
        show_password_button_senha.config(image=eye_closed_icon)
    else:
        senha_entry.config(show="")
        show_password_button_senha.config(image=eye_open_icon)






def toggle_show_password_listbox():
    global show_passwords_open

    selected_item = password_list.curselection()
    if not selected_item:
        messagebox.showinfo("Mostrar Senha", "Selecione uma senha para mostrar.")
        return

    index = selected_item[0]
    senha_item = password_entries[index]

    show_passwords_open = not show_passwords_open

    # Mantenha o item selecionado
    selected_index = index
    password_list.selection_set(selected_index)




# Função para editar senha
def edit_password():
    global edit_password_id
    global selected_index

    selected_item = password_list.curselection()
    if not selected_item:
        messagebox.showinfo("Editar Senha", "Selecione uma senha para editar.")
        return

    index = selected_item[0]
    senha_item = password_entries[index]

    # Preencha os campos de entrada com os detalhes da senha
    onde_entry.delete(0, tk.END)
    onde_entry.insert(0, senha_item[0])  # Onde
    tipo_combobox.set(senha_item[1])  # Tipo
    id_login_entry.delete(0, tk.END)
    id_login_entry.insert(0, senha_item[2])  # ID/Login

    # Atualize o campo de senha, verificando se a senha está presente
    senha = senha_item[3]
    senha_entry.delete(0, tk.END)
    senha_entry.insert(0, senha)

    # Defina a variável global para o ID da senha em edição
    edit_password_id = senha_item[2]  # Use o índice correto (2) para o ID

    # Mantenha o item selecionado
    selected_index = index

    # Atualize o estado do botão "Salvar" para a função de atualização
    save_button.config(text="Atualizar Senha", command=update_password)
    
    # Mantenha o item selecionado
    password_list.selection_set(selected_index)





# Função para atualizar senha
def update_password():
    global selected_index

    selected_item = password_list.curselection()
    if not selected_item:
        messagebox.showinfo("Editar Senha", "Selecione uma senha para editar.")
        return

    index = selected_item[0]
    senha_item = password_entries[index]

    onde = onde_entry.get()
    tipo = tipo_combobox.get()
    id_login = id_login_entry.get()
    senha = senha_entry.get()

    if not senha:
        messagebox.showerror("Erro", "A senha não pode estar em branco.")
        return

    try:
        conn = mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='',
            database='Armazem'
        )
        cursor = conn.cursor()

        # Use o ID da senha selecionada para a atualização
        update_query = "UPDATE passwords SET Onde = %s, Tipo = %s, IDdeLogin = %s, senha = %s WHERE IDdeLogin = %s"
        data = (onde, tipo, id_login, senha, senha_item[2])  # Use o ID de login da senha selecionada
        cursor.execute(update_query, data)

        conn.commit()
        conn.close()

        refresh_password_list()  # Atualize a lista de senhas

        onde_entry.delete(0, tk.END)
        tipo_combobox.set("")
        id_login_entry.delete(0, tk.END)
        senha_entry.delete(0, tk.END)

        save_button.config(text="Salvar Senha", command=save_or_update_password)  # Restaure o botão "Salvar"

        messagebox.showinfo("Editar Senha", "Senha atualizada com sucesso.")

    except mysql.connector.Error as err:
        print(f"Erro no MySQL: {err}")
        messagebox.showerror("Erro no MySQL", f"Erro ao atualizar a senha: {err}")

















def delete_password():
    selected_item = password_list.curselection()
    if not selected_item:
        messagebox.showinfo("Excluir Senha", "Selecione uma senha para excluir.")
        return

    index = selected_item[0]
    senha_item = password_entries[index]
    id_login = senha_item[2]  # Obtenha o ID de login da senha selecionada

    confirm = messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir esta senha?")
    if confirm:
        try:
            conn = mysql.connector.connect(
                host='localhost',
                port=3306,
                user='root',
                password='',
                database='Armazem'
            )
            cursor = conn.cursor()

            delete_query = "DELETE FROM passwords WHERE IDdeLogin = %s"
            data = (id_login,)
            cursor.execute(delete_query, data)

            conn.commit()
            conn.close()

            # Remova a senha da lista de senhas
            password_entries.pop(index)

            password_list.delete(index)
            messagebox.showinfo("Excluir Senha", "Senha excluída com sucesso.")

        except mysql.connector.Error as err:
            messagebox.showerror("Erro no MySQL", f"Erro ao excluir a senha: {err}")


# Variáveis booleanas para rastrear o estado dos botões "eye_open" e "eye_closed"
show_passwords_open = False
show_passwords_closed = True

# Função para alternar a visibilidade da senha selecionada na Listbox
selected_index = None  # Variável para armazenar o índice selecionado

# Variável global para armazenar o índice selecionado
selected_index = None

# Função para alternar a visibilidade da senha selecionada na Listbox
def toggle_show_password_selected():
    global show_passwords_open
    global selected_index

    selected_item = password_list.curselection()
    if not selected_item:
        messagebox.showinfo("Mostrar Senha", "Selecione uma senha para mostrar.")
        return

    index = selected_item[0]
    senha_item = password_entries[index]

    if show_passwords_open:
        senha = senha_item[3]
        show_password_button_list.config(image=eye_open_icon)
    else:
        senha = "*" * len(senha_item[3])
        show_password_button_list.config(image=eye_closed_icon)

    password_list.delete(index)
    password_list.insert(index, f"{senha_item[0]} | {senha_item[1]} | {senha_item[2]}: {senha}")

    show_passwords_open = not show_passwords_open

    # Mantenha o item selecionado
    selected_index = index
    password_list.selection_set(selected_index)


# Função para atualizar o banco de dados antes de encerrar a aplicação
def update_database_on_exit():
    global add_mode

    # Verifique se há alterações pendentes
    if not add_mode:
        messagebox.showinfo("Aviso", "Há uma edição pendente. Por favor, termine a edição antes de sair.")
        return

    # Se não houver alterações pendentes, você pode fechar a aplicação diretamente
    root.destroy()


# Crie uma janela principal
root = tk.Tk()
root.title("Armazenador de Senhas")

# Crie os widgets da interface
frame = ttk.Frame(root, padding=10)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Obtém o diretório do script atual
script_dir = os.path.dirname(__file__)

# Constrói o caminho completo para os ícones
eye_open_icon = Image.open(os.path.join(script_dir, "eye_open.png"))
eye_open_icon = eye_open_icon.resize((20, 20), Image.LANCZOS)
eye_open_icon = ImageTk.PhotoImage(eye_open_icon)

eye_closed_icon = Image.open(os.path.join(script_dir, "eye_closed.png"))
eye_closed_icon = eye_closed_icon.resize((20, 20), Image.LANCZOS)
eye_closed_icon = ImageTk.PhotoImage(eye_closed_icon)



onde_label = ttk.Label(frame, text="Onde:")
onde_label.grid(row=0, column=0, sticky=tk.W)

onde_entry = ttk.Entry(frame, width=40)
onde_entry.grid(row=0, column=1, columnspan=2, sticky=(tk.W, tk.E))

tipo_label = ttk.Label(frame, text="Tipo:")
tipo_label.grid(row=1, column=0, sticky=tk.W)

tipo_combobox = ttk.Combobox(frame, values=["App", "site","Email", "Rede Social", "Site/App","Banco de Dados", "Outro"], state="")
tipo_combobox.grid(row=1, column=1, sticky=(tk.W, tk.E))

id_login_label = ttk.Label(frame, text="ID/Login:")
id_login_label.grid(row=2, column=0, sticky=tk.W)

id_login_entry = ttk.Entry(frame, width=40)
id_login_entry.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E))

senha_label = ttk.Label(frame, text="Senha:")
senha_label.grid(row=3, column=0, sticky=(tk.W, tk.N))

senha_entry = ttk.Entry(frame, width=40, show="*")
senha_entry.grid(row=3, column=1, sticky=(tk.W, tk.E))

save_button = ttk.Button(frame, text="Salvar Senha", command=save_or_update_password)
save_button.grid(row=4, column=0, columnspan=3)
save_button.config(text="Salvar Senha")

search_label = ttk.Label(frame, text="Pesquisar:")
search_label.grid(row=5, column=0, sticky=tk.W)

search_entry = ttk.Entry(frame, width=40)
search_entry.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E))

search_button = ttk.Button(frame, text="Pesquisar", command=search_password)
search_button.grid(row=6, column=0, columnspan=3)

password_list = tk.Listbox(frame, width=60, height=10)
password_list.grid(row=7, column=0, columnspan=3, rowspan=3, sticky=(tk.W, tk.E))

scrollbar = ttk.Scrollbar(frame, orient="vertical", command=password_list.yview)
scrollbar.grid(row=7, column=3, rowspan=3, sticky=(tk.N, tk.S))
password_list.configure(yscrollcommand=scrollbar.set)

show_password_button_senha = ttk.Button(frame, image=eye_closed_icon, command=toggle_show_password_senha)
show_password_button_senha.grid(row=3, column=2, sticky=tk.W)

senha_entry = ttk.Entry(frame, width=40, show="*")  # Campo de entrada de senha inicialmente oculto
senha_entry.grid(row=3, column=1, sticky=(tk.W, tk.E))

edit_button = ttk.Button(frame, text="Editar Senha", command=edit_password)
edit_button.grid(row=10, column=0, columnspan=2, sticky=tk.W)

delete_button = ttk.Button(frame, text="Excluir Senha", command=delete_password)
delete_button.grid(row=10, column=2, sticky=(tk.W, tk.E))

generate_password_button = ttk.Button(frame, text="Gerar Senha", command=generate_and_display_password)
generate_password_button.grid(row=4, column=2, sticky=tk.W)


# Crie um único botão (eye_open ou eye_closed) para alternar a visibilidade da senha selecionada
show_password_button_list = ttk.Button(frame, image=eye_closed_icon, command=toggle_show_password_selected)
show_password_button_list.grid(row=10, column=1)



# Dentro da função que atualiza a lista de senhas (refresh_password_list), adicione:
def refresh_password_list():
    password_list.delete(0, tk.END)  # Limpa a Listbox
    password_entries.clear()
    search_password()


    # Se um item estava selecionado antes de atualizar, selecione-o novamente
    global selected_index
    if selected_index is not None:
        password_list.selection_set(selected_index)
        password_list.see(selected_index)


# Atualizar a lista de senhas ao iniciar o aplicativo
refresh_password_list()

# Configure a função update_database_on_exit para ser chamada quando a janela for fechada
root.protocol("WM_DELETE_WINDOW", update_database_on_exit)

root.mainloop()
