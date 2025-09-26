import tkinter as tk
from tkinter import messagebox
import random
import string

def generate_password(length):
    characters = string.ascii_letters + string.digits
    password = ''
    while len(password) < length:
        password += random.choice(characters)
    return password

def generate_and_copy_password():
    length = slider.get()  # Obtém o comprimento da senha a partir do controle deslizante
    password = generate_password(length)
    root.clipboard_clear()  # Limpa a área de transferência
    root.clipboard_append(password)  # Copia a senha para a área de transferência
    root.update()  # Atualiza a interface gráfica após a cópia
    messagebox.showinfo("Senha gerada", f"Senha gerada: {password}\nCopiada para a área de transferência")

root = tk.Tk()
root.title("Gerador de Senhas")

# Controle deslizante para escolher o comprimento da senha
slider = tk.Scale(root, from_=4, to=20, orient="horizontal", label="Comprimento da Senha")
slider.pack()

# Botão para gerar e copiar a senha
generate_button = tk.Button(root, text="Gerar e Copiar Senha", command=generate_and_copy_password)
generate_button.pack()

root.mainloop()
