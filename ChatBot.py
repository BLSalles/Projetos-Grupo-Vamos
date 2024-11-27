import pywhatkit as kit

from PIL import Image, ImageTk

import pandas as pd

import sqlite3

from datetime import datetime

import time

import pyautogui as p

import tkinter as tk

from tkinter import filedialog, messagebox, ttk

# Configura o banco de dados

def configurar_bd():

    conn = sqlite3.connect('envios.db')

    cursor = conn.cursor()

    cursor.execute('''

    CREATE TABLE IF NOT EXISTS envios (

        id INTEGER PRIMARY KEY AUTOINCREMENT,

        telefone TEXT,

        status TEXT,

        mensagem TEXT,

        data_hora TEXT

    )

    ''')

    conn.commit()

    cursor.execute("PRAGMA table_info(envios);")

    columns = [info[1] for info in cursor.fetchall()]

    if 'data_hora' not in columns:

        cursor.execute('ALTER TABLE envios ADD COLUMN data_hora TEXT;')

    conn.commit()

    conn.close()

# Função para salvar o status do envio

def salvar_status(telefone, status, mensagem):

    conn = sqlite3.connect('envios.db')

    cursor = conn.cursor()

    data_hora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    cursor.execute('''

    INSERT INTO envios (telefone, status, mensagem, data_hora)

    VALUES (?, ?, ?, ?)

    ''', (telefone, status, mensagem, data_hora))

    conn.commit()

    conn.close()

# Função para consultar dados do banco de dados com filtros

def consultar_dados():

    conn = sqlite3.connect('envios.db')

    cursor = conn.cursor()

    contato = entrada_contato.get().strip()

    status = entrada_status.get().strip()

    data_inicio = entrada_data_inicio.get().strip()

    data_fim = entrada_data_fim.get().strip()

    query = "SELECT * FROM envios WHERE 1=1"

    if contato:

        query += f" AND telefone LIKE '%{contato}%'"

    if status:

        query += f" AND status LIKE '%{status}%'"

    if data_inicio:

        query += f" AND data_hora >= '{data_inicio} 00:00:00'"

    if data_fim:

        query += f" AND data_hora <= '{data_fim} 23:59:59'"

    cursor.execute(query)

    rows = cursor.fetchall()

    conn.close()

    for item in tree.get_children():

        tree.delete(item)

    for row in rows:

        tree.insert("", "end", text=row[0], values=(row[1], row[2], row[3], row[4]))

# Função para enviar mensagens

def enviar_mensagens(file_path):

    df = pd.read_excel(file_path)

    coluna_contato = 'Contato'

    coluna_mensagens = [f'Mensagem{i}' for i in range(1, 8)]

    codigo_pais = '+55'

    for index, row in df.iterrows():

        telefone_destinatario = int(row[coluna_contato])

        telefone_destinatario = str(row[coluna_contato]).strip()

        if not telefone_destinatario.startswith('+'):

            telefone_destinatario = codigo_pais + telefone_destinatario

        mensagens = '\n'.join([str(row[col]).strip() if not pd.isna(row[col]) else '' for col in coluna_mensagens])

        if pd.isna(telefone_destinatario) or telefone_destinatario.strip() == '':

            messagebox.showerror(title=None, message="Telefone vazio encontrado, parando o envio.")

            break

        try:

            kit.sendwhatmsg_instantly(telefone_destinatario, mensagens)

            time.sleep(2)

            screenshot = p.screenshot()

            screenshot.save("img/screenshot.png")

            error_location = p.locateOnScreen("img/ERROR.png", confidence=0.8)

            if error_location:

                error_message = "O número de telefone compartilhado por URL é inválido."

                salvar_status(telefone_destinatario, error_message, "Mensagem não enviada")

                print("Mensagem não enviada", telefone_destinatario)

                p.hotkey('ctrl', 'shift', 'tab')

                p.hotkey('ctrl', 'w')

        except Exception as e:

            salvar_status(telefone_destinatario, f"Enviado {str(e)}", mensagens)

            print("Mensagem enviada", telefone_destinatario)

            p.hotkey('ctrl', 'shift', 'tab')

            p.hotkey('ctrl', 'w')

    messagebox.showinfo(title='ChatBot', message='Mensagens enviadas com sucesso!')

def selecionar_arquivo():

    file_path = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])

    entrada_var.set(file_path)

# Função para executar o envio de mensagens

def executar():

    file_path = entrada_var.get()

    if file_path:

        enviar_mensagens(file_path)

    else:

        messagebox.showerror(title=None, message="Nenhum arquivo selecionado.")

def exportar_dados():

    conn = sqlite3.connect('envios.db')

    df = pd.read_sql_query("SELECT * FROM envios", conn)

    conn.close()

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if file_path:

        df.to_excel(file_path, index=False)

        messagebox.showinfo("Informação", "Dados exportados com sucesso.")

# Função para iniciar a interface gráfica

def iniciar_interface():

    global entrada_var, entrada_contato, entrada_status, entrada_data_inicio, entrada_data_fim, tree

    root = tk.Tk()

    root.title("Processador de Arquivos Excel")

    root.configure(bg='#ED1B24')

    root.geometry("800x600")

    logo_path = "Logo Vm 1.png"

    logo_image = Image.open(logo_path)

    logo_image = logo_image.resize((100, 50), Image.Resampling.LANCZOS)

    logo_photo = ImageTk.PhotoImage(logo_image)

    logo_label = tk.Label(root, image=logo_photo, bg='#ED1B24')

    logo_label.grid(row=0, column=0, padx=10, pady=10, columnspan=3, sticky='w')

    entrada_var = tk.StringVar()

    tk.Label(root, bg='#ED1B24', fg='white', text="Selecione o arquivo de entrada:").grid(row=1, column=0, padx=10, pady=5, sticky='w')

    file_frame = tk.Frame(root, bg='#ED1B24')

    file_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky='w')

    tk.Entry(file_frame, textvariable=entrada_var, width=50).pack(side='left', padx=5)

    tk.Button(file_frame, text="Procurar", command=selecionar_arquivo).pack(side='left', padx=5)

    tk.Button(root, text="Executar", command=executar).grid(row=3, column=0, padx=10, pady=5, sticky='w')

    tk.Label(root, bg='#ED1B24', fg='white', text="Filtro por Contato:").grid(row=4, column=0, padx=10, pady=10, sticky='w')

    entrada_contato = tk.Entry(root, width=30)

    entrada_contato.grid(row=4, column=1, padx=10, pady=5, sticky='w')

    tk.Label(root, bg='#ED1B24', fg='white', text="Filtro por Status:").grid(row=5, column=0, padx=10, pady=5, sticky='w')

    entrada_status = tk.Entry(root, width=30)

    entrada_status.grid(row=5, column=1, padx=10, pady=5, sticky='w')

    tk.Label(root, bg='#ED1B24', fg='white', text="Filtro por Data Início (YYYY-MM-DD):").grid(row=6, column=0, padx=10, pady=5, sticky='w')

    entrada_data_inicio = tk.Entry(root, width=30)

    entrada_data_inicio.grid(row=6, column=1, padx=10, pady=5, sticky='w')

    tk.Label(root, bg='#ED1B24', fg='white', text="Filtro por Data Fim (YYYY-MM-DD):").grid(row=7, column=0, padx=10, pady=5, sticky='w')

    entrada_data_fim = tk.Entry(root, width=30)

    entrada_data_fim.grid(row=7, column=1, padx=10, pady=5, sticky='w')

    tk.Button(root, text="Consultar", command=consultar_dados).grid(row=8, column=0, columnspan=3, padx=10, pady=5, sticky='w')

    tk.Button(root, text="Exportar", command=exportar_dados).grid(row=8, column=1, padx=10, pady=5, sticky='w')

    tree = ttk.Treeview(root, columns=("Telefone", "Status", "Mensagem", "Data/Hora"), show="headings")

    tree.heading('Telefone', text='Telefone')

    tree.heading('Status', text='Status')

    tree.heading('Mensagem', text='Mensagem')

    tree.heading('Data/Hora', text='Data/Hora')

    tree.grid(row=9, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")

    root.grid_rowconfigure(9, weight=1)

    root.grid_columnconfigure(1, weight=1)

    root.mainloop()

if __name__ == "__main__":

    configurar_bd()

    iniciar_interface()
 