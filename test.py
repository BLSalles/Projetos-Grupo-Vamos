import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import sqlite3
def automate_process():
   login = login_entry.get()
   password = password_entry.get()
   creditor_number = creditor_number_entry.get()
   file_path = file_path_entry.get()
   if not login or not password or not creditor_number or not file_path:
       messagebox.showerror("Error", "Todos os campos devem ser preenchidos.")
       return
   if not creditor_number.isdigit():
       messagebox.showerror("Error", "O número do credor deve conter apenas números.")
       return
   try:
       df = pd.read_excel(file_path)
   except FileNotFoundError:
       messagebox.showerror("Error", "Arquivo não encontrado.")
       return
   driver = webdriver.Edge()
   driver.get("https://www.serasaexperian.com.br/meus-produtos/login")
   driver.maximize_window()
   try:
       login_field = WebDriverWait(driver, 10).until(
           EC.presence_of_element_located((By.XPATH, "//input[@id='loginUser']"))
       )
       password_field = driver.find_element(By.XPATH, "//input[@id='loginPassword']")
       btn_acess = driver.find_element(By.XPATH, "//button[@id='loginFormSubmit']")
       login_field.send_keys(login)
       password_field.send_keys(password)
       btn_acess.click()
       time.sleep(5)
       tour = WebDriverWait(driver, 10).until(
           EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-mdc-dialog-0"]/div/div/div/mat-dialog-content/div/div[2]/div/a'))
       )
       driver.execute_script("arguments[0].scrollIntoView();", tour)
       tour.click()
       rec_div = WebDriverWait(driver, 10).until(
           EC.element_to_be_clickable((By.XPATH, '//*[@id="prod-64adce65a625ed4687bde841"]/div/div[3]/button'))
       )
       rec_div.click()
       incluc_div = WebDriverWait(driver, 10).until(
           EC.element_to_be_clickable((By.XPATH, "//span[@class='sc-e17868fb-2 dlssDw']"))
       )
       incluc_div.click()
       time.sleep(3)
       conn = sqlite3.connect('negativacao.db')
       c = conn.cursor()
       c.execute('''CREATE TABLE IF NOT EXISTS registros
                     (CNPJ TEXT, Razao_Social TEXT, Valor REAL, Titulo TEXT, Status TEXT, Data_Hora TEXT)''')
       def automat_negative(row):
           try:
               credor = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='creditor.documentNumber']"))
               )
               credor.clear()
               credor_number_str = str(creditor_number).zfill(7)
               credor.send_keys(credor_number_str)
               CNPJ = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.documentNumber']"))
               )
               CNPJ.clear()
               cnpj_cpf_str = str(row['CNPJ/CPF'])
               if len(cnpj_cpf_str) > 11 and len(cnpj_cpf_str) <= 13:
                   cnpj_cpf_str = cnpj_cpf_str.zfill(14)
               elif len(cnpj_cpf_str) < 11:
                   cnpj_cpf_str = cnpj_cpf_str.zfill(11)
               CNPJ.send_keys(cnpj_cpf_str)
               Data_oc = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='dueDate']"))
               )
               Data_oc.clear()
               data_vencimento = row['Vencimento'].strftime("%d/%m/%Y")
               Data_oc.send_keys(data_vencimento)
               valor_div = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='value']"))
               )
               valor_div.clear()
               valor_formatado = "{:.2f}".format(row['Valor']).replace('.', ',')
               valor_div.send_keys(valor_formatado)
               natureza = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, '//*[@id="categoryId"]/option[38]')))
               natureza.click()
               titulo = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='contractNumber']"))
               )
               titulo.clear()
               documento = str(int(row['Documento'])).zfill(9)
               if row['Parcela'] and str(row['Parcela']).strip():
                   titulo.send_keys(f"{documento}/{row['Parcela']}")
               else:
                   titulo.send_keys(documento)
               rz_social = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.name']"))
               )
               rz_social.clear()
               rz_social.send_keys(row['Nome cliente'])
               cep = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.address.zipCode']"))
               )
               cep.clear()
               cep.send_keys(row['CEP'])
               time.sleep(3)
               logradouro = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.address.addressLine']"))
               )
               logradouro.clear()
               logradouro.send_keys(row['Endereco'])
               numero = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.address.number']"))
               )
               numero.clear()
               numero.send_keys(str(row['Nro']))
               complem = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.address.complement']"))
               )
               complem.clear()
               complem.send_keys(str(row['Complem']))
               bairro = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, '//*[@id="debtor.address.district"]')))
               bairro.clear()
               bairro.send_keys(row['Bairro'])
               uf = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.address.state']"))
               )
               uf.clear()
               uf.send_keys(row['UF'])
               cidade = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, "//input[@id='debtor.address.city']"))
               )
               cidade.clear()
               cidade.send_keys(row['Cidade'])
               try:
                   btn_enviar_div = WebDriverWait(driver, 10).until(
                       EC.presence_of_element_located((By.XPATH, '//*[@id="__next"]/main/div/form/div[9]/div/button')))
                   btn_enviar_div.click()
                   status = "Enviado"
               except:
                   status = "Não Enviado"
               time.sleep(3)
               btn_modal = WebDriverWait(driver, 10).until(
                   EC.presence_of_element_located((By.XPATH, '//*[@id="modal"]/div[1]/div/div/div[2]/div/div/button[2]')))
               btn_modal.click()
           except Exception as e:
               if driver.find_elements(By.XPATH, '//*[@id="toast-notification-c38fb503-75e0-4a27-a737-22afe3841796"]'):
                   status = "Erro: Notificação de erro detectada"
               else:
                   status = f"Erro: {str(e)}"
           data_hora = time.strftime("%Y-%m-%d %H:%M:%S")
           c.execute("INSERT INTO registros (CNPJ, Razao_Social, Valor, Titulo, Status, Data_Hora) VALUES (?, ?, ?, ?, ?, ?)",
                     (str(row['CNPJ/CPF']), row['Nome cliente'], row['Valor'], f"{documento}/{row['Parcela']}", status, data_hora))
           conn.commit()
       for index, row in df.iterrows():
           if pd.notna(row['CNPJ/CPF']):
               automat_negative(row)
               time.sleep(3)
       driver.close()
       conn.close()
       messagebox.showinfo("Informação", "Processo concluído.")
   except Exception as e:
       driver.quit()
       messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
def browse_file():
   file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xls;.xlsx")])
   file_path_entry.delete(0, tk.END)
   file_path_entry.insert(0, file_path)
def consultar_dados():
   conn = sqlite3.connect('negativacao.db')
   df = pd.read_sql_query("SELECT * FROM registros", conn)
   conn.close()
   for i in tree.get_children():
       tree.delete(i)
   for index, row in df.iterrows():
       tree.insert("", "end", values=(row['CNPJ'], row['Razao_Social'], row['Valor'], row['Titulo'], row['Status'], row['Data_Hora']))
root = tk.Tk()
root.title("Automatização de Negativação")
root.geometry("800x600")
login_label = tk.Label(root, text="Login:")
login_label.pack()
login_entry = tk.Entry(root)
login_entry.pack()
password_label = tk.Label(root, text="Senha:")
password_label.pack()
password_entry = tk.Entry(root, show='*')
password_entry.pack()
creditor_number_label = tk.Label(root, text="Número do Credor:")
creditor_number_label.pack()
creditor_number_entry = tk.Entry(root)
creditor_number_entry.pack()
file_path_label = tk.Label(root, text="Caminho do Arquivo:")
file_path_label.pack()
file_path_entry = tk.Entry(root)
file_path_entry.pack()
browse_button = tk.Button(root, text="Procurar", command=browse_file)
browse_button.pack()
start_button = tk.Button(root, text="Iniciar", command=automate_process)
start_button.pack()
tree = ttk.Treeview(root, columns=("CNPJ", "Razao_Social", "Valor", "Titulo", "Status", "Data_Hora"), show="headings")
tree.pack(fill=tk.BOTH, expand=True)
for col in tree["columns"]:
   tree.heading(col, text=col)
consultar_button = tk.Button(root, text="Consultar Dados", command=consultar_dados)
consultar_button.pack()
root.mainloop()