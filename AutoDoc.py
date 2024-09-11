import tkinter as tk
from tkinter import messagebox
import win32com.client
def executar_processo():
   
   cota = str(entry_cota.get())
   lote = int(entry_lote.get())
   # Chamar a função para executar as macros com os valores inseridos
   automatizar_processo1(cota, lote)
   automatizar_processo2()
   tk.messagebox.showinfo(title=None, message="Progresso concluído")
   root.destroy()
   
def automatizar_processo1(cota, lote):
   

   caminho_arquivo = r'C:\Users\thiago.slima\Desktop\AutoDoc.v1\Automat_OMAR.xlsm'

   nome_macro = 'Módulo1.AutomatizarProcesso'

   excel = win32com.client.Dispatch('Excel.Application')
   excel.Visible = True
   
   workbook = excel.Workbooks.Open(Filename=caminho_arquivo)
   sheet = workbook.Sheets('Form_Omar')
 
   sheet.Range("A1").Value = cota
   sheet.Range("B1").Value = lote
  
   excel.Application.Run(f'{workbook.Name}!{nome_macro}')
def automatizar_processo2():
  
   caminho_arquivo = r'C:\Users\thiago.slima\Desktop\AutoDoc.v1\Automat_OMAR.xlsm'
   
   nome_macro = 'Planilha1.PreencherEEnviarEmail'
   
   excel = win32com.client.Dispatch('Excel.Application')
   excel.Visible = True
  
   workbook = excel.Workbooks.Open(Filename=caminho_arquivo)
  
   excel.Application.Run(f'{workbook.Name}!{nome_macro}')

root = tk.Tk()
root.title("Formulário de Entrada")
root.geometry('300x100')

label_cota = tk.Label(root, text="Cota:")
label_cota.grid(row=0, column=0, padx=10, pady=5)
entry_cota = tk.Entry(root,width=30)
entry_cota.grid(row=0, column=1, padx=10, pady=5)

label_lote = tk.Label(root, text="Lote:")
label_lote.grid(row=1, column=0, padx=10, pady=5)
entry_lote = tk.Entry(root,width=30)
entry_lote.grid(row=1, column=1, padx=10, pady=5)

button_executar = tk.Button(root, text="Executar Processo", command=executar_processo)
button_executar.grid(row=2, column=0, columnspan=2, padx=10, pady=5)
root.mainloop()