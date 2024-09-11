import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk
class Application(tk.Tk):
   def __init__(self):
       super().__init__()
       self.title("Contract Processor")
       self.configure(bg='#ED1B24')
       # Carregar e exibir logotipo
       logo_path = "Logo Vm 1.png"
       logo_image = Image.open(logo_path)
       logo_image = logo_image.resize((100, 50), Image.Resampling.LANCZOS)
       logo_photo = ImageTk.PhotoImage(logo_image)
       logo_label = tk.Label(self, image=logo_photo, bg='#ED1B24')
       logo_label.grid(row=0, column=0, padx=10, pady=10, columnspan=3, sticky='w')
       self.excel_path = ""
       self.contracts_folder = ""
       ttk.Button(self, text="Selecionar Arquivo Excel", command=self.select_excel_file).grid(row=2, column=0, padx=10, pady=10)
       self.excel_file_label = tk.Label(self, text="Nenhum arquivo selecionado", bg='#ED1B24', fg='white')
       self.excel_file_label.grid(row=2, column=1, padx=10, pady=10)
       ttk.Button(self, text="Selecionar Pasta de Contratos", command=self.select_contract_folder).grid(row=3, column=0, padx=10, pady=10)
       self.folder_label = tk.Label(self, text="Nenhuma pasta selecionada", bg='#ED1B24', fg='white')
       self.folder_label.grid(row=3, column=1, padx=10, pady=10)
       ttk.Button(self, text="Executar", command=self.run_processing).grid(row=4, column=0, padx=10, pady=10)
       ttk.Button(self, text="Exportar para Excel", command=self.export_data).grid(row=4, column=1, padx=10, pady=10)
   def select_excel_file(self):
       self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
       self.excel_file_label.config(text=f"Excel File: {self.excel_path}")
   def select_contract_folder(self):
       self.contracts_folder = filedialog.askdirectory()
       self.folder_label.config(text=f"Contracts Folder: {self.contracts_folder}")
   def run_processing(self):
       from controller import process_contracts
       try:
           process_contracts(self.excel_path, self.contracts_folder)
           messagebox.showinfo("Success", "Processamento concluído!")
       except Exception as e:
           messagebox.showerror("Error", f"Error: {str(e)}")
   def export_data(self):
       from controller import export_to_excel
       try:
           export_to_excel()
           messagebox.showinfo("Success", "Dados exportados para 'status.xlsx'!")
       except Exception as e:
           messagebox.showerror("Error", f"Error: {str(e)}")