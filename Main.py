import pandas as pd
import openpyxl as op

from Suporte import *

import tkinter as tk  # Biblioteca para GUI em Python
from tkinter import filedialog, messagebox  # Componentes específicos do tkinter para diálogos e mensagens

global x
x=1

def processo(reset,continuar):
    app.iconify()
    
    if reset == True:
        # Zera o arquivo
        pdVazio = pd.DataFrame({"":[""]})
        pdVazio.to_excel("Tabela-ClientesAnalisados.xlsx",sheet_name="Faltantes", index=False)
        with pd.ExcelWriter("Tabela-ClientesAnalisados.xlsx",mode='a') as writer:
            pdVazio.to_excel(writer, sheet_name="Realizados",index=False)

    
    if continuar == False:
            if CodSAP.get() == "":
                if file_path_entry.get() != "":
                    # Pega os dados dos Clientes da Tabela Excel
                    print(str(file_path_entry.get()))
                    tabela1 = pd.read_excel(str(file_path_entry.get()))
                    print(tabela1)
                    # Filtra as colunas da tabela inserida
                    tabela = tabela1[["Conta","Nome 1","Descrição Setor de Atividade"]]
                    # print (tabela)
                    # Filtra por Ativos a tabela inserida
                    df_mask = tabela["Descrição Setor de Atividade"] == "Ativos"
                    tabela = tabela[df_mask].reset_index(drop=True)
                    # print ("\n **Filto ativos**\n",tabela)
                    # Retira as duplicatas de Cod SAP
                    tabela = tabela.drop_duplicates(subset="Conta", keep="first")
                    tabela = tabela.reset_index(drop=True)
                    print ("\n **Sem duplicata**\n",tabela)
                else:
                    app.state("normal")
            else:
                tabela = pd.DataFrame({"Nome 1":["None"],"Conta":["none"]})
                tabela.at[0,"Conta"] = CodSAP.get()
                print(tabela)
            
    else:
        tabela = pd.read_excel("Tabela-ClientesAnalisados.xlsx",sheet_name="Faltantes")

    # Faz um loop por todos os clientes na base de dados
    if len(tabela.index) <=40:
        for linha in tabela.index:
            cliente = tabela.loc[linha,"Conta"]
            supCliente = True
            supCliente = abrirCliente(cliente)
            print("supCliente")
            if supCliente == True:
                print("PreCOnt")
                # Contando quantas linhas tem na tabela do cliente
                contagem = contador()
                print("PosCOnt")
                i=0
                while i<contagem:
                    print("PreNum")
                    print(i)
                    documento = pegarNumDoc(i)
                    print("PosNum")
                    placa = pegarPlaca(documento)
                    if placa != False:
                        alterarAtribuicao(placa,cliente,i)
                    else:
                        abrirCliente(cliente)
                    i=i+1
                retorno()

        app.state("normal")
        supIndex = tabela.iloc[len(tabela.index):] 
        supSup = tabela.iloc[:len(tabela.index)]

        supFim = pd.read_excel("Tabela-ClientesAnalisados.xlsx",sheet_name="Realizados")
        supFim = pd.concat([supFim,supSup],ignore_index=False)

                        
        print(supIndex,"INdez")

        supIndex.to_excel("Tabela-ClientesAnalisados.xlsx",sheet_name="Faltantes", index=False) #Sobrescreve o arquivo

        with pd.ExcelWriter("Tabela-ClientesAnalisados.xlsx",mode='a') as writer:
            supFim.to_excel(writer, sheet_name="Realizados",index=False) #Adiciona ao arquivo
        abaFinal("Para realizar a execução novamente retorne à tela inicial")

    else:
        linha=0
        while linha <= 40:
            cliente = tabela.loc[linha,"Conta"]
            abrirCliente(cliente)
            # Contando quantas linhas tem na tabela do cliente
            contagem = contador()
            i=0
            while i<contagem:
                print(i)
                documento = pegarNumDoc(i)
                placa = pegarPlaca(documento)
                if placa != False:
                    alterarAtribuicao(placa,cliente,i)
                else:
                    abrirCliente(cliente)
                i=i+1

        supIndex = tabela.iloc[40:] 
        supSup = tabela.iloc[:40]

        supFim = pd.read_excel("Tabela-ClientesAnalisados.xlsx",sheet_name="Realizados")
        supFim = pd.concat([supFim,supSup],ignore_index=False)

                        
        print(supIndex)

        supIndex.to_excel("Tabela-ClientesAnalisados.xlsx",sheet_name="Faltantes", index=False) #Sobrescreve o arquivo

        with pd.ExcelWriter("Tabela-ClientesAnalisados.xlsx",mode='a') as writer:
            supFim.to_excel(writer, sheet_name="Realizados",index=False) #Adiciona ao arquivo
        abaFinal("Para continuar a execução da mesma tabela selecione \"Continuar\"")
        

# Função para navegar e selecionar arquivo Excel na interface
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_path_entry.delete(0, tk.END)  # Limpando o campo de entrada atual
    file_path_entry.insert(0, file_path)  # Inserindo o caminho do arquivo selecionado no campo de entrada

def iniciar():
    processo(True,False)

# Função para fechar as abas da interface
def finalizar():
    app.destroy()
    # appEnd.destroy()

    workbook = op.load_workbook("Tabela-ClientesAnalisados.xlsx")

    # Salvar as alterações
    workbook.save(f"Tabela-ClientesFinalizados.xlsx")


# Função para fechar as abas da interface
def continuar():
    appEnd.destroy()
    processo(False,True)

# Função para redefinir as abas da interface
def voltar():   
    CodSAP.delete(0, "end")
    file_path_entry.delete(0, "end")

    app.state("normal")
    appEnd.destroy()

    ss = op.load_workbook("Tabela-ClientesAnalisados.xlsx")
    ss.save(f"Tabela-ClientesFinalizados.xlsx")
    # file.close(f"Tabela-Finalizada{i}")


def abaFinal(texto):
    global appEnd 
    appEnd = tk.Toplevel()
    appEnd.title("Relacao Placa-Atribuicao")
    appEnd.configure(bg='#ED1B24')  # Configurando cor de fundo da janela principal

    # Botões para selecionar arquivo, executar processo, consultar e exportar dados
    tk.Label(appEnd, text="Processo finalizado").grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky='w')
    tk.Label(appEnd, text=texto).grid(row=1, column=0, columnspan=3, padx=10, pady=15, sticky='w')

    tk.Button(appEnd, text="Voltar para tela inicial", command=voltar).grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky='w')
    tk.Button(appEnd, text="Continuar", command=continuar).grid(row=2, column=1, columnspan=3, padx=140, pady=10, sticky='w')
    tk.Button(appEnd, text="Fechar", command=finalizar).grid(row=2, column=2, padx=210, pady=10, sticky='w')
    appEnd.attributes("-topmost", True)

# Configuração da interface gráfica tkinter
app = tk.Tk()
app.title("Relacao Placa-Atribuicao")
app.configure(bg='#ED1B24')  # Configurando cor de fundo da janela principal
tk.Label(app, text="Arquivo:", bg='#ED1B24', fg='white').grid(row=5, column=0, padx=10, pady=0, sticky='w')

file_path_entry = tk.Entry(app)
file_path_entry.grid(row=6, column=0, padx=11, pady=0, sticky='w')


#Labels e campos de entrada para login, senha, número do credor e arquivo
tk.Label(app, text="CodSAP:", bg='#ED1B24', fg='white').grid(row=1, column=0, padx=10, pady=10, sticky='w')
CodSAP = tk.Entry(app)
CodSAP.grid(row=1, column=0, padx=60, pady=10, sticky='w')

# Botões para selecionar arquivo, executar processo, consultar e exportar dados
tk.Button(app, text="Procurar", command=browse_file).grid(row=6, column=0, padx=140, pady=1, sticky='w')
tk.Button(app, text="Executar", command=iniciar).grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky='w')
app.attributes("-topmost", True)

app.mainloop()




