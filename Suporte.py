import win32com.client
import subprocess
import time
import sys

# Faz o login
def loginSAP():    
    session.findById("wnd[0]").maximize
    #Espera o usuaio inserir os dados de login
    validador = False
    while validador != True:
        try:  
            session.findById("wnd[0]/usr/boxMESSAGE_FRAME")
        except:
            validador=True
            
        else:
            time.sleep(1)
    session.createSession()
    
    


def abrirCliente(CodSAP):

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "fbl5n"
    session.findById("wnd[0]").sendVKey(0)

    timerizacao(session.findById("wnd[0]/usr/radX_AISEL").select()) #Selecionar todas as partidas
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = CodSAP
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "6800"
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/COB" # Seleciona o layout correto
     
    session.findById("wnd[0]/usr/chkX_SHBV").selected = True #Partida normal
    session.findById("wnd[0]/usr/chkX_NORM").selected = True #Partida Razão especial
    session.findById("wnd[0]/usr/chkX_SHBV").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    if session.findById("wnd[0]/sbar/pane[0]").text == ("Nenhuma conta preenche as condições de seleção"):
        return False
    elif session.findById("wnd[0]/sbar/pane[0]").text == ("Nenhuma partida selecionada (ver texto descritivo)"):
        return False
    else:
        timerizacao(session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("BELNR"))

        session.findById("wnd[0]/tbar[1]/btn[41]").press()
        return True
    
def retorno():
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()



def pegarNumDoc(num):
    numDoc = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(num, "BELNR")      
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    return numDoc

def pegarPlaca(numDoc):    
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "vf03"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = numDoc
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[19]").press()
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").selectItem("          1", "&Hierarchy")
    except:
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        return False
    else:
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem("          1", "&Hierarchy")
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell[1]").doubleClickItem("          1", "&Hierarchy")
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]").sendVKey(2) 
        # time.sleep(1)
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\15").select()
        conteudo = str(session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZZKFZKZ").text)
        
        if conteudo == "":
            conteudo = str(session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZZANLHTXT").text)
        print(conteudo)
        # time.sleep(5)

        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        return conteudo

def alterarAtribuicao(conteudo,CodSAP,num):
    session.findById("wnd[0]/tbar[0]/okcd").text = "fbl5n"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = CodSAP
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "6800"
    session.findById("wnd[0]/usr/radX_AISEL").select()
    session.findById("wnd[0]/usr/chkX_SHBV").selected = True
    session.findById("wnd[0]/usr/chkX_NORM").selected = True
    session.findById("wnd[0]/usr/chkX_SHBV").setFocus()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    timerizacao(session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn("BELNR"))
    session.findById("wnd[0]/tbar[1]/btn[41]").press()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell(num, "ZUONR")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleColumn = "DMSHB"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").doubleClickCurrentCell()
    session.findById("wnd[0]/tbar[1]/btn[13]").press()
    # time.sleep(5)
    session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = conteudo
    # time.sleep(5)
    session.findById("wnd[0]/usr/txtBSEG-ZUONR").setFocus()
    session.findById("wnd[0]/usr/txtBSEG-ZUONR").caretPosition = 8
    session.findById("wnd[0]/tbar[0]/btn[11]").press()

def timerizacao(CodValidacao):
    tempInicial = 0 #Valor inicial em segundos
    tempFinal = 20 #Valor Final em segundos
    while tempInicial<=tempFinal:
        try:
            CodValidacao
            #Código que deve ser validado
            #Tentativa de acessar uma entrada SAP enquanto o sistema carrega
        except:
            #Se não conseguir acessar a entrada esperar 0.5 segundos
            time.sleep(0.5)
        else:
            #Se conseguir acessar a entrada sair do loop
            break
        finally:
            #Quantidade de tentativas realizadas
            tempInicial = tempInicial + 1
    if tempInicial == tempFinal:
        print("Erro: Tempo sem interacao excedido")

def contador():
    cont = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").RowCount
    print("Total de linhas:", cont)
    return cont

#Abre o SAP
try:
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(5)
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection("# JSL -  ECC - Produção (ECP)", True)
    time.sleep(3)
    session = connection.Children(0)
    loginSAP()
except:
    print(sys.exc_info()[0])
    session = None
    connection = None
    application = None
    SapGuiAuto = None