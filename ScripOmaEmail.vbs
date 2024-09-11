Sub PreencherEEnviarEmail()

   Dim wdApp As Object
   
   Dim wdDoc As Object
   
   Dim ws As Worksheet
   
   Dim caminhoDocumento As String
   
   Dim outlookApp As Object
   
   Dim outlookMail As Object
   
   Dim celula As Range
   
   Dim ultimaLinha As Long
   
   Dim valorG As String
   
   
   ' Defina o caminho do arquivo do Word
   caminhoDocumento = "C:\Users\thiago.slima\Desktop\AutoDoc.v1\Termo de Autorização de Faturamento_Omar.docx" ' Substitua pelo caminho do seu documento
   
   ' Inicialize a planilha
   Set ws = ThisWorkbook.Sheets("Form_Omar")
   
   ' Inicialize o Word
   Set wdApp = CreateObject("Word.Application")
   wdApp.Visible = True ' Deixe visível para ver o que está acontecendo
   
   ' Abra o documento do Word
   Set wdDoc = wdApp.Documents.Open(caminhoDocumento)
 
   ' Percorra as células da coluna A enquanto houver dados
   For Each celula In ws.Range("A3:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
   
            ' Verificar se a coluna G está marcada como "concluído"
            If celula.Offset(0, 6).Value = "CONCLUÍDO" Then
      
                valor = celula.Offset(0, 23).Value ' Coluna X
                
                dt_valor = celula.Value ' Coluna A
                
                nome = celula.Offset(0, 13).Value ' Coluna N
                
                Placa = celula.Offset(0, 5).Value ' Coluna F
                
                modelo = celula.Offset(0, 7).Value ' Coluna H
                
                ano_mod = celula.Offset(0, 9).Value ' Coluna J
                
                chassi = celula.Offset(0, 8).Value ' Coluna I
                
                endereco = celula.Offset(0, 17).Value ' Coluna R
                
                bairro = celula.Offset(0, 18).Value ' Coluna S
                
                estado = celula.Offset(0, 21).Value ' Coluna V
                
                Cpf_cnpj = celula.Offset(0, 14).Value ' Coluna O
                
                mes = celula.Offset(0, 25).Value ' Coluna Z
                
                dia = celula.Offset(0, 24).Value ' Coluna Y
                
                ano = celula.Offset(0, 26).Value ' Coluna AA
                
                email_resp = celula.Offset(0, 27).Value ' Coluna AB
                
       
                ' Preencha o formulário com os dados
                With wdDoc
                
                    .FormFields("Valor").Result = valor
                    
                    .FormFields("dt_valor").Result = dt_valor
                    
                    .FormFields("Nome_cliente").Result = nome
                    
                    .FormFields("Endereco").Result = endereco
                    
                    .FormFields("Bairro").Result = bairro
                    
                    .FormFields("Placa").Result = Placa
                    
                    .FormFields("Modelo").Result = modelo
                    
                    .FormFields("Ano_Mod").Result = ano_mod
                    
                    .FormFields("Chassi").Result = chassi
                    
                    .FormFields("Estado").Result = estado
                    
                    .FormFields("Cpf_cnpj").Result = Cpf_cnpj
                    
                    .FormFields("mes").Result = mes
                    
                    .FormFields("dia").Result = dia
                    
                    .FormFields("ano").Result = ano
                    
                End With
            
        
            
                ' Salve o documento do Word
                wdDoc.Save
                
                ' Configurar o e-mail
                Set outlookApp = CreateObject("Outlook.Application")
                
                Set outlookMail = outlookApp.CreateItem(0)
                
                     
                With outlookMail
                
                    .To = email_resp ' Endereço de e-mail do destinatário
                    
                    .Cc = "thiago.slima@grupovamos.com.br" ' 'Endereço de e-mail do Remetendo para Arquivar o Doc
                    
                    .Subject = "Termo de Autorização de Faturamento Veículo :" & Placa
                    
                    .Body = "Bom dia "
                    
                    .Attachments.Add caminhoDocumento ' Anexar o documento preenchido
                    
                    .Send ' Enviar o e-mail
                    
                End With
                
                email_resp = ""
            
        End If
       
   Next celula
   
   
   ' Feche o documento e o Word
   
   wdDoc.Close
   
   wdApp.Quit
   
   ' Libere a memória
   Set wdDoc = Nothing
   
   Set wdApp = Nothing
   
   Set outlookMail = Nothing
   
   Set outlookApp = Nothing
      
End Sub


