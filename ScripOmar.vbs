Sub AutomatizarProcesso()
 
    Dim wsAutomat As Workbook
 
    Dim wsPgtos As Workbook
 
    Dim wsAutomatSheet As Worksheet
 
    Dim wsPgtosSheet As Worksheet
 
    Dim wsResumoSheet As Worksheet
 
    Dim i As Long
 
    Dim j As Long
 
    Dim lastRowF As Long
 
    Dim ultimaLinha As Long
 
    Dim foundCell As Range
 
    Dim valorG As Variant
    
    Dim CotaFixa As Double
 
    Dim nextLote As Long
    
    Dim erroPlaca As String
 
 
    Set wsAutomat = Workbooks.Open("C:\Users\thiago.slima\Desktop\AutoDoc.v1\Automat_OMAR.xlsm")
 
    Set wsPgtos = Workbooks.Open("C:\Users\thiago.slima\Desktop\AutoDoc.v1\PGTOS LOTES OMAR VAMOS.xls")
 
    ' Definir as folhas de trabalho
 
    Set wsAutomatSheet = wsAutomat.Sheets("Form_Omar")
 
    Set wsPgtosSheet = wsPgtos.Sheets("Relação Ativos")
    'wsPgtosSheet.AutoFilterMode = False
 
    Set wsResumoSheet = wsPgtos.Sheets("Resumo por Lote")
 
    ' Obter a última linha da Coluna F na planilha Automat
    
    CotaFixa = wsAutomatSheet.Cells(1, 1).Value
 
    nextLote = wsAutomatSheet.Cells(1, 2).Value
    
    'MsgBox "Cota:" & CotaFixa & "Lote:" & nextLote
    
    lastRowF = wsAutomatSheet.Cells(wsAutomatSheet.Rows.Count, "F").End(xlUp).Row
 
    ultimaLinha = wsResumoSheet.Cells(wsResumoSheet.Rows.Count, "G").End(xlUp).Row
 
    ' Loop pela Coluna F a partir da linha 3
 
    For i = 3 To lastRowF
 
        ' Verificar se a célula está vazia
 
        If wsAutomatSheet.Cells(i, "F").Value <> "" Then
 
            valorG = ""
 
            ' Procurar o primeiro valor válido na coluna G
 
            For j = 7 To ultimaLinha
 
                If wsResumoSheet.Cells(j, 7).Value <> 0 And wsResumoSheet.Cells(j, 7).Value <> "-" Then
 
                    valorG = wsResumoSheet.Cells(j, 7).Value
 
                    Exit For
 
                End If
 
            Next j
 
            ' Encontrar o valor na Coluna B da planilha PGTOS
 
            Set foundCell = wsPgtosSheet.Columns("B").Find(wsAutomatSheet.Cells(i, "F").Value, LookIn:=xlValues, LookAt:=xlWhole)

 
            If Not foundCell Is Nothing Then
 
                ' Inserir o Lote da Coluna B na Coluna P
 
                wsPgtosSheet.Cells(foundCell.Row, "P").Value = nextLote - 1
 
                ' Inserir o valor da Coluna G na Coluna Q
 
                wsPgtosSheet.Cells(foundCell.Row, "Q").Value = valorG
 
                ' Inserir o próximo lote na Coluna R
 
                wsPgtosSheet.Cells(foundCell.Row, "R").Value = nextLote
                 ' Calcular o valor na Coluna S (N - Q)
 
                wsPgtosSheet.Cells(foundCell.Row, "S").Value = wsPgtosSheet.Cells(foundCell.Row, "N").Value - wsPgtosSheet.Cells(foundCell.Row, "Q").Value


 
                If wsPgtosSheet.Cells(foundCell.Row, "Q").Value > wsPgtosSheet.Cells(foundCell.Row, "N").Value Then
                
                    wsPgtosSheet.Cells(foundCell.Row, "Q").Value = wsPgtosSheet.Cells(foundCell.Row, "N").Value
                    wsPgtosSheet.Cells(foundCell.Row, "R").Value = wsPgtosSheet.Cells(foundCell.Row, "P").Value
                    
                    wsPgtosSheet.Cells(foundCell.Row, "S").Value = wsPgtosSheet.Cells(foundCell.Row, "N").Value - wsPgtosSheet.Cells(foundCell.Row, "Q").Value
                    nextLote = nextLote - 1
                End If

 
                ' Verificar se o valor em S é maior que CotaFixa
 
                If wsPgtosSheet.Cells(foundCell.Row, "S").Value > CotaFixa Then
 
                    ' Inserir o Valor CotaFixa na Coluna S
 
                    wsPgtosSheet.Cells(foundCell.Row, "S").Value = CotaFixa
 
                    ' Inserir o lote + 1 na Coluna T
 
                    wsPgtosSheet.Cells(foundCell.Row, "T").Value = nextLote + 1
 
                    ' Fazer o cálculo T = N - Q - S
 
                    wsPgtosSheet.Cells(foundCell.Row, "U").Value = wsPgtosSheet.Cells(foundCell.Row, "N").Value - wsPgtosSheet.Cells(foundCell.Row, "Q").Value - wsPgtosSheet.Cells(foundCell.Row, "S").Value
                    
                    nextLote = nextLote + 1
                End If
 
                ' Incrementar o número do próximo lote
 
                nextLote = nextLote + 1
                
                 If wsPgtosSheet.Cells(foundCell.Row, "U").Value > CotaFixa Then
 
                    ' Inserir o Valor CotaFixa na Coluna S
 
                    wsPgtosSheet.Cells(foundCell.Row, "U").Value = CotaFixa
 
                    ' Inserir o lote + 1 na Coluna T
 
                    wsPgtosSheet.Cells(foundCell.Row, "V").Value = nextLote
 
                    ' Fazer o cálculo W = N - Q - S - U
 
                    wsPgtosSheet.Cells(foundCell.Row, "W").Value = wsPgtosSheet.Cells(foundCell.Row, "N").Value - wsPgtosSheet.Cells(foundCell.Row, "Q").Value - wsPgtosSheet.Cells(foundCell.Row, "S").Value - wsPgtosSheet.Cells(foundCell.Row, "U").Value
                    
                    
                End If
        Else
            
            erroPlaca = erroPlacas & wsAutomatSheet.Cells(i, "F").Value & vbCrLf
       
        End If
        
        If erroPlaca <> "" Then
        
            wsAutomatSheet.Cells(i, "G").Value = "NÃO LOCALIZADO"
            erroPlaca = ""
        
        Else
        
            wsAutomatSheet.Cells(i, "G").Value = "CONCLUÍDO"
            
        End If
        
    End If
 
    Next i
 
    ' Salvar e fechar as planilhas
 
    'wsPgtos.Close SaveChanges:=True
 
End Sub
