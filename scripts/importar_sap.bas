## Sub principal que é executada inicialmente.
Sub Importar_SAP()

    On Error GoTo Erro
    
    '========================
    ' DECLARAÇÕES
    '========================
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsDesconsiderar As Worksheet
    
    Dim caminhoArquivo As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim texto As String
    Dim resultado As Variant
    
    '========================
    ' CONFIGURAÇÕES INICIAIS
    '========================
    Set wsDestino = ThisWorkbook.Sheets("SAP_RAW")
    Set wsDesconsiderar = ThisWorkbook.Sheets("desconsiderar")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '========================
    ' EXTRAÇÃO SAP (AUTOMAÇÃO)
    '========================
    ' Esta função executa a extração via SAP GUI scripting
    Call Baixar_Relatorio_SAP

    ' Caminho genérico (REMOVIDO DADO SENSÍVEL)
    caminhoArquivo = "C:\CAMINHO\GENERIC\backlog.xlsx"
    
    '========================
    ' IMPORTAÇÃO DE DADOS
    '========================
    Set wbOrigem = Workbooks.Open(caminhoArquivo)
    Set wsOrigem = wbOrigem.Sheets(1)

    ' Limpar dados anteriores
    wsDestino.Cells.Clear

    ' Copiar dados do relatório SAP
    wsOrigem.UsedRange.Copy
    wsDestino.Range("A1").PasteSpecial xlPasteValues

    ' Remover linhas vazias (coluna H)
    wsDestino.Range("A1").AutoFilter Field:=8, Criteria1:="="
    
    On Error Resume Next
    wsDestino.Range("A2:A" & wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row) _
        .SpecialCells(xlCellTypeVisible).EntireRow.Delete
    On Error GoTo 0
    
    wsDestino.AutoFilterMode = False

    ' Fechar arquivo origem
    wbOrigem.Close False

    '========================
    ' FORMATAÇÃO
    '========================
    wsDestino.Columns("B").NumberFormat = "hh:mm"
    wsDestino.Columns("I").NumberFormat = "dd/mm/yyyy"

    ' Forçar valores (remover fórmulas)
    wsDestino.Columns("A").Value = wsDestino.Columns("A").Value
    wsDestino.Columns("C").Value = wsDestino.Columns("C").Value
    wsDestino.Columns("D").Value = wsDestino.Columns("D").Value

    ' Ajuste de prioridade (exemplo)
    wsDestino.Columns("C").Replace What:="0", Replacement:="4", LookAt:=xlWhole

    '========================
    ' ADICIONAR COLUNAS AUXILIARES
    '========================
    wsDestino.Cells(1, "P").Value = "STATUS"
    wsDestino.Cells(1, "Q").Value = "CONSIDERAR"
    wsDestino.Cells(1, "R").Value = "JOB_FINALIZADO"

    '========================
    ' REGRA 1: CLASSIFICAÇÃO DE ITENS
    '========================
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "E").End(xlUp).Row

    For i = 2 To ultimaLinha

        texto = wsDestino.Cells(i, "E").Value

        ' Exemplo de classificação por tipo de produto
        If Left(texto, 7) = "EXEMPLO1" _
        Or Left(texto, 3) = "ABC" _
        Or Left(texto, 6) = "ITEMX" Then
        
            wsDestino.Cells(i, "C").Value = "CLASSIFICADO"
        
        End If

    Next i

    '========================
    ' REGRA 2: VALIDAR CONSIDERAÇÃO
    '========================
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "D").End(xlUp).Row

    For i = 2 To ultimaLinha

        resultado = Application.VLookup(wsDestino.Cells(i, "D").Value, _
                    wsDesconsiderar.Range("A:C"), 3, False)

        If IsError(resultado) Then
            wsDestino.Cells(i, "Q").Value = "SIM"
        Else
            wsDestino.Cells(i, "Q").Value = resultado
        End If

    Next i

    '========================
    ' PROCESSOS COMPLEMENTARES
    '========================
    Call Atualizar_Apex_Andamento
    Call Aplicar_Regras_STU
    Call Criar_Tabela_Dinamica
    Call Cruzar_Apex_Finalizado

    '========================
    ' REGRA FINAL
    '========================
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "R").End(xlUp).Row

    For i = 2 To ultimaLinha
        If Trim(wsDestino.Cells(i, "R").Value) <> "" _
        And Trim(wsDestino.Cells(i, "R").Value) <> "0" Then
            
            wsDestino.Cells(i, "P").Value = "AGUARDANDO PROCESSAMENTO"
        
        End If
    Next i

Fim:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

Erro:
    MsgBox "Erro: " & Err.Description
    Resume Fim

End Sub

--------------------------------------------------------------------------------------------------------------------------------
## sub´s secundarias que a primeira chama.

Sub Aplicar_Regras_STU()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim valorN As String

    Set ws = ThisWorkbook.Sheets("SAP_RAW")

    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaLinha

        'Só aplica se STU estiver vazio
        If ws.Cells(i, "P").Value = "" Then
            
            valorN = ws.Cells(i, "N").Value
            
            If valorN = "A" Then
                ws.Cells(i, "P").Value = "AG. SEPARAÇÃO"
            Else
                ws.Cells(i, "P").Value = "AG. FATURAMENTO"
            End If

        End If

    Next i

End Sub

------------------------------------------------------------------------------------------------------------------------
Sub Aplicar_Regras_STU()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim valorN As String

    Set ws = ThisWorkbook.Sheets("SAP_RAW")

    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaLinha

        'Só aplica se STU estiver vazio
        If ws.Cells(i, "P").Value = "" Then
            
            valorN = ws.Cells(i, "N").Value
            
            If valorN = "A" Then
                ws.Cells(i, "P").Value = "AG. SEPARAÇÃO"
            Else
                ws.Cells(i, "P").Value = "AG. FATURAMENTO"
            End If

        End If

    Next i

End Sub
---------------------------------------------------------------------------------------------------------------

Sub Criar_Tabela_Dinamica()

Dim wsDados As Worksheet
Dim wsPivot As Worksheet
Dim ultimaLinha As Long
Dim ultimaColuna As Long
Dim intervalo As Range

Dim pc As PivotCache
Dim pt As PivotTable

Set wsDados = ThisWorkbook.Sheets("SAP_RAW")
Set wsPivot = ThisWorkbook.Sheets("DASHBOARD")

wsPivot.Cells.Clear

ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "A").End(xlUp).Row
ultimaColuna = wsDados.Cells(1, wsDados.Columns.Count).End(xlToLeft).Column

Set intervalo = wsDados.Range(wsDados.Cells(1, 1), wsDados.Cells(ultimaLinha, ultimaColuna))

Set pc = ThisWorkbook.PivotCaches.Create( _
SourceType:=xlDatabase, _
SourceData:=intervalo)

Set pt = pc.CreatePivotTable( _
TableDestination:=wsPivot.Range("A3"), _
TableName:="Pivot_Apex")

With pt

    'FILTRO
    .PivotFields("considerar").Orientation = xlPageField
    .PivotFields("considerar").Position = 1
    
    'LINHA 1
    .PivotFields("STU").Orientation = xlRowField
    .PivotFields("STU").Position = 1

    'LINHA 2
    .PivotFields("Prioridade remessa").Orientation = xlRowField
    .PivotFields("Prioridade remessa").Position = 2

    'COLUNAS
    .PivotFields("Criado em").Orientation = xlColumnField
    .PivotFields("Criado em").Position = 1

    'VALOR
    .AddDataField .PivotFields("Qtd.remessa"), _
    "Soma de Qtd.remessa", xlSum

End With

'Aplicar filtro automaticamente
With pt.PivotFields("considerar")
    .ClearAllFilters
    .CurrentPage = "sim"
End With

End Sub

-------------------------------------------------------------------------------------------------------

Sub Preencher_STU()

Dim wsSAP As Worksheet
Dim wsApex As Worksheet
Dim ultimaLinhaSAP As Long
Dim i As Long
Dim resultado As Variant

Set wsSAP = ThisWorkbook.Sheets("SAP_RAW")
Set wsApex = ThisWorkbook.Sheets("APEX_ANDAMENTO")

ultimaLinhaSAP = wsSAP.Cells(wsSAP.Rows.Count, "A").End(xlUp).Row

For i = 2 To ultimaLinhaSAP

    resultado = application.VLookup(wsSAP.Cells(i, "A").Value, _
                wsApex.Range("B:E"), 4, False)

    If Not IsError(resultado) Then
        wsSAP.Cells(i, "P").Value = resultado
    End If

Next i

End Sub

-----------------------------------------------------------------

Sub Baixar_Relatorio_SAP()

    Dim SapGuiAuto As Object
    Dim application As Object
    Dim Connection As Object
    Dim Session As Object

    On Error GoTo ErroSAP

    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
    Set Connection = application.Children(0)
    Set Session = Connection.Children(0)

    Session.findById("wnd[0]").maximize
    
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "vl06f"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]").sendVKey 17
    Session.findById("wnd[1]/tbar[0]/btn[6]").press
    Session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
    Session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
    Session.findById("wnd[0]").sendVKey 8
    Session.findById("wnd[0]/mbar/menu[3]/menu[2]/menu[1]").Select
    Session.findById("wnd[1]/usr").verticalScrollbar.Position = 20
    Session.findById("wnd[1]/tbar[0]/btn[71]").press
    Session.findById("wnd[2]/usr/txtRSYSF-STRING").Text = "den"
    Session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = 3
    Session.findById("wnd[2]/tbar[0]/btn[0]").press
    Session.findById("wnd[3]/usr/sub/1[0,0]/sub/1/3[0,2]/lbl[1,2]").SetFocus
    Session.findById("wnd[3]/usr/sub/1[0,0]/sub/1/3[0,2]/lbl[1,2]").caretPosition = 5
    Session.findById("wnd[3]").sendVKey 2
    Session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/3[0,3]/lbl[1,3]").caretPosition = 4
    Session.findById("wnd[1]").sendVKey 2
    Session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/7[0,7]/lbl[15,7]").SetFocus
    Session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/7[0,7]/lbl[15,7]").caretPosition = 2
    Session.findById("wnd[0]/mbar/menu[0]/menu[5]/menu[1]").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "N:\SERVICOS\Logística de Serviços\Logística Estoque Expedição\2026\3-OPERAÇÃO DIARIA\BACKLOG"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "backlogatual.XLSX"
    Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
    Session.findById("wnd[1]/tbar[0]/btn[11]").press
    Session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/12[0,12]/lbl[4,12]").SetFocus
    Session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/12[0,12]/lbl[4,12]").caretPosition = 2
    Session.findById("wnd[0]").sendVKey 3
    Session.findById("wnd[0]").sendVKey 3
    
    Exit Sub

ErroSAP:
    MsgBox "Erro ao rodar SAP: " & Err.Description

End Sub

---------------------------------------------------------------------------------

Sub Cruzar_Apex_Finalizado()

    Dim wsSAP As Worksheet
    Dim wsApex As Worksheet
    Dim ultimaLinhaSAP As Long
    Dim i As Long
    Dim resultado As Variant

    Set wsSAP = ThisWorkbook.Sheets("SAP_RAW")
    Set wsApex = ThisWorkbook.Sheets("APEX_FINALIZADO")

    ultimaLinhaSAP = wsSAP.Cells(wsSAP.Rows.Count, "A").End(xlUp).Row

    'Cabeçalho (opcional)
    wsSAP.Cells(1, "R").Value = "APEX_FINALIZADO"

    For i = 2 To ultimaLinhaSAP

        resultado = application.VLookup(wsSAP.Cells(i, "A").Value, _
                    wsApex.Range("A:D"), 4, False)

        If Not IsError(resultado) Then
            wsSAP.Cells(i, "R").Value = resultado
        Else
            wsSAP.Cells(i, "R").Value = ""
        End If

    Next i

End Sub


