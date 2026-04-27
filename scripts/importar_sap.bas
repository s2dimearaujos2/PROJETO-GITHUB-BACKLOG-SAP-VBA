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
