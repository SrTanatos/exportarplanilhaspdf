Attribute VB_Name = "M�dulo5"
Sub ExportarTodasPlanilhasPDF()

    ' Declara as vari�veis
    Dim pastaBase As String
    Dim pastaDestino As String
    Dim nomeArquivo As String
    Dim ws As Worksheet

    ' Obt�m o caminho da pasta atual
    pastaBase = ThisWorkbook.Path

    ' Verifica se a pasta de trabalho foi salva (ter� um caminho v�lido)
    If pastaBase = "" Then
        MsgBox "Por favor, salve sua pasta de trabalho antes de continuar.", vbExclamation
        Exit Sub
    End If

    ' Itera sobre todas as planilhas
    For Each ws In ThisWorkbook.Sheets

        ' Verifica se a planilha deve ser ignorada
        If DeveIgnorar(ws) Then
            ' Pula a itera��o atual do loop e passa para a pr�xima planilha
            GoTo ProximaPlanilha
        End If

        ' Obt�m o nome da planilha
        nomeArquivo = ws.Name

        ' Cria o nome da subpasta de destino
        pastaDestino = pastaBase & "\" & nomeArquivo

        ' Verifica se a subpasta existe, caso contr�rio, cria a subpasta
        If Dir(pastaDestino, vbDirectory) = "" Then
            MkDir pastaDestino
        End If

        ' Exporta a planilha como PDF para a subpasta criada
        On Error Resume Next
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=pastaDestino & "\" & nomeArquivo & ".pdf"
        On Error GoTo 0

        If Err.Number <> 0 Then
            MsgBox "Erro ao exportar '" & nomeArquivo & "': " & Err.Description, vbCritical
        End If

ProximaPlanilha:
    Next ws

End Sub

Function NomeValido(ByVal nome As String) As Boolean
   
End Function

Function DeveIgnorar(ws As Worksheet) As Boolean
    Dim nomesIgnorados As Variant
   
    ' Lista de nomes de planilhas a serem ignorados
    nomesIgnorados = Array("CAPA", "Resumo", "Guia", "Datas BM`s", "PQ") ' Adicione ou remova nomes conforme necess�rio

    ' Define a cor vermelho (em RGB)
    corVermelho = RGB(255, 0, 0)
    
    ' Verifica o nome
    If IsInArray(ws.Name, nomesIgnorados) Then
        DeveIgnorar = True
        Exit Function
    End If
    
        ' Verifica o valor da c�lula H11
    If ws.Range("H11").Value = 0 Then
        DeveIgnorar = True
        Exit Function
    End If

    
    DeveIgnorar = False
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function



