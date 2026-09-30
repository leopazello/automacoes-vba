Attribute VB_Name = "Escalas"
Option Explicit

Public Sub Escala_varios_eps_por_doc_e_contagem_de_aparicoes()
    Dim wsOrigem As Worksheet, wsEscala As Worksheet
    Dim dictPalavras As Object, dictEpisodios As Object
    Dim regEx As Object, matches As Object
    Dim ultimaLinha As Long, i As Long
    Dim valTimecode As Variant
    Dim personagem As String, fala As String, episodioAtual As String
    Dim k As Variant, eps As Variant, strEps As String

    On Error GoTo Falha
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set wsOrigem = ActiveSheet
    If wsOrigem.Name = "Escala" Then
        MsgBox "Selecione a aba com as falas antes de executar.", vbExclamation
        Exit Sub
    End If

    Set dictPalavras = CreateObject("Scripting.Dictionary")
    Set dictEpisodios = CreateObject("Scripting.Dictionary")
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = "\S+"
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "B").End(xlUp).Row
    episodioAtual = "Indefinido"

    For i = 1 To ultimaLinha
        If IsError(wsOrigem.Cells(i, 1).Value) Or _
           IsError(wsOrigem.Cells(i, 2).Value) Or _
           IsError(wsOrigem.Cells(i, 3).Value) Then GoTo ProximaLinha

        valTimecode = wsOrigem.Cells(i, 1).Value
        personagem = Trim(CStr(wsOrigem.Cells(i, 2).Value))
        fala = CStr(wsOrigem.Cells(i, 3).Value)
        If UCase(Left(Trim(CStr(valTimecode)), 2)) = "EP" Then
            episodioAtual = Trim(CStr(valTimecode))
        End If

        If (IsNumeric(valTimecode) Or IsDate(valTimecode)) And _
           CStr(valTimecode) <> "" And PersonagemValido(personagem) Then
            Set matches = regEx.Execute(fala)
            If matches.Count > 0 Then
                If Not dictPalavras.Exists(personagem) Then
                    dictPalavras.Add personagem, 0&
                    dictEpisodios.Add personagem, CreateObject("Scripting.Dictionary")
                End If
                dictPalavras(personagem) = dictPalavras(personagem) + matches.Count
                If Not dictEpisodios(personagem).Exists(episodioAtual) Then
                    dictEpisodios(personagem).Add episodioAtual, True
                End If
            End If
        End If
ProximaLinha:
    Next i

    Set wsEscala = PrepararEscala(wsOrigem)
    If wsEscala Is Nothing Then Exit Sub
    wsEscala.Range("A1:C1").Value = Array("Personagem", "Palavras", "Episodios")
    i = 2
    For Each k In dictPalavras.Keys
        wsEscala.Cells(i, 1).Value = k
        wsEscala.Cells(i, 2).Value = dictPalavras(k)
        strEps = ""
        For Each eps In dictEpisodios(k).Keys
            If strEps <> "" Then strEps = strEps & ", "
            strEps = strEps & eps
        Next eps
        wsEscala.Cells(i, 3).Value = strEps
        i = i + 1
    Next k
    FormatarEscala wsEscala, i - 1, 3
    MsgBox "Escala gerada com palavras e episodios por personagem.", vbInformation
    Exit Sub
Falha:
    MsgBox "Nao foi possivel concluir a escala: " & Err.Description, vbExclamation
End Sub

Public Sub Escala_Simples_Contagem_Palavras()
    Dim wsOrigem As Worksheet, wsEscala As Worksheet
    Dim dictPalavras As Object, regEx As Object, matches As Object
    Dim ultimaLinha As Long, i As Long, numColFala As Long
    Dim personagem As String, fala As String
    Dim colFala As Variant, k As Variant
    Dim numero As Double, indiceLetra As Long, codigoLetra As Long

    On Error GoTo Falha
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set wsOrigem = ActiveSheet
    If wsOrigem.Name = "Escala" Then
        MsgBox "Selecione a aba com as falas antes de executar.", vbExclamation
        Exit Sub
    End If

    colFala = Application.InputBox( _
        Prompt:="Qual coluna voce quer contar? Informe a letra ou o numero.", _
        Title:="Coluna das falas", Default:="C", Type:=2)
    If VarType(colFala) = vbBoolean Then Exit Sub
    colFala = Trim(UCase(CStr(colFala)))
    If colFala = "" Then Exit Sub

    If IsNumeric(colFala) Then
        ' Limita antes de converter para Long, evitando arredondamento e overflow.
        If Len(colFala) > 10 Then GoTo ColunaInvalida
        numero = CDbl(colFala)
        If numero < 1 Or numero > wsOrigem.Columns.Count Then GoTo ColunaInvalida
        If numero <> Fix(numero) Then GoTo ColunaInvalida
        numColFala = CLng(numero)
    Else
        If Len(colFala) > 3 Then GoTo ColunaInvalida
        For indiceLetra = 1 To Len(colFala)
            codigoLetra = Asc(Mid(colFala, indiceLetra, 1))
            If codigoLetra < 65 Or codigoLetra > 90 Then GoTo ColunaInvalida
            numColFala = numColFala * 26 + codigoLetra - 64
        Next indiceLetra
        If numColFala > wsOrigem.Columns.Count Then GoTo ColunaInvalida
    End If

    Set dictPalavras = CreateObject("Scripting.Dictionary")
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = "\S+"
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "B").End(xlUp).Row

    For i = 1 To ultimaLinha
        If IsError(wsOrigem.Cells(i, 2).Value) Or _
           IsError(wsOrigem.Cells(i, numColFala).Value) Then GoTo ProximaLinha
        personagem = Trim(CStr(wsOrigem.Cells(i, 2).Value))
        fala = CStr(wsOrigem.Cells(i, numColFala).Value)
        If PersonagemValido(personagem) Then
            Set matches = regEx.Execute(fala)
            If matches.Count > 0 Then
                If Not dictPalavras.Exists(personagem) Then dictPalavras.Add personagem, 0&
                dictPalavras(personagem) = dictPalavras(personagem) + matches.Count
            End If
        End If
ProximaLinha:
    Next i

    Set wsEscala = PrepararEscala(wsOrigem)
    If wsEscala Is Nothing Then Exit Sub
    wsEscala.Range("A1:B1").Value = Array("Personagem", "Palavras")
    i = 2
    For Each k In dictPalavras.Keys
        wsEscala.Cells(i, 1).Value = k
        wsEscala.Cells(i, 2).Value = dictPalavras(k)
        i = i + 1
    Next k
    FormatarEscala wsEscala, i - 1, 2
    MsgBox "Escala gerada! Coluna " & colFala & ".", vbInformation
    Exit Sub
ColunaInvalida:
    MsgBox "Coluna invalida. Use uma letra como C ou um numero como 3.", vbExclamation
    Exit Sub
Falha:
    MsgBox "Nao foi possivel concluir a escala: " & Err.Description, vbExclamation
End Sub

Private Function PersonagemValido(ByVal nome As String) As Boolean
    PersonagemValido = nome <> "" And UCase(nome) <> "PERSONAGEM" And UCase(nome) <> "PERSONAGENS"
End Function

Private Function PrepararEscala(ByVal origem As Worksheet) As Worksheet
    Dim existente As Worksheet, candidata As Worksheet
    Dim alertasAnteriores As Boolean
    Dim numeroErro As Long, descricaoErro As String

    alertasAnteriores = Application.DisplayAlerts
    On Error GoTo Falha
    For Each candidata In origem.Parent.Worksheets
        If candidata.Name = "Escala" Then
            Set existente = candidata
            Exit For
        End If
    Next candidata
    If Not existente Is Nothing Then
        If MsgBox("A aba Escala sera substituida. Continuar?", _
                  vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
        Application.DisplayAlerts = False
        existente.Delete
        Application.DisplayAlerts = alertasAnteriores
    End If
    Set PrepararEscala = origem.Parent.Worksheets.Add(After:=origem)
    PrepararEscala.Name = "Escala"
    Exit Function
Falha:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    Application.DisplayAlerts = alertasAnteriores
    Err.Raise numeroErro, "PrepararEscala", descricaoErro
End Function

Private Sub FormatarEscala(ByVal folha As Worksheet, ByVal ultimaLinha As Long, ByVal colunas As Long)
    Dim tabela As Range, cabecalho As Range, dados As Range
    Set tabela = folha.Range(folha.Cells(1, 1), folha.Cells(ultimaLinha, colunas))
    Set cabecalho = folha.Range(folha.Cells(1, 1), folha.Cells(1, colunas))
    If ultimaLinha > 1 Then
        tabela.Sort Key1:=folha.Range("A2"), Order1:=xlAscending, Header:=xlYes
    End If
    With cabecalho
        .Font.Bold = True
        .Interior.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlLeft
    End With
    tabela.Borders.LineStyle = xlNone
    tabela.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(0, 0, 0)
    cabecalho.Borders(xlEdgeBottom).LineStyle = xlContinuous
    If ultimaLinha > 1 Then
        Set dados = folha.Range(folha.Cells(2, 1), folha.Cells(ultimaLinha, colunas))
        With dados.Borders(xlInsideVertical)
            .LineStyle = xlDot
            .Weight = xlThin
            .Color = RGB(180, 180, 180)
        End With
        If ultimaLinha > 2 Then dados.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End If
    tabela.Columns.AutoFit
End Sub
