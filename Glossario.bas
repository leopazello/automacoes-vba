Attribute VB_Name = "Glossario"
Option Explicit

Public Sub SubstituirPalavrasDoGlossario()
    Dim wsOrigem As Worksheet, wsGlossario As Worksheet, candidata As Worksheet
    Dim celula As Range, repositorio As Object
    Dim ultimaLinha As Long, colunasAlvo As Variant, c As Variant
    Dim telaAnterior As Boolean, calculoAnterior As XlCalculation
    Dim descricaoErro As String

    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set wsOrigem = ActiveSheet
    telaAnterior = Application.ScreenUpdating
    calculoAnterior = Application.Calculation
    On Error GoTo Falha

    ' ChrW evita depender da codificacao do arquivo .bas para o nome acentuado.
    For Each candidata In wsOrigem.Parent.Worksheets
        If UCase(candidata.Name) = "GLOSSARIO" Or _
           UCase(candidata.Name) = "GLOSS" & ChrW(193) & "RIO" Then
            Set wsGlossario = candidata
            Exit For
        End If
    Next candidata
    If wsGlossario Is Nothing Then
        MsgBox "Aba GLOSSARIO nao encontrada nesta pasta de trabalho.", vbExclamation
        Exit Sub
    End If
    If wsOrigem Is wsGlossario Then
        MsgBox "Selecione a aba com os textos antes de executar.", vbExclamation
        Exit Sub
    End If

    Set repositorio = CreateObject("Scripting.Dictionary")
    CarregarRepositorio wsGlossario, repositorio
    If repositorio.Count = 0 Then
        MsgBox "O glossario nao possui termos abaixo do cabecalho.", vbInformation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    colunasAlvo = Array(3, 4)

    For Each c In colunasAlvo
        ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, c).End(xlUp).Row
        For Each celula In wsOrigem.Range(wsOrigem.Cells(1, c), wsOrigem.Cells(ultimaLinha, c))
            If Not celula.HasFormula And VarType(celula.Value) = vbString Then
                ProcessarCelula celula, repositorio
            End If
        Next celula
    Next c

    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior
    MsgBox "Substituicoes concluidas. Revise os trechos destacados.", vbInformation
    Exit Sub
Falha:
    descricaoErro = Err.Description
    Application.ScreenUpdating = telaAnterior
    Application.Calculation = calculoAnterior
    MsgBox "A execucao foi interrompida: " & descricaoErro & vbCrLf & _
           "Algumas substituicoes podem ter sido aplicadas. Confira sua copia.", vbExclamation
End Sub

Private Sub ProcessarCelula(ByVal celula As Range, ByVal repositorio As Object)
    Dim chave As Variant, novoTexto As String, textoAtual As String
    Dim regEx As Object, matches As Object, match As Object
    Dim deslocamento As Long, diferencaTam As Long

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    textoAtual = celula.Value
    For Each chave In repositorio.Keys
        regEx.Pattern = "\b" & EscapeRegExp(CStr(chave)) & "\b"
        Set matches = regEx.Execute(textoAtual)
        If matches.Count > 0 Then
            novoTexto = repositorio(chave)
            diferencaTam = Len(novoTexto) - Len(chave)
            deslocamento = 0
            For Each match In matches
                ' O deslocamento acompanha a mudanca de tamanho a cada troca.
                celula.Characters(match.FirstIndex + 1 + deslocamento, match.Length).Text = novoTexto
                If Len(novoTexto) > 0 Then
                    celula.Characters(match.FirstIndex + 1 + deslocamento, Len(novoTexto)).Font.Bold = True
                End If
                deslocamento = deslocamento + diferencaTam
            Next match
            textoAtual = celula.Value
        End If
    Next chave
End Sub

Private Sub CarregarRepositorio(ByVal ws As Worksheet, ByVal repo As Object)
    Dim i As Long, ultimaLinha As Long
    Dim chave As String, valor As String
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultimaLinha
        If Not IsError(ws.Cells(i, 1).Value) And Not IsError(ws.Cells(i, 2).Value) Then
            chave = Trim(Replace(CStr(ws.Cells(i, 1).Value), ChrW(160), " "))
            valor = Trim(CStr(ws.Cells(i, 2).Value))
            If chave <> "" Then
                If Not repo.Exists(chave) Then repo.Add chave, valor
            End If
        End If
    Next i
End Sub

Private Function EscapeRegExp(ByVal texto As String) As String
    Dim chars As Variant, i As Long
    chars = Array("\", ".", "*", "+", "?", "|", "(", ")", "[", "]", "{", "}", "^", "$")
    For i = LBound(chars) To UBound(chars)
        texto = Replace(texto, chars(i), "\" & chars(i))
    Next i
    EscapeRegExp = texto
End Function
