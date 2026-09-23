Attribute VB_Name = "modTrocaTurno"
'==============================================================================
' TROCA DE TURNO 2026 - macros opcionais
'   RegistrarLancamento : grava o formulário da aba LANÇAMENTO na BASE_OPERACIONAL
'   LimparFormulario    : limpa as células amarelas do formulário
'   CopiarResumoTurno   : copia o resumo da PASSAGEM_DE_TURNO para a área de transferência
'   RegistrarLog        : grava uma linha no LOG_ALTERAÇÕES (usado também pelos eventos)
' Instalação: Alt+F11 > Arquivo > Importar arquivo... > modTrocaTurno.bas
'             e cole EstaPastaDeTrabalho.cls em "EstaPasta_de_trabalho". Salve como .xlsm.
'==============================================================================
Option Explicit

Public Const SENHA As String = "tt2026"
Public Const ABA_BASE As String = "BASE_OPERACIONAL"
Public Const ABA_FORM As String = "LANÇAMENTO"
Public Const ABA_LOG As String = "LOG_ALTERAÇÕES"
Public Const ABA_PASSAGEM As String = "PASSAGEM_DE_TURNO"
Public Const PRIMEIRA_LINHA As Long = 2

' ordem dos campos do formulário (C5 em diante) -> cabeçalho na BASE
Private Function CamposFormulario() As Variant
    CamposFormulario = Array("DATA", "TURNO", "RESPONSÁVEL PELO TURNO", "TIPO DE ATIVIDADE", "PROGRAMAÇÃO", _
        "ROTA", "MEDIDA", "PLANEJADO", "EXECUTADO", "FATURAMENTO", "EMBARQUE", "VOLUME FATURADO", _
        "TRANSPORTADORA", "PENDÊNCIA", "TIPO DE PENDÊNCIA", "CAUSA", "RESPONSÁVEL PELA TRATATIVA", _
        "PRAZO", "AÇÃO", "IMPACTO OPERACIONAL", "OBSERVAÇÃO")
End Function

Public Sub ProtegerTudo()
    ' Proteção com UserInterfaceOnly: o usuário não altera fórmulas, mas as macros podem gravar.
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        sh.Protect Password:=SENHA, UserInterfaceOnly:=True, AllowFiltering:=True, _
                   AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next sh
End Sub

Public Function ColunaPorCabecalho(ws As Worksheet, cabecalho As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(What:=cabecalho, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If c Is Nothing Then ColunaPorCabecalho = 0 Else ColunaPorCabecalho = c.Column
End Function

Public Function ProximaLinhaLivre(ws As Worksheet) As Long
    Dim colData As Long, r As Long, ultima As Long
    colData = ColunaPorCabecalho(ws, "DATA")
    ultima = ws.ListObjects(1).Range.Rows.Count
    For r = PRIMEIRA_LINHA To ultima
        If IsEmpty(ws.Cells(r, colData).Value) Then
            ProximaLinhaLivre = r
            Exit Function
        End If
    Next r
    ProximaLinhaLivre = 0
End Function

Public Sub RegistrarLancamento()
    Dim wf As Worksheet, wb As Worksheet
    Dim campos As Variant, i As Long, linha As Long, col As Long
    Dim obrig As Variant, faltando As String, idNovo As String

    Set wf = ThisWorkbook.Worksheets(ABA_FORM)
    Set wb = ThisWorkbook.Worksheets(ABA_BASE)
    campos = CamposFormulario()

    ' obrigatórios: data, turno, responsável, tipo de atividade, programação, medida, planejado, executado
    obrig = Array(0, 1, 2, 3, 4, 6, 7, 8)
    For i = LBound(obrig) To UBound(obrig)
        If Trim(CStr(wf.Range("C5").Offset(obrig(i), 0).Value)) = "" Then
            faltando = faltando & vbCrLf & " - " & campos(obrig(i))
        End If
    Next i
    If Trim(CStr(wf.Range("C18").Value)) <> "" Then
        If Trim(CStr(wf.Range("C19").Value)) = "" Then faltando = faltando & vbCrLf & " - TIPO DE PENDÊNCIA"
        If Trim(CStr(wf.Range("C21").Value)) = "" Then faltando = faltando & vbCrLf & " - RESPONSÁVEL PELA TRATATIVA"
        If Trim(CStr(wf.Range("C22").Value)) = "" Then faltando = faltando & vbCrLf & " - PRAZO"
    End If
    If faltando <> "" Then
        MsgBox "Preencha os campos obrigatórios:" & faltando, vbExclamation, "Lançamento incompleto"
        Exit Sub
    End If

    linha = ProximaLinhaLivre(wb)
    If linha = 0 Then
        MsgBox "A BASE_OPERACIONAL está cheia. Amplie a capacidade (ver COMO_USAR).", vbCritical
        Exit Sub
    End If

    Application.EnableEvents = False
    On Error GoTo Falha
    For i = LBound(campos) To UBound(campos)
        col = ColunaPorCabecalho(wb, CStr(campos(i)))
        If col > 0 Then
            If Trim(CStr(wf.Range("C5").Offset(i, 0).Value)) <> "" Then
                wb.Cells(linha, col).Value = wf.Range("C5").Offset(i, 0).Value
            End If
        End If
    Next i
    wb.Cells(linha, ColunaPorCabecalho(wb, "DATA/HORA DA ATUALIZAÇÃO")).Value = Now
    Application.Calculate
    idNovo = CStr(wb.Cells(linha, ColunaPorCabecalho(wb, "ID")).Value)
    RegistrarLog idNovo, "(novo registro)", "", "linha " & linha, "FORMULÁRIO"
    Application.EnableEvents = True

    LimparFormulario
    MsgBox "Registro " & idNovo & " gravado na linha " & linha & ".", vbInformation, "Lançamento"
    Exit Sub
Falha:
    Application.EnableEvents = True
    MsgBox "Erro ao gravar: " & Err.Description, vbCritical
End Sub

Public Sub LimparFormulario()
    ThisWorkbook.Worksheets(ABA_FORM).Range("C5:C25").ClearContents
End Sub

Public Sub RegistrarLog(registro As String, campo As String, antes As String, depois As String, origem As String)
    Dim wl As Worksheet, r As Long
    Set wl = ThisWorkbook.Worksheets(ABA_LOG)
    r = wl.Cells(wl.Rows.Count, 1).End(xlUp).Row + 1
    wl.Cells(r, 1).Value = Now
    wl.Cells(r, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    wl.Cells(r, 2).Value = Application.UserName & " (" & Environ$("USERNAME") & ")"
    wl.Cells(r, 3).Value = registro
    wl.Cells(r, 4).Value = campo
    wl.Cells(r, 5).Value = Left$(antes, 250)
    wl.Cells(r, 6).Value = Left$(depois, 250)
    wl.Cells(r, 7).Value = origem
End Sub

Public Sub CopiarResumoTurno()
    Dim txt As String, obj As Object, c As Range
    For Each c In ThisWorkbook.Worksheets(ABA_PASSAGEM).UsedRange.Columns(2).Cells
        If Left$(CStr(c.Value), 6) = "TURNO " And InStr(CStr(c.Value), "Planejado:") > 0 Then
            txt = CStr(c.Value)
            Exit For
        End If
    Next c
    If txt = "" Then
        MsgBox "Resumo não encontrado.", vbExclamation
        Exit Sub
    End If
    Set obj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  ' MSForms.DataObject
    obj.SetText txt
    obj.PutInClipboard
    MsgBox "Resumo copiado. Cole no WhatsApp/e-mail (Ctrl+V).", vbInformation
End Sub
