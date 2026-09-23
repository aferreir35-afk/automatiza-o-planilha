'==============================================================================
' Cole este código no módulo "EstaPasta_de_trabalho" (ThisWorkbook) do arquivo .xlsm.
' - Ao abrir: protege as abas com UserInterfaceOnly (macros podem gravar).
' - Ao alterar a BASE_OPERACIONAL: grava DATA/HORA DA ATUALIZAÇÃO e o LOG_ALTERAÇÕES
'   (data/hora, usuário, ID, campo, valor anterior e novo).
'==============================================================================
Option Explicit

Private valoresAntes As Object   ' endereço -> valor antes da edição

Private Sub Workbook_Open()
    modTrocaTurno.ProtegerTudo
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim c As Range
    If Sh.Name <> modTrocaTurno.ABA_BASE Then Exit Sub
    Set valoresAntes = CreateObject("Scripting.Dictionary")
    If Target.CountLarge > 2000 Then Exit Sub
    For Each c In Target.Cells
        valoresAntes(c.Address) = CStr(c.Value)
    Next c
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim c As Range, colAtual As Long, colId As Long, antes As String
    If Sh.Name <> modTrocaTurno.ABA_BASE Then Exit Sub
    If Target.CountLarge > 2000 Then Exit Sub

    colAtual = modTrocaTurno.ColunaPorCabecalho(Sh, "DATA/HORA DA ATUALIZAÇÃO")
    colId = modTrocaTurno.ColunaPorCabecalho(Sh, "ID")
    Application.EnableEvents = False
    On Error GoTo Fim
    For Each c In Target.Cells
        If c.Row >= modTrocaTurno.PRIMEIRA_LINHA And c.Column <> colAtual And Not c.HasFormula Then
            antes = ""
            If Not valoresAntes Is Nothing Then
                If valoresAntes.Exists(c.Address) Then antes = valoresAntes(c.Address)
            End If
            If antes <> CStr(c.Value) Then
                Sh.Cells(c.Row, colAtual).Value = Now
                Application.Calculate
                modTrocaTurno.RegistrarLog CStr(Sh.Cells(c.Row, colId).Value), CStr(Sh.Cells(1, c.Column).Value), _
                                           antes, CStr(c.Value), "EDIÇÃO NA BASE"
            End If
        End If
    Next c
Fim:
    Application.EnableEvents = True
End Sub
