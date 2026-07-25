Attribute VB_Name = "DesignModule"
Option Explicit

' Met à niveau la feuille sans effacer des colonnes entières ni les données
' de l'utilisateur. Cette macro est également appelée par le script
' scripts/Sync-VbaProject.ps1.
Public Sub Update_Design()
    Dim previousScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo DESIGN_ERROR
    Application.ScreenUpdating = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Portfolio.SHEET_PORT)

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    MigrateLegacySampleData ws

    ' Paramètre optionnel de dividende continu q.
    ws.Range("B12").Value = "Dividende (q)"
    If Len(Trim$(CStr(ws.Range("B13").Value2))) = 0 Then
        ws.Range("B13").Value = 0#
    End If

    With ws.Range("B12")
        .Font.Color = vbWhite
        .Font.Size = 10
    End With

    With ws.Range("B13")
        .Interior.Color = vbWhite
        .Font.Color = RGB(0, 0, 255)
        .HorizontalAlignment = xlRight
        .NumberFormat = "0.00%"
    End With

    ' Les hypothèses éditables sont affichées en bleu.
    ws.Range("B5,B8,B11,B13,B17,B20,B23").Font.Color = RGB(0, 0, 255)
    ws.Range("B5,B17,B20,B23").NumberFormat = "0.00"
    ws.Range("B8,B11,B13").NumberFormat = "0.00%"

    ws.Range("M2").Value = "ANALYSE PnL & RISQUE — À RECALCULER"
    ws.Range("M2").Font.Size = 14
    ws.Range("M2").Font.Bold = True
    ws.Range("M2").Font.Color = RGB(45, 62, 80)

    ' Validation de la table de positions.
    ApplyListValidation ws.Range("G4:G200"), "CALL", "PUT"
    ApplyListValidation ws.Range("H4:H200"), "Achat", "Vente"
    ApplyPositiveDecimalValidation ws.Range("I4:J200")
    ApplyPositiveWholeValidation ws.Range("K4:K200")
    ApplyNonNegativeDecimalValidation ws.Range("L4:L200")

    ws.Range("I4:I200").NumberFormat = "0.00"
    ws.Range("J4:J200").NumberFormat = "0.00"
    ws.Range("K4:K200").NumberFormat = "0"
    ws.Range("L4:L200").NumberFormat = "0.0000"

    ws.Range("E3:L3").Font.Bold = True
    ws.Range("E3:L3").Font.Color = vbWhite
    ws.Range("E3:L3").Interior.Color = RGB(0, 112, 192)

    Portfolio.InvalidateResults
    GoTo DESIGN_EXIT

DESIGN_ERROR:
    MsgBox "La mise en forme n'a pas pu être terminée : " & Err.Description, _
           vbExclamation, "Mise en forme"

DESIGN_EXIT:
    Application.ScreenUpdating = previousScreenUpdating
End Sub

' Corrige uniquement le jeu d'exemple historique livré dans le dépôt.
' Les portefeuilles personnalisés ne sont jamais réécrits silencieusement.
Private Sub MigrateLegacySampleData(ByVal ws As Worksheet)
    Dim cell As Range
    For Each cell In ws.Range("E4:E6,F4:F6,J4:J6").Cells
        If IsError(cell.Value2) Then Exit Sub
    Next cell

    If ws.Range("E4").Value2 <> 1 Or _
       ws.Range("E5").Value2 <> 2 Or _
       ws.Range("E6").Value2 <> 3 Then Exit Sub

    If ws.Range("J4").Value2 <> 100 Or _
       ws.Range("J5").Value2 <> 105 Or _
       ws.Range("J6").Value2 <> 120 Then Exit Sub

    If StrComp(Trim$(CStr(ws.Range("F4").Value2)), "Asset", vbTextCompare) = 0 And _
       StrComp(Trim$(CStr(ws.Range("F5").Value2)), "Asset", vbTextCompare) = 0 And _
       StrComp(Trim$(CStr(ws.Range("F6").Value2)), "Tesla", vbTextCompare) = 0 Then
        ws.Range("F6").Value = "Asset"
    End If
End Sub

Private Sub ApplyListValidation(ByVal target As Range, ByVal firstChoice As String, _
                                ByVal secondChoice As String)
    Dim separator As String
    separator = Application.International(xlListSeparator)

    target.Validation.Delete
    target.Validation.Add Type:=xlValidateList, _
                          AlertStyle:=xlValidAlertStop, _
                          Formula1:=firstChoice & separator & secondChoice
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.ErrorTitle = "Valeur invalide"
    target.Validation.ErrorMessage = "Choisissez une valeur dans la liste."
End Sub

Private Sub ApplyPositiveDecimalValidation(ByVal target As Range)
    target.Validation.Delete
    target.Validation.Add Type:=xlValidateDecimal, _
                          AlertStyle:=xlValidAlertStop, _
                          Operator:=xlGreater, Formula1:="0"
    target.Validation.IgnoreBlank = True
    target.Validation.ErrorTitle = "Valeur invalide"
    target.Validation.ErrorMessage = "Saisissez une valeur strictement positive."
End Sub

Private Sub ApplyPositiveWholeValidation(ByVal target As Range)
    target.Validation.Delete
    target.Validation.Add Type:=xlValidateWholeNumber, _
                          AlertStyle:=xlValidAlertStop, _
                          Operator:=xlGreater, Formula1:="0"
    target.Validation.IgnoreBlank = True
    target.Validation.ErrorTitle = "Maturité invalide"
    target.Validation.ErrorMessage = "Saisissez un nombre entier de jours positif."
End Sub

Private Sub ApplyNonNegativeDecimalValidation(ByVal target As Range)
    target.Validation.Delete
    target.Validation.Add Type:=xlValidateDecimal, _
                          AlertStyle:=xlValidAlertStop, _
                          Operator:=xlGreaterEqual, Formula1:="0"
    target.Validation.IgnoreBlank = True
    target.Validation.ErrorTitle = "Prime invalide"
    target.Validation.ErrorMessage = "La prime doit être positive ou nulle."
End Sub
