VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPortfolio 
   Caption         =   "Portfolio "
   ClientHeight    =   3815
   ClientLeft      =   7
   ClientTop       =   -14
   ClientWidth     =   5495
   OleObjectBlob   =   "frmPortfolio.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPortfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.BackColor = RGB(250, 250, 250)

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CommandButton" Then
            ctrl.BackColor = RGB(0, 120, 215)
            ctrl.ForeColor = vbWhite
            ctrl.Font.Bold = True
        ElseIf TypeName(ctrl) = "Label" Then
            ctrl.ForeColor = RGB(80, 80, 80)
        ElseIf TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
            ctrl.BackColor = vbWhite
            ctrl.BorderStyle = fmBorderStyleSingle
            ctrl.BorderColor = RGB(200, 200, 200)
        End If
    Next ctrl

    txtUnderlying.Text = "Asset"
    cboType.Clear
    cboType.AddItem "CALL"
    cboType.AddItem "PUT"
    cboType.Value = "CALL"

    cboSide.Clear
    cboSide.AddItem "Achat"
    cboSide.AddItem "Vente"
    cboSide.Value = "Achat"

    txtQty.Text = "1"
    txtK.Text = "100"
    txtMatDays.Text = "30"
    txtPrime.Text = ""

    RefreshList
End Sub

Private Sub cmdAdd_Click()
    Dim underlying As String
    Dim quantity As Double
    Dim strike As Double
    Dim maturityDays As Double
    Dim premium As Double
    Dim hasPremium As Boolean

    underlying = Trim$(txtUnderlying.Text)
    If Len(underlying) = 0 Then
        MsgBox "Le nom du sous-jacent est obligatoire.", _
               vbExclamation, "Position invalide"
        Exit Sub
    End If

    If Not TryParseDouble(txtQty.Text, quantity) Or quantity <= 0# Then
        MsgBox "La quantité doit être strictement positive.", _
               vbExclamation, "Position invalide"
        Exit Sub
    End If

    If Not TryParseDouble(txtK.Text, strike) Or strike <= 0# Then
        MsgBox "Le strike doit être strictement positif.", _
               vbExclamation, "Position invalide"
        Exit Sub
    End If

    If Not TryParseDouble(txtMatDays.Text, maturityDays) Or maturityDays <= 0# Then
        MsgBox "La maturité doit être un nombre de jours strictement positif.", _
               vbExclamation, "Position invalide"
        Exit Sub
    End If

    If Abs(maturityDays - Fix(maturityDays)) > 0.0000001 Then
        MsgBox "La maturité doit être saisie en nombre entier de jours.", _
               vbExclamation, "Position invalide"
        Exit Sub
    End If

    hasPremium = Len(Trim$(txtPrime.Text)) > 0
    If hasPremium Then
        If Not TryParseDouble(txtPrime.Text, premium) Or premium < 0# Then
            MsgBox "La prime doit être positive ou nulle. " & _
                   "Laissez le champ vide pour utiliser la prime théorique.", _
                   vbExclamation, "Position invalide"
            Exit Sub
        End If
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Portfolio.SHEET_PORT)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    If lastRow >= 4 Then
        Dim existingUnderlying As String
        Dim existingMaturity As Double

        existingUnderlying = Trim$(CStr(ws.Cells(4, "F").Value2))
        If Len(existingUnderlying) > 0 Then
            If StrComp(existingUnderlying, underlying, vbTextCompare) <> 0 Then
                MsgBox "Ce modèle accepte un seul sous-jacent : " & _
                       existingUnderlying & ".", _
                       vbExclamation, "Sous-jacent différent"
                Exit Sub
            End If
        End If

        If IsNumeric(ws.Cells(4, "K").Value2) Then
            existingMaturity = CDbl(ws.Cells(4, "K").Value2)
            If Abs(existingMaturity - maturityDays) > 0.0000001 Then
                MsgBox "Toutes les positions doivent avoir la même maturité " & _
                       "pour construire un PnL à l'échéance cohérent.", _
                       vbExclamation, "Maturité différente"
                Exit Sub
            End If
        End If
    End If

    Dim targetRow As Long
    targetRow = lastRow + 1
    If targetRow < 4 Then targetRow = 4

    Dim positionId As Long
    positionId = 1
    If targetRow > 4 Then
        positionId = CLng(Application.Max(ws.Range("E4:E" & targetRow - 1))) + 1
    End If

    ws.Cells(targetRow, "E").Value = positionId
    ws.Cells(targetRow, "F").Value = underlying
    ws.Cells(targetRow, "G").Value = UCase$(cboType.Value)
    ws.Cells(targetRow, "H").Value = cboSide.Value
    ws.Cells(targetRow, "I").Value = quantity
    ws.Cells(targetRow, "J").Value = strike
    ws.Cells(targetRow, "K").Value = CLng(maturityDays)

    If hasPremium Then
        ws.Cells(targetRow, "L").Value = premium
    Else
        ws.Cells(targetRow, "L").ClearContents
    End If

    Portfolio.InvalidateResults
    RefreshList
End Sub

Private Sub cmdDelete_Click()
    If lstPositions.ListIndex < 0 Then
        MsgBox "Sélectionnez d'abord une position.", _
               vbInformation, "Aucune sélection"
        Exit Sub
    End If

    Dim selectedId As Long
    selectedId = CLng(lstPositions.List(lstPositions.ListIndex, 0))

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Portfolio.SHEET_PORT)

    Dim rowIndex As Long
    Dim lastRow As Long
    Dim found As Boolean
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    For rowIndex = lastRow To 4 Step -1
        If ws.Cells(rowIndex, "E").Value2 = selectedId Then
            ' Ne jamais supprimer toute la ligne : les paramètres de marché
            ' se trouvent dans les colonnes A:C de la même feuille.
            ws.Range("E" & rowIndex & ":L" & rowIndex).Delete Shift:=xlUp
            found = True
            Exit For
        End If
    Next rowIndex

    If Not found Then
        MsgBox "La position sélectionnée n'a pas été trouvée.", _
               vbExclamation, "Suppression impossible"
    Else
        Portfolio.InvalidateResults
    End If

    RefreshList
End Sub

Private Sub RefreshList()
    lstPositions.Clear
    lstPositions.ColumnCount = 8
    lstPositions.ColumnWidths = "24;60;36;42;42;48;48;48"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Portfolio.SHEET_PORT)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow < 4 Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = 4 To lastRow
        If Len(Trim$(CStr(ws.Cells(rowIndex, "E").Value2))) > 0 Then
            lstPositions.AddItem ws.Cells(rowIndex, "E").Value
            lstPositions.List(lstPositions.ListCount - 1, 1) = ws.Cells(rowIndex, "F").Value
            lstPositions.List(lstPositions.ListCount - 1, 2) = ws.Cells(rowIndex, "G").Value
            lstPositions.List(lstPositions.ListCount - 1, 3) = ws.Cells(rowIndex, "H").Value
            lstPositions.List(lstPositions.ListCount - 1, 4) = ws.Cells(rowIndex, "I").Value
            lstPositions.List(lstPositions.ListCount - 1, 5) = ws.Cells(rowIndex, "J").Value
            lstPositions.List(lstPositions.ListCount - 1, 6) = ws.Cells(rowIndex, "K").Value
            lstPositions.List(lstPositions.ListCount - 1, 7) = ws.Cells(rowIndex, "L").Value
        End If
    Next rowIndex
End Sub

Private Function TryParseDouble(ByVal rawText As String, _
                                ByRef numberValue As Double) As Boolean
    Dim normalized As String
    normalized = Trim$(rawText)
    If Len(normalized) = 0 Then Exit Function

    If Application.DecimalSeparator = "," Then
        normalized = Replace(normalized, ".", ",")
    Else
        normalized = Replace(normalized, ",", ".")
    End If

    If Not IsNumeric(normalized) Then Exit Function

    numberValue = CDbl(normalized)
    TryParseDouble = True
End Function
