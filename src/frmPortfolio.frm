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

' --- UTILS ---
Private Function ToDbl(ByVal v As Variant) As Double
    ToDbl = Val(Replace(CStr(v), ",", "."))
End Function

Private Sub UserForm_Initialize()
    ' Style Flat / Clean
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
    
    ' Defaults
    txtUnderlying.Text = "Asset"
    cboType.Clear: cboType.AddItem "CALL": cboType.AddItem "PUT": cboType.Value = "CALL"
    cboSide.Clear: cboSide.AddItem "Achat": cboSide.AddItem "Vente": cboSide.Value = "Achat"
    txtQty.Text = "1"
    txtK.Text = "100"
    txtMatDays.Text = "30"
    txtPrime.Text = "0"
    
    RefreshList
End Sub

Private Sub cmdAdd_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Portfolio")
    If ToDbl(txtK.Text) <= 0 Then MsgBox "Strike invalide": Exit Sub
    
    Dim row&: row = ws.Cells(ws.Rows.Count, "E").End(xlUp).row + 1
    If row < 4 Then row = 4
    
    Dim id&: id = 1
    If row > 4 Then id = Application.Max(ws.Range("E4:E" & row - 1)) + 1
    
    ws.Cells(row, "E").Value = id
    ws.Cells(row, "F").Value = txtUnderlying.Text
    ws.Cells(row, "G").Value = cboType.Value
    ws.Cells(row, "H").Value = cboSide.Value
    ws.Cells(row, "I").Value = ToDbl(txtQty.Text)
    ws.Cells(row, "J").Value = ToDbl(txtK.Text)
    ws.Cells(row, "K").Value = ToDbl(txtMatDays.Text)
    If ToDbl(txtPrime.Text) > 0 Then ws.Cells(row, "L").Value = ToDbl(txtPrime.Text) Else ws.Cells(row, "L").Value = ""
    
    RefreshList
End Sub

Private Sub cmdDelete_Click()
    If lstPositions.ListIndex < 0 Then Exit Sub
    Dim idSel&: idSel = CLng(lstPositions.List(lstPositions.ListIndex, 0))
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Portfolio")
    Dim i&, last&
    last = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    For i = last To 4 Step -1
        If ws.Cells(i, "E").Value = idSel Then ws.Rows(i).Delete: Exit For
    Next i
    RefreshList
End Sub

Private Sub RefreshList()
    lstPositions.Clear
    lstPositions.ColumnCount = 8
    lstPositions.ColumnWidths = "20;40;30;30;30;30;30;30"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Portfolio")
    Dim last&: last = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    If last < 4 Then Exit Sub
    
    Dim i&
    For i = 4 To last
        lstPositions.AddItem ws.Cells(i, "E").Value
        lstPositions.List(lstPositions.ListCount - 1, 1) = ws.Cells(i, "F").Value
        lstPositions.List(lstPositions.ListCount - 1, 2) = ws.Cells(i, "G").Value
        lstPositions.List(lstPositions.ListCount - 1, 3) = ws.Cells(i, "H").Value
        lstPositions.List(lstPositions.ListCount - 1, 4) = ws.Cells(i, "I").Value
        lstPositions.List(lstPositions.ListCount - 1, 5) = ws.Cells(i, "J").Value
        lstPositions.List(lstPositions.ListCount - 1, 6) = ws.Cells(i, "K").Value
        lstPositions.List(lstPositions.ListCount - 1, 7) = ws.Cells(i, "L").Value
    Next i
End Sub
