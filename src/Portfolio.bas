Attribute VB_Name = "Portfolio"
Option Explicit

' ===== Constantes =====
Public Const SHEET_PORT As String = "Portfolio"
Public Const SHEET_GRID As String = "GridPort"
Public Const SHEET_GREEKS As String = "Greeks" ' Nouvelle feuille

' =========================================================================
' 1) MATHS (BLACK-SCHOLES & UTILITAIRES)
' =========================================================================

' --- Fonctions de Loi Normale (Utilise les fonctions Excel natives pour la précision) ---
Private Function Ncdf(ByVal x As Double) As Double
    Ncdf = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function Npdf(ByVal x As Double) As Double
    Npdf = Application.WorksheetFunction.Norm_S_Dist(x, False)
End Function

' --- Calculs intermédiaires d1 et d2 ---
Private Function d1_(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                     ByVal sig As Double, ByVal T As Double) As Double
    If S <= 0 Or K <= 0 Or sig <= 0 Or T <= 0 Then
        d1_ = 0
        Exit Function
    End If
    d1_ = (Log(S / K) + (r + 0.5 * sig ^ 2) * T) / (sig * Sqr(T))
End Function

Private Function d2_(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                     ByVal sig As Double, ByVal T As Double) As Double
    If S <= 0 Or K <= 0 Or sig <= 0 Or T <= 0 Then
        d2_ = 0
        Exit Function
    End If
    d2_ = d1_(S, K, r, sig, T) - sig * Sqr(T)
End Function

' --- Formules de Pricing (Call & Put) ---
Public Function BS_Call(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                        ByVal sig As Double, ByVal T As Double) As Double
    ' Cas particuliers pour éviter les erreurs de division par zéro
    If T <= 0 Then
        BS_Call = Application.Max(S - K, 0)
        Exit Function
    End If
    If sig <= 0 Then
        BS_Call = Application.Max(S - K * Exp(-r * T), 0)
        Exit Function
    End If
    
    Dim d1 As Double, d2 As Double
    d1 = d1_(S, K, r, sig, T)
    d2 = d2_(S, K, r, sig, T)
    
    BS_Call = S * Ncdf(d1) - K * Exp(-r * T) * Ncdf(d2)
End Function

Public Function BS_Put(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                       ByVal sig As Double, ByVal T As Double) As Double
    ' Cas particuliers
    If T <= 0 Then
        BS_Put = Application.Max(K - S, 0)
        Exit Function
    End If
    If sig <= 0 Then
        BS_Put = Application.Max(K * Exp(-r * T) - S, 0)
        Exit Function
    End If

    Dim d1 As Double, d2 As Double
    d1 = d1_(S, K, r, sig, T)
    d2 = d2_(S, K, r, sig, T)
    
    BS_Put = K * Exp(-r * T) * Ncdf(-d2) - S * Ncdf(-d1)
End Function
' Grecs
Public Function Delta_Call(S, K, r, sig, T) As Double: Delta_Call = Ncdf(d1_(S, K, r, sig, T)): End Function
Public Function Delta_Put(S, K, r, sig, T) As Double: Delta_Put = Ncdf(d1_(S, K, r, sig, T)) - 1: End Function
Public Function Gamma_(S, K, r, sig, T) As Double: Gamma_ = Npdf(d1_(S, K, r, sig, T)) / (S * sig * Sqr(T)): End Function
Public Function Vega_(S, K, r, sig, T) As Double: Vega_ = S * Npdf(d1_(S, K, r, sig, T)) * Sqr(T): End Function
Public Function Theta_Call(S, K, r, sig, T) As Double
    Dim d1 As Double, d2 As Double: d1 = d1_(S, K, r, sig, T): d2 = d2_(S, K, r, sig, T)
    Theta_Call = -(S * Npdf(d1) * sig) / (2 * Sqr(T)) - r * K * Exp(-r * T) * Ncdf(d2)
End Function
Public Function Theta_Put(S, K, r, sig, T) As Double
    Dim d1 As Double, d2 As Double: d1 = d1_(S, K, r, sig, T): d2 = d2_(S, K, r, sig, T)
    Theta_Put = -(S * Npdf(d1) * sig) / (2 * Sqr(T)) + r * K * Exp(-r * T) * Ncdf(-d2)
End Function
Public Function Rho_Call(S, K, r, sig, T) As Double: Rho_Call = K * T * Exp(-r * T) * Ncdf(d2_(S, K, r, sig, T)): End Function
Public Function Rho_Put(S, K, r, sig, T) As Double: Rho_Put = -K * T * Exp(-r * T) * Ncdf(-d2_(S, K, r, sig, T)): End Function

' =========================================================================
' 2) GESTION & CALCUL
' =========================================================================
Public Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.name = sheetName
    End If
    On Error GoTo 0
End Function

Private Function ReadPortfolio(S0, r, ByRef outSigma) As Collection
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(SHEET_PORT)
    Dim col As New Collection, last&, i&
    outSigma = ws.Range("B11").Value: If outSigma <= 0 Then outSigma = 0.25
    last = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    If last < 4 Then Set ReadPortfolio = col: Exit Function
    
    For i = 4 To last
        If Len(ws.Cells(i, "E").Value) > 0 Then
            Dim p As Classe1: Set p = New Classe1
            p.Label = CStr(ws.Cells(i, "F").Value)
            p.isCall = (UCase(ws.Cells(i, "G").Value) Like "*CALL*")
            p.Sign = IIf(UCase(ws.Cells(i, "H").Value) Like "*VENTE*", -1, 1)
            p.qty = Val(Replace(ws.Cells(i, "I").Value, ",", "."))
            p.K = Val(Replace(ws.Cells(i, "J").Value, ",", "."))
            Dim rawMat#: rawMat = Val(Replace(ws.Cells(i, "K").Value, ",", "."))
            p.T = IIf(rawMat > 10, rawMat / 365, rawMat)
            p.Sigma = outSigma
            Dim vPr: vPr = ws.Cells(i, "L").Value
            If Val(vPr) = 0 Then
                If p.isCall Then p.Premium0 = BS_Call(S0, p.K, r, p.Sigma, p.T) Else p.Premium0 = BS_Put(S0, p.K, r, p.Sigma, p.T)
            Else
                p.Premium0 = Val(Replace(vPr, ",", "."))
            End If
            If p.qty > 0 Then col.Add p
        End If
    Next i
    Set ReadPortfolio = col
End Function

Public Sub Port_BuildGrid()
    Dim wsP As Worksheet: Set wsP = GetOrCreateSheet(SHEET_PORT)
    Dim S0#, r#, Smin#, Smax#, stepS#, sig#
    On Error Resume Next
    S0 = wsP.Range("B5").Value: r = wsP.Range("B8").Value: sig = wsP.Range("B11").Value
    Smin = wsP.Range("B17").Value: Smax = wsP.Range("B20").Value: stepS = wsP.Range("B23").Value
    If stepS <= 0 Then stepS = 1
    On Error GoTo 0
    
    Dim positions As Collection: Set positions = ReadPortfolio(S0, r, sig)
    If positions.Count = 0 Then MsgBox "Portefeuille vide.", vbExclamation: Exit Sub

    Dim wsG As Worksheet: Set wsG = GetOrCreateSheet(SHEET_GRID)
    wsG.Cells.Clear
    wsG.Range("A1:I1").Value = Array("S", "Value", "Delta", "Gamma", "Vega", "Theta", "Rho", "PnL_T", "PnL_MTM")
    
    Dim row&, S_curr#, i&
    row = 2
    Application.ScreenUpdating = False
    
    For S_curr = Smin To Smax Step stepS
        Dim Ssafe#: Ssafe = Application.Max(S_curr, 0.01)
        Dim tVal#, tD#, tG#, tV#, tT#, tR#, tPnLT#, tPnLM#
        
        For i = 1 To positions.Count
            Dim p As Classe1: Set p = positions(i)
            Dim pr#, d#, g#, v#, th#, rh#, pay#
            If p.isCall Then
                pr = BS_Call(Ssafe, p.K, r, p.Sigma, p.T): d = Delta_Call(Ssafe, p.K, r, p.Sigma, p.T)
                g = Gamma_(Ssafe, p.K, r, p.Sigma, p.T): v = Vega_(Ssafe, p.K, r, p.Sigma, p.T)
                th = Theta_Call(Ssafe, p.K, r, p.Sigma, p.T): rh = Rho_Call(Ssafe, p.K, r, p.Sigma, p.T)
                pay = Application.Max(Ssafe - p.K, 0)
            Else
                pr = BS_Put(Ssafe, p.K, r, p.Sigma, p.T): d = Delta_Put(Ssafe, p.K, r, p.Sigma, p.T)
                g = Gamma_(Ssafe, p.K, r, p.Sigma, p.T): v = Vega_(Ssafe, p.K, r, p.Sigma, p.T)
                th = Theta_Put(Ssafe, p.K, r, p.Sigma, p.T): rh = Rho_Put(Ssafe, p.K, r, p.Sigma, p.T)
                pay = Application.Max(p.K - Ssafe, 0)
            End If
            
            Dim f#: f = p.Sign * p.qty
            tVal = tVal + f * pr: tD = tD + f * d: tG = tG + f * g
            tV = tV + f * (v / 100): tT = tT + f * (th / 365): tR = tR + f * (rh / 100)
            tPnLT = tPnLT + f * (pay - p.Premium0): tPnLM = tPnLM + f * (pr - p.Premium0)
        Next i
        
        wsG.Cells(row, 1).Value = Ssafe: wsG.Cells(row, 2).Value = tVal
        wsG.Cells(row, 3).Value = tD: wsG.Cells(row, 4).Value = tG
        wsG.Cells(row, 5).Value = tV: wsG.Cells(row, 6).Value = tT
        wsG.Cells(row, 7).Value = tR: wsG.Cells(row, 8).Value = tPnLT
        wsG.Cells(row, 9).Value = tPnLM
        row = row + 1
    Next S_curr
    Application.ScreenUpdating = True
End Sub

' =========================================================================
' 3) GRAPHIQUES (SEPARÉS)
' =========================================================================
Public Sub Port_PlotCharts_All()
    On Error GoTo FAIL
    Dim wsG As Worksheet: Set wsG = GetOrCreateSheet(SHEET_GRID)
    Dim lastRow&: lastRow = wsG.Cells(wsG.Rows.Count, "A").End(xlUp).row
    If lastRow < 2 Then MsgBox "Aucune donnée. Cliquez sur Calculer.", vbExclamation: Exit Sub
    
    Dim rngX As Range: Set rngX = wsG.Range("A2:A" & lastRow)
    
    ' --- 1. GRAPHIQUE PNL SUR FEUILLE PORTFOLIO (A DROITE DU TABLEAU) ---
    Dim wsP As Worksheet: Set wsP = GetOrCreateSheet(SHEET_PORT)
    wsP.Activate
    DeleteCharts wsP
    
    ' Positionnement en colonne M (à droite du tableau)
    Dim ch As ChartObject
    Set ch = wsP.ChartObjects.Add(Left:=wsP.Range("M3").Left, Top:=wsP.Range("M3").Top, Width:=600, Height:=350)
    With ch.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = True: .ChartTitle.Text = "Profil de PnL (Gains/Pertes)"
        
        ' Série MTM (Aujourd'hui)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = rngX
        .SeriesCollection(1).Values = wsG.Range("I2:I" & lastRow)
        .SeriesCollection(1).name = "MTM (Auj.)"
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192) ' Bleu
        .SeriesCollection(1).Format.Line.Weight = 2
        
        ' Série Echéance
        .SeriesCollection.NewSeries
        .SeriesCollection(2).XValues = rngX
        .SeriesCollection(2).Values = wsG.Range("H2:H" & lastRow)
        .SeriesCollection(2).name = "Echéance"
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' Rouge
        .SeriesCollection(2).Format.Line.DashStyle = msoLineDash
        
        .Legend.Position = xlLegendPositionBottom
    End With
    
    ' --- 2. GRAPHIQUES GRECS SUR FEUILLE "GREEKS" ---
    Dim wsGr As Worksheet: Set wsGr = GetOrCreateSheet(SHEET_GREEKS)
    wsGr.Activate
    wsGr.Cells.Interior.Color = RGB(250, 250, 250) ' Fond propre
    DeleteCharts wsGr
    
    ' Titre de la page
    wsGr.Range("B2").Value = "ANALYSES DE SENSIBILITÉ (GRECS)"
    wsGr.Range("B2").Font.Size = 16: wsGr.Range("B2").Font.Bold = True
    wsGr.Range("B2").Font.Color = RGB(45, 62, 80)
    
    ' Création des 5 graphes en grille
    Dim configs, i&
    configs = Array( _
        Array("Delta", 3, 0, 0), _
        Array("Gamma", 4, 1, 0), _
        Array("Vega", 5, 0, 1), _
        Array("Theta", 6, 1, 1), _
        Array("Rho", 7, 0, 2) _
    )
    
    Dim startLeft#: startLeft = 20
    Dim startTop#: startTop = 50
    Dim w#: w = 450: Dim h#: h = 250
    Dim gap#: gap = 20
    
    For i = LBound(configs) To UBound(configs)
        Dim conf: conf = configs(i)
        Dim lPos#: lPos = startLeft + (conf(2) * (w + gap))
        Dim tPos#: tPos = startTop + (conf(3) * (h + gap))
        
        Set ch = wsGr.ChartObjects.Add(Left:=lPos, Top:=tPos, Width:=w, Height:=h)
        With ch.Chart
            .ChartType = xlXYScatterLinesNoMarkers
            .HasTitle = True: .ChartTitle.Text = "Portefeuille " & conf(0)
            .SeriesCollection.NewSeries
            .SeriesCollection(1).XValues = rngX
            .SeriesCollection(1).Values = wsG.Range(wsG.Cells(2, conf(1)), wsG.Cells(lastRow, conf(1)))
            .Legend.Delete
            .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
            .ChartArea.Format.Line.Visible = msoFalse
        End With
    Next i
    
    ' Retour au portfolio
    wsP.Activate
    MsgBox "Graphiques mis à jour !\n- PnL : à droite.\n- Grecs : voir feuille 'Greeks'.", vbInformation
    Exit Sub
FAIL:
    MsgBox "Erreur Graphes : " & Err.Description
End Sub

Private Sub DeleteCharts(ws As Worksheet)
    Dim co As ChartObject
    For Each co In ws.ChartObjects: co.Delete: Next co
End Sub

Public Sub Ouvrir_Portfolio()
    frmPortfolio.Show
End Sub

Sub LanceTout()
    Portfolio.Port_BuildGrid
    Portfolio.Port_PlotCharts_All
End Sub

