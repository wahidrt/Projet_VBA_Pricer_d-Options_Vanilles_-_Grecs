Attribute VB_Name = "Portfolio"
Option Explicit

' =========================================================================
' PRICER BLACK-SCHOLES ET PORTEFEUILLE MONO-SOUS-JACENT
' =========================================================================

Public Const SHEET_PORT As String = "Portfolio"
Public Const SHEET_GRID As String = "GridPort"
Public Const SHEET_GREEKS As String = "Greeks"

Private Const CELL_SPOT As String = "B5"
Private Const CELL_RATE As String = "B8"
Private Const CELL_VOLATILITY As String = "B11"
Private Const CELL_DIVIDEND As String = "B13"
Private Const CELL_S_MIN As String = "B17"
Private Const CELL_S_MAX As String = "B20"
Private Const CELL_S_STEP As String = "B23"

Private Const FIRST_POSITION_ROW As Long = 4
Private Const MAX_GRID_POINTS As Long = 2000
Private Const DAYS_IN_YEAR As Double = 365#
Private Const EPSILON As Double = 0.0000001

Private Const CHART_PNL As String = "pricer_PnL"
Private Const CHART_DELTA As String = "pricer_Delta"
Private Const CHART_GAMMA As String = "pricer_Gamma"
Private Const CHART_VEGA As String = "pricer_Vega"
Private Const CHART_THETA As String = "pricer_Theta"
Private Const CHART_RHO As String = "pricer_Rho"

' =========================================================================
' 1) MOTEUR BLACK-SCHOLES
' Hypothèses : options européennes, taux/volatilité/dividende constants.
' =========================================================================

Private Function Ncdf(ByVal x As Double) As Double
    Ncdf = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function Npdf(ByVal x As Double) As Double
    Npdf = Application.WorksheetFunction.Norm_S_Dist(x, False)
End Function

Private Sub ValidatePriceInputs(ByVal S As Double, ByVal K As Double, _
                                ByVal sigma As Double, ByVal T As Double)
    If S <= 0# Then
        Err.Raise vbObjectError + 1001, "BlackScholes", _
                  "Le spot doit être strictement positif."
    End If
    If K <= 0# Then
        Err.Raise vbObjectError + 1002, "BlackScholes", _
                  "Le strike doit être strictement positif."
    End If
    If sigma < 0# Then
        Err.Raise vbObjectError + 1003, "BlackScholes", _
                  "La volatilité ne peut pas être négative."
    End If
    If T < 0# Then
        Err.Raise vbObjectError + 1004, "BlackScholes", _
                  "La maturité ne peut pas être négative."
    End If
End Sub

Private Sub ValidateGreekInputs(ByVal S As Double, ByVal K As Double, _
                                ByVal sigma As Double, ByVal T As Double)
    ValidatePriceInputs S, K, sigma, T
    If sigma <= 0# Then
        Err.Raise vbObjectError + 1005, "BlackScholes", _
                  "Les Grecs exigent une volatilité strictement positive."
    End If
    If T <= 0# Then
        Err.Raise vbObjectError + 1006, "BlackScholes", _
                  "Les Grecs exigent une maturité strictement positive."
    End If
End Sub

Private Function d1_(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                     ByVal sigma As Double, ByVal T As Double, _
                     ByVal q As Double) As Double
    d1_ = (Log(S / K) + (r - q + 0.5 * sigma ^ 2) * T) / _
          (sigma * Sqr(T))
End Function

Private Function d2_(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                     ByVal sigma As Double, ByVal T As Double, _
                     ByVal q As Double) As Double
    d2_ = d1_(S, K, r, sigma, T, q) - sigma * Sqr(T)
End Function

Public Function BS_Call(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                        ByVal sigma As Double, ByVal T As Double, _
                        Optional ByVal q As Double = 0#) As Double
    ValidatePriceInputs S, K, sigma, T

    If T = 0# Then
        BS_Call = Application.Max(S - K, 0#)
        Exit Function
    End If

    If sigma = 0# Then
        BS_Call = Application.Max(S * Exp(-q * T) - K * Exp(-r * T), 0#)
        Exit Function
    End If

    Dim d1 As Double
    Dim d2 As Double
    d1 = d1_(S, K, r, sigma, T, q)
    d2 = d2_(S, K, r, sigma, T, q)

    BS_Call = S * Exp(-q * T) * Ncdf(d1) - _
              K * Exp(-r * T) * Ncdf(d2)
End Function

Public Function BS_Put(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                       ByVal sigma As Double, ByVal T As Double, _
                       Optional ByVal q As Double = 0#) As Double
    ValidatePriceInputs S, K, sigma, T

    If T = 0# Then
        BS_Put = Application.Max(K - S, 0#)
        Exit Function
    End If

    If sigma = 0# Then
        BS_Put = Application.Max(K * Exp(-r * T) - S * Exp(-q * T), 0#)
        Exit Function
    End If

    Dim d1 As Double
    Dim d2 As Double
    d1 = d1_(S, K, r, sigma, T, q)
    d2 = d2_(S, K, r, sigma, T, q)

    BS_Put = K * Exp(-r * T) * Ncdf(-d2) - _
             S * Exp(-q * T) * Ncdf(-d1)
End Function

Public Function Delta_Call(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                           ByVal sigma As Double, ByVal T As Double, _
                           Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T
    Delta_Call = Exp(-q * T) * Ncdf(d1_(S, K, r, sigma, T, q))
End Function

Public Function Delta_Put(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                          ByVal sigma As Double, ByVal T As Double, _
                          Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T
    Delta_Put = Exp(-q * T) * (Ncdf(d1_(S, K, r, sigma, T, q)) - 1#)
End Function

Public Function Gamma_(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                       ByVal sigma As Double, ByVal T As Double, _
                       Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T
    Gamma_ = Exp(-q * T) * Npdf(d1_(S, K, r, sigma, T, q)) / _
             (S * sigma * Sqr(T))
End Function

Public Function Vega_(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                      ByVal sigma As Double, ByVal T As Double, _
                      Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T
    Vega_ = S * Exp(-q * T) * Npdf(d1_(S, K, r, sigma, T, q)) * Sqr(T)
End Function

Public Function Theta_Call(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                           ByVal sigma As Double, ByVal T As Double, _
                           Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T

    Dim d1 As Double
    Dim d2 As Double
    d1 = d1_(S, K, r, sigma, T, q)
    d2 = d2_(S, K, r, sigma, T, q)

    Theta_Call = -(S * Exp(-q * T) * Npdf(d1) * sigma) / (2# * Sqr(T)) _
                 - r * K * Exp(-r * T) * Ncdf(d2) _
                 + q * S * Exp(-q * T) * Ncdf(d1)
End Function

Public Function Theta_Put(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                          ByVal sigma As Double, ByVal T As Double, _
                          Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T

    Dim d1 As Double
    Dim d2 As Double
    d1 = d1_(S, K, r, sigma, T, q)
    d2 = d2_(S, K, r, sigma, T, q)

    Theta_Put = -(S * Exp(-q * T) * Npdf(d1) * sigma) / (2# * Sqr(T)) _
                + r * K * Exp(-r * T) * Ncdf(-d2) _
                - q * S * Exp(-q * T) * Ncdf(-d1)
End Function

Public Function Rho_Call(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                         ByVal sigma As Double, ByVal T As Double, _
                         Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T
    Rho_Call = K * T * Exp(-r * T) * Ncdf(d2_(S, K, r, sigma, T, q))
End Function

Public Function Rho_Put(ByVal S As Double, ByVal K As Double, ByVal r As Double, _
                        ByVal sigma As Double, ByVal T As Double, _
                        Optional ByVal q As Double = 0#) As Double
    ValidateGreekInputs S, K, sigma, T
    Rho_Put = -K * T * Exp(-r * T) * Ncdf(-d2_(S, K, r, sigma, T, q))
End Function

' =========================================================================
' 2) LECTURE ET VALIDATION DES DONNÉES
' =========================================================================

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws
End Function

Private Function TryConvertNumber(ByVal rawValue As Variant, ByVal fieldName As String, _
                                  ByRef numberValue As Double, _
                                  ByRef errorMessage As String) As Boolean
    If IsError(rawValue) Then
        errorMessage = fieldName & " contient une erreur Excel."
        Exit Function
    End If

    If Len(Trim$(CStr(rawValue))) = 0 Then
        errorMessage = fieldName & " est obligatoire."
        Exit Function
    End If

    If Not IsNumeric(rawValue) Then
        errorMessage = fieldName & " doit être numérique."
        Exit Function
    End If

    numberValue = CDbl(rawValue)
    TryConvertNumber = True
End Function

Private Function TryReadCellNumber(ByVal ws As Worksheet, ByVal address As String, _
                                   ByVal fieldName As String, _
                                   ByRef numberValue As Double, _
                                   ByRef errorMessage As String, _
                                   Optional ByVal allowBlank As Boolean = False, _
                                   Optional ByVal blankValue As Double = 0#) As Boolean
    Dim rawValue As Variant
    rawValue = ws.Range(address).Value2

    If Not IsError(rawValue) Then
        If Len(Trim$(CStr(rawValue))) = 0 And allowBlank Then
            numberValue = blankValue
            TryReadCellNumber = True
            Exit Function
        End If
    End If

    TryReadCellNumber = TryConvertNumber(rawValue, fieldName, numberValue, errorMessage)
End Function

Private Function ReadScenarioInputs(ByVal ws As Worksheet, _
                                    ByRef spot As Double, ByRef rate As Double, _
                                    ByRef sigma As Double, ByRef dividendYield As Double, _
                                    ByRef sMin As Double, ByRef sMax As Double, _
                                    ByRef stepS As Double, ByRef pointCount As Long, _
                                    ByRef errorMessage As String) As Boolean
    If Not TryReadCellNumber(ws, CELL_SPOT, "Le prix spot", spot, errorMessage) Then Exit Function
    If Not TryReadCellNumber(ws, CELL_RATE, "Le taux sans risque", rate, errorMessage) Then Exit Function
    If Not TryReadCellNumber(ws, CELL_VOLATILITY, "La volatilité", sigma, errorMessage) Then Exit Function
    If Not TryReadCellNumber(ws, CELL_DIVIDEND, "Le rendement de dividende", _
                             dividendYield, errorMessage, True, 0#) Then Exit Function
    If Not TryReadCellNumber(ws, CELL_S_MIN, "S min", sMin, errorMessage) Then Exit Function
    If Not TryReadCellNumber(ws, CELL_S_MAX, "S max", sMax, errorMessage) Then Exit Function
    If Not TryReadCellNumber(ws, CELL_S_STEP, "Le pas de simulation", stepS, errorMessage) Then Exit Function

    If spot <= 0# Then
        errorMessage = "Le prix spot doit être strictement positif."
        Exit Function
    End If
    If sigma <= 0# Or sigma > 5# Then
        errorMessage = "La volatilité doit être comprise entre 0 et 500 %."
        Exit Function
    End If
    If rate < -1# Or rate > 1# Then
        errorMessage = "Le taux sans risque doit être saisi sous forme décimale."
        Exit Function
    End If
    If dividendYield < -1# Or dividendYield > 1# Then
        errorMessage = "Le rendement de dividende doit être saisi sous forme décimale."
        Exit Function
    End If
    If sMin <= 0# Then
        errorMessage = "S min doit être strictement positif."
        Exit Function
    End If
    If sMax <= sMin Then
        errorMessage = "S max doit être strictement supérieur à S min."
        Exit Function
    End If
    If stepS <= 0# Then
        errorMessage = "Le pas de simulation doit être strictement positif."
        Exit Function
    End If

    pointCount = CLng(Int((sMax - sMin) / stepS)) + 1
    If pointCount < 2 Then
        errorMessage = "La simulation doit contenir au moins deux points."
        Exit Function
    End If
    If pointCount > MAX_GRID_POINTS Then
        errorMessage = "La simulation contient " & pointCount & _
                       " points. Augmentez le pas pour rester sous " & _
                       MAX_GRID_POINTS & " points."
        Exit Function
    End If

    ReadScenarioInputs = True
End Function

Private Function LoadPositions(ByVal ws As Worksheet, ByVal spot As Double, _
                               ByVal rate As Double, ByVal sigma As Double, _
                               ByVal dividendYield As Double, _
                               ByRef errorMessage As String) As Collection
    On Error GoTo LOAD_ERROR

    Dim positions As New Collection
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim firstUnderlying As String
    Dim firstMaturityDays As Double

    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow < FIRST_POSITION_ROW Then
        errorMessage = "Le portefeuille est vide."
        Exit Function
    End If

    For rowIndex = FIRST_POSITION_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(rowIndex, "E").Value2))) > 0 Then
            Dim position As Classe1
            Dim underlying As String
            Dim optionType As String
            Dim positionSide As String
            Dim quantity As Double
            Dim strike As Double
            Dim maturityDays As Double
            Dim premium As Double
            Dim premiumRaw As Variant

            Set position = New Classe1
            underlying = Trim$(CStr(ws.Cells(rowIndex, "F").Value2))
            optionType = UCase$(Trim$(CStr(ws.Cells(rowIndex, "G").Value2)))
            positionSide = UCase$(Trim$(CStr(ws.Cells(rowIndex, "H").Value2)))

            If Len(underlying) = 0 Then
                errorMessage = "Sous-jacent manquant à la ligne " & rowIndex & "."
                Exit Function
            End If

            If Len(firstUnderlying) = 0 Then
                firstUnderlying = underlying
            ElseIf StrComp(firstUnderlying, underlying, vbTextCompare) <> 0 Then
                errorMessage = "Le modèle utilise un seul sous-jacent. " & _
                               "La ligne " & rowIndex & " contient """ & _
                               underlying & """ au lieu de """ & firstUnderlying & """."
                Exit Function
            End If

            Select Case optionType
                Case "CALL"
                    position.IsCall = True
                Case "PUT"
                    position.IsCall = False
                Case Else
                    errorMessage = "Le type de la ligne " & rowIndex & _
                                   " doit être CALL ou PUT."
                    Exit Function
            End Select

            Select Case positionSide
                Case "ACHAT"
                    position.PositionSign = 1#
                Case "VENTE"
                    position.PositionSign = -1#
                Case Else
                    errorMessage = "La position de la ligne " & rowIndex & _
                                   " doit être Achat ou Vente."
                    Exit Function
            End Select

            If Not TryConvertNumber(ws.Cells(rowIndex, "I").Value2, _
                                    "La quantité de la ligne " & rowIndex, _
                                    quantity, errorMessage) Then Exit Function
            If Not TryConvertNumber(ws.Cells(rowIndex, "J").Value2, _
                                    "Le strike de la ligne " & rowIndex, _
                                    strike, errorMessage) Then Exit Function
            If Not TryConvertNumber(ws.Cells(rowIndex, "K").Value2, _
                                    "La maturité de la ligne " & rowIndex, _
                                    maturityDays, errorMessage) Then Exit Function

            If quantity <= 0# Then
                errorMessage = "La quantité de la ligne " & rowIndex & _
                               " doit être strictement positive."
                Exit Function
            End If
            If strike <= 0# Then
                errorMessage = "Le strike de la ligne " & rowIndex & _
                               " doit être strictement positif."
                Exit Function
            End If
            If maturityDays <= 0# Then
                errorMessage = "La maturité de la ligne " & rowIndex & _
                               " doit être strictement positive."
                Exit Function
            End If
            If Abs(maturityDays - Fix(maturityDays)) > EPSILON Then
                errorMessage = "La maturité de la ligne " & rowIndex & _
                               " doit être un nombre entier de jours."
                Exit Function
            End If

            If firstMaturityDays = 0# Then
                firstMaturityDays = maturityDays
            ElseIf Abs(firstMaturityDays - maturityDays) > EPSILON Then
                errorMessage = "Le profil de PnL à l'échéance exige une maturité " & _
                               "commune. Vérifiez la ligne " & rowIndex & "."
                Exit Function
            End If

            position.Underlying = underlying
            position.Quantity = quantity
            position.Strike = strike
            position.MaturityYears = maturityDays / DAYS_IN_YEAR
            position.Volatility = sigma

            premiumRaw = ws.Cells(rowIndex, "L").Value2
            If IsError(premiumRaw) Then
                errorMessage = "La prime de la ligne " & rowIndex & _
                               " contient une erreur Excel."
                Exit Function
            ElseIf Len(Trim$(CStr(premiumRaw))) = 0 Then
                If position.IsCall Then
                    premium = BS_Call(spot, strike, rate, sigma, _
                                      position.MaturityYears, dividendYield)
                Else
                    premium = BS_Put(spot, strike, rate, sigma, _
                                     position.MaturityYears, dividendYield)
                End If
            Else
                If Not TryConvertNumber(premiumRaw, _
                                        "La prime de la ligne " & rowIndex, _
                                        premium, errorMessage) Then Exit Function
                If premium < 0# Then
                    errorMessage = "La prime de la ligne " & rowIndex & _
                                   " ne peut pas être négative."
                    Exit Function
                End If
            End If

            position.InitialPremium = premium
            positions.Add position
        End If
    Next rowIndex

    If positions.Count = 0 Then
        errorMessage = "Le portefeuille est vide."
        Exit Function
    End If

    Set LoadPositions = positions
    Exit Function

LOAD_ERROR:
    errorMessage = "Erreur pendant la lecture du portefeuille : " & Err.Description
End Function

' =========================================================================
' 3) CONSTRUCTION DE LA GRILLE
' =========================================================================

Public Function Port_BuildGrid() As Boolean
    Dim previousScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo BUILD_ERROR

    Dim wsPortfolio As Worksheet
    Dim wsGrid As Worksheet
    Dim positions As Collection
    Dim errorMessage As String
    Dim spot As Double
    Dim rate As Double
    Dim sigma As Double
    Dim dividendYield As Double
    Dim sMin As Double
    Dim sMax As Double
    Dim stepS As Double
    Dim pointCount As Long

    Set wsPortfolio = GetOrCreateSheet(SHEET_PORT)

    If Not ReadScenarioInputs(wsPortfolio, spot, rate, sigma, dividendYield, _
                              sMin, sMax, stepS, pointCount, errorMessage) Then
        MsgBox errorMessage, vbExclamation, "Paramètres invalides"
        GoTo BUILD_EXIT
    End If

    Set positions = LoadPositions(wsPortfolio, spot, rate, sigma, _
                                  dividendYield, errorMessage)
    If positions Is Nothing Then
        MsgBox errorMessage, vbExclamation, "Portefeuille invalide"
        GoTo BUILD_EXIT
    End If

    Set wsGrid = GetOrCreateSheet(SHEET_GRID)
    Application.ScreenUpdating = False
    Application.StatusBar = "Calcul du portefeuille..."

    Dim results() As Variant
    ReDim results(1 To pointCount, 1 To 9)

    Dim pointIndex As Long
    Dim positionIndex As Long
    Dim currentSpot As Double
    Dim optionPrice As Double
    Dim delta As Double
    Dim gamma As Double
    Dim vega As Double
    Dim theta As Double
    Dim rho As Double
    Dim payoff As Double
    Dim exposure As Double
    Dim position As Classe1

    For pointIndex = 1 To pointCount
        currentSpot = sMin + (pointIndex - 1) * stepS

        Dim totalValue As Double
        Dim totalDelta As Double
        Dim totalGamma As Double
        Dim totalVega As Double
        Dim totalTheta As Double
        Dim totalRho As Double
        Dim totalExpiryPnl As Double
        Dim totalMtmPnl As Double

        ' Réinitialisation obligatoire pour chaque valeur du spot.
        totalValue = 0#
        totalDelta = 0#
        totalGamma = 0#
        totalVega = 0#
        totalTheta = 0#
        totalRho = 0#
        totalExpiryPnl = 0#
        totalMtmPnl = 0#

        For positionIndex = 1 To positions.Count
            Set position = positions(positionIndex)

            If position.IsCall Then
                optionPrice = BS_Call(currentSpot, position.Strike, rate, _
                                      position.Volatility, position.MaturityYears, _
                                      dividendYield)
                delta = Delta_Call(currentSpot, position.Strike, rate, _
                                   position.Volatility, position.MaturityYears, _
                                   dividendYield)
                theta = Theta_Call(currentSpot, position.Strike, rate, _
                                   position.Volatility, position.MaturityYears, _
                                   dividendYield)
                rho = Rho_Call(currentSpot, position.Strike, rate, _
                               position.Volatility, position.MaturityYears, _
                               dividendYield)
                payoff = Application.Max(currentSpot - position.Strike, 0#)
            Else
                optionPrice = BS_Put(currentSpot, position.Strike, rate, _
                                     position.Volatility, position.MaturityYears, _
                                     dividendYield)
                delta = Delta_Put(currentSpot, position.Strike, rate, _
                                  position.Volatility, position.MaturityYears, _
                                  dividendYield)
                theta = Theta_Put(currentSpot, position.Strike, rate, _
                                  position.Volatility, position.MaturityYears, _
                                  dividendYield)
                rho = Rho_Put(currentSpot, position.Strike, rate, _
                              position.Volatility, position.MaturityYears, _
                              dividendYield)
                payoff = Application.Max(position.Strike - currentSpot, 0#)
            End If

            gamma = Gamma_(currentSpot, position.Strike, rate, _
                           position.Volatility, position.MaturityYears, _
                           dividendYield)
            vega = Vega_(currentSpot, position.Strike, rate, _
                         position.Volatility, position.MaturityYears, _
                         dividendYield)

            exposure = position.PositionSign * position.Quantity
            totalValue = totalValue + exposure * optionPrice
            totalDelta = totalDelta + exposure * delta
            totalGamma = totalGamma + exposure * gamma
            totalVega = totalVega + exposure * (vega / 100#)
            totalTheta = totalTheta + exposure * (theta / DAYS_IN_YEAR)
            totalRho = totalRho + exposure * (rho / 100#)
            totalExpiryPnl = totalExpiryPnl + _
                             exposure * (payoff - position.InitialPremium)
            totalMtmPnl = totalMtmPnl + _
                          exposure * (optionPrice - position.InitialPremium)
        Next positionIndex

        results(pointIndex, 1) = currentSpot
        results(pointIndex, 2) = totalValue
        results(pointIndex, 3) = totalDelta
        results(pointIndex, 4) = totalGamma
        results(pointIndex, 5) = totalVega
        results(pointIndex, 6) = totalTheta
        results(pointIndex, 7) = totalRho
        results(pointIndex, 8) = totalExpiryPnl
        results(pointIndex, 9) = totalMtmPnl
    Next pointIndex

    With wsGrid
        .Cells.ClearContents
        .Range("A1:I1").Value = Array( _
            "Spot", "Valeur", "Delta", "Gamma", _
            "Vega / 1 pt vol", "Theta / jour", "Rho / 1 pt taux", _
            "PnL échéance", "PnL MTM")
        .Range("A2").Resize(pointCount, 9).Value = results
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(0, 112, 192)
        .Range("A1:I1").Font.Color = vbWhite
        .Range("A2:A" & pointCount + 1).NumberFormat = "0.00"
        .Range("B2:I" & pointCount + 1).NumberFormat = "0.000000"
        .Columns("A:I").AutoFit
        .Visible = xlSheetVeryHidden
    End With

    wsPortfolio.Range("M2").Value = "ANALYSE PnL & RISQUE"
    Port_BuildGrid = True
    GoTo BUILD_EXIT

BUILD_ERROR:
    MsgBox "Erreur pendant le calcul : " & Err.Description, _
           vbCritical, "Calcul interrompu"

BUILD_EXIT:
    Application.StatusBar = False
    Application.ScreenUpdating = previousScreenUpdating
End Function

Public Sub InvalidateResults()
    Dim wsGrid As Worksheet
    Dim wsPortfolio As Worksheet

    Set wsGrid = GetOrCreateSheet(SHEET_GRID)
    Set wsPortfolio = GetOrCreateSheet(SHEET_PORT)
    wsGrid.Cells.ClearContents
    wsPortfolio.Range("M2").Value = "ANALYSE PnL & RISQUE — À RECALCULER"
End Sub

' =========================================================================
' 4) GRAPHIQUES
' =========================================================================

Private Function IsManagedChart(ByVal chartObject As ChartObject) As Boolean
    On Error GoTo NOT_MANAGED

    If LCase$(Left$(chartObject.Name, 7)) = "pricer_" Then
        IsManagedChart = True
        Exit Function
    End If

    Dim seriesIndex As Long
    For seriesIndex = 1 To chartObject.Chart.SeriesCollection.Count
        If InStr(1, chartObject.Chart.SeriesCollection(seriesIndex).Formula, _
                 SHEET_GRID, vbTextCompare) > 0 Then
            IsManagedChart = True
            Exit Function
        End If
    Next seriesIndex

NOT_MANAGED:
End Function

Private Sub DeleteManagedCharts(ByVal ws As Worksheet)
    Dim chartIndex As Long
    For chartIndex = ws.ChartObjects.Count To 1 Step -1
        If IsManagedChart(ws.ChartObjects(chartIndex)) Then
            ws.ChartObjects(chartIndex).Delete
        End If
    Next chartIndex
End Sub

Private Sub ConfigureAxes(ByVal chart As Chart, ByVal xTitle As String, _
                          ByVal yTitle As String, ByVal yNumberFormat As String)
    With chart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = xTitle
        .TickLabels.NumberFormat = "0.00"
    End With

    With chart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = yTitle
        .TickLabels.NumberFormat = yNumberFormat
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
    End With
End Sub

Public Sub Port_PlotCharts_All(Optional ByVal showConfirmation As Boolean = True)
    Dim previousScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo CHART_ERROR

    Dim wsGrid As Worksheet
    Dim wsPortfolio As Worksheet
    Dim wsGreeks As Worksheet
    Dim lastRow As Long
    Dim xRange As Range

    Set wsGrid = GetOrCreateSheet(SHEET_GRID)
    Set wsPortfolio = GetOrCreateSheet(SHEET_PORT)
    Set wsGreeks = GetOrCreateSheet(SHEET_GREEKS)

    lastRow = wsGrid.Cells(wsGrid.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Aucune donnée calculée. Cliquez sur « Calculer & Tracer ».", _
               vbExclamation, "Données absentes"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set xRange = wsGrid.Range("A2:A" & lastRow)

    DeleteManagedCharts wsPortfolio
    DeleteManagedCharts wsGreeks

    Dim underlying As String
    underlying = Trim$(CStr(wsPortfolio.Range("F4").Value2))
    If Len(underlying) = 0 Then underlying = "Sous-jacent"

    Dim pnlChart As ChartObject
    Set pnlChart = wsPortfolio.ChartObjects.Add( _
        Left:=wsPortfolio.Range("M3").Left, _
        Top:=wsPortfolio.Range("M3").Top, Width:=600, Height:=350)
    pnlChart.Name = CHART_PNL

    With pnlChart.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .HasTitle = True
        .ChartTitle.Text = "Profil de PnL — " & underlying
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .ChartArea.Format.Line.Visible = msoFalse

        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "MTM (aujourd'hui)"
        .SeriesCollection(1).XValues = xRange
        .SeriesCollection(1).Values = wsGrid.Range("I2:I" & lastRow)
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(1).Format.Line.Weight = 2#

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "À l'échéance"
        .SeriesCollection(2).XValues = xRange
        .SeriesCollection(2).Values = wsGrid.Range("H2:H" & lastRow)
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(220, 53, 69)
        .SeriesCollection(2).Format.Line.Weight = 2#
        .SeriesCollection(2).Format.Line.DashStyle = msoLineDash
    End With
    ConfigureAxes pnlChart.Chart, "Prix du sous-jacent", _
                  "PnL par unité", "0.00"

    wsGreeks.Range("A1:P60").Interior.Color = RGB(250, 250, 250)
    wsGreeks.Range("B2").Value = "ANALYSES DE SENSIBILITÉ (GRECS)"
    wsGreeks.Range("B2").Font.Size = 16
    wsGreeks.Range("B2").Font.Bold = True
    wsGreeks.Range("B2").Font.Color = RGB(45, 62, 80)

    Dim configs As Variant
    configs = Array( _
        Array(CHART_DELTA, "Delta", 3, "Delta", 0, 0), _
        Array(CHART_GAMMA, "Gamma", 4, "Gamma", 1, 0), _
        Array(CHART_VEGA, "Vega", 5, "Vega / +1 pt vol", 0, 1), _
        Array(CHART_THETA, "Theta", 6, "Theta / jour", 1, 1), _
        Array(CHART_RHO, "Rho", 7, "Rho / +1 pt taux", 0, 2))

    Dim configIndex As Long
    Dim chartConfig As Variant
    Dim greekChart As ChartObject
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim chartWidth As Double
    Dim chartHeight As Double
    Dim chartGap As Double

    chartWidth = 450#
    chartHeight = 250#
    chartGap = 20#

    For configIndex = LBound(configs) To UBound(configs)
        chartConfig = configs(configIndex)
        chartLeft = 20# + chartConfig(4) * (chartWidth + chartGap)
        chartTop = 50# + chartConfig(5) * (chartHeight + chartGap)

        Set greekChart = wsGreeks.ChartObjects.Add( _
            Left:=chartLeft, Top:=chartTop, _
            Width:=chartWidth, Height:=chartHeight)
        greekChart.Name = chartConfig(0)

        With greekChart.Chart
            .ChartType = xlXYScatterLinesNoMarkers
            .HasTitle = True
            .ChartTitle.Text = "Portefeuille — " & chartConfig(1)
            .HasLegend = False
            .ChartArea.Format.Line.Visible = msoFalse
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Name = chartConfig(1)
            .SeriesCollection(1).XValues = xRange
            .SeriesCollection(1).Values = wsGrid.Range( _
                wsGrid.Cells(2, chartConfig(2)), _
                wsGrid.Cells(lastRow, chartConfig(2)))
            .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
            .SeriesCollection(1).Format.Line.Weight = 1.75
        End With

        ConfigureAxes greekChart.Chart, "Prix du sous-jacent", _
                      chartConfig(3), "0.0000"
    Next configIndex

    wsGreeks.Visible = xlSheetVisible
    wsPortfolio.Activate

    If showConfirmation Then
        MsgBox "Calcul et graphiques mis à jour." & vbCrLf & _
               "• PnL : feuille Portfolio" & vbCrLf & _
               "• Grecs : feuille Greeks", _
               vbInformation, "Mise à jour terminée"
    End If

    GoTo CHART_EXIT

CHART_ERROR:
    MsgBox "Erreur pendant la création des graphiques : " & Err.Description, _
           vbCritical, "Graphiques interrompus"

CHART_EXIT:
    Application.ScreenUpdating = previousScreenUpdating
End Sub

Public Sub Ouvrir_Portfolio()
    frmPortfolio.Show
End Sub

Public Sub LanceTout(Optional ByVal showConfirmation As Boolean = True)
    DesignModule.Update_Design
    If Port_BuildGrid() Then
        Port_PlotCharts_All showConfirmation
    End If
End Sub
