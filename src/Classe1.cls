VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Classe1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Position sur une option européenne vanille.
' Le nom historique "Classe1" est conservé pour rester compatible avec le
' projet VBA déjà intégré dans le classeur.

Private mIsCall As Boolean
Private mPositionSign As Double
Private mQuantity As Double
Private mStrike As Double
Private mVolatility As Double
Private mMaturityYears As Double
Private mInitialPremium As Double
Private mUnderlying As String

Public Property Get IsCall() As Boolean
    IsCall = mIsCall
End Property

Public Property Let IsCall(ByVal value As Boolean)
    mIsCall = value
End Property

Public Property Get PositionSign() As Double
    PositionSign = mPositionSign
End Property

Public Property Let PositionSign(ByVal value As Double)
    If value <> 1# And value <> -1# Then
        Err.Raise vbObjectError + 2001, "Classe1.PositionSign", _
                  "Le sens de la position doit être 1 (achat) ou -1 (vente)."
    End If
    mPositionSign = value
End Property

Public Property Get Quantity() As Double
    Quantity = mQuantity
End Property

Public Property Let Quantity(ByVal value As Double)
    If value <= 0# Then
        Err.Raise vbObjectError + 2002, "Classe1.Quantity", _
                  "La quantité doit être strictement positive."
    End If
    mQuantity = value
End Property

Public Property Get Strike() As Double
    Strike = mStrike
End Property

Public Property Let Strike(ByVal value As Double)
    If value <= 0# Then
        Err.Raise vbObjectError + 2003, "Classe1.Strike", _
                  "Le strike doit être strictement positif."
    End If
    mStrike = value
End Property

Public Property Get Volatility() As Double
    Volatility = mVolatility
End Property

Public Property Let Volatility(ByVal value As Double)
    If value <= 0# Then
        Err.Raise vbObjectError + 2004, "Classe1.Volatility", _
                  "La volatilité doit être strictement positive."
    End If
    mVolatility = value
End Property

Public Property Get MaturityYears() As Double
    MaturityYears = mMaturityYears
End Property

Public Property Let MaturityYears(ByVal value As Double)
    If value <= 0# Then
        Err.Raise vbObjectError + 2005, "Classe1.MaturityYears", _
                  "La maturité doit être strictement positive."
    End If
    mMaturityYears = value
End Property

Public Property Get InitialPremium() As Double
    InitialPremium = mInitialPremium
End Property

Public Property Let InitialPremium(ByVal value As Double)
    If value < 0# Then
        Err.Raise vbObjectError + 2006, "Classe1.InitialPremium", _
                  "La prime ne peut pas être négative."
    End If
    mInitialPremium = value
End Property

Public Property Get Underlying() As String
    Underlying = mUnderlying
End Property

Public Property Let Underlying(ByVal value As String)
    value = Trim$(value)
    If Len(value) = 0 Then
        Err.Raise vbObjectError + 2007, "Classe1.Underlying", _
                  "Le nom du sous-jacent est obligatoire."
    End If
    mUnderlying = value
End Property
