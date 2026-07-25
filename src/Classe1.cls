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

' Propriétés de l'option
Public isCall As Boolean
Public Sign As Double       ' 1 pour Achat (Long), -1 pour Vente (Short)
Public qty As Double
Public K As Double
Public Sigma As Double
Public T As Double          ' En années
Public Premium0 As Double   ' Prime payée/reçue unitaire
Public Label As String
