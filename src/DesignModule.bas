Attribute VB_Name = "DesignModule"
Sub Update_Design()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Portfolio")
    ws.Activate
    
    ' Nettoie la zone graphique à droite (Colonnes M à T)
    ws.Range("M:T").Clear
    ws.Range("M2").Value = "ANALYSE PnL & RISQUE"
    ws.Range("M2").Font.Size = 14
    ws.Range("M2").Font.Bold = True
    ws.Range("M2").Font.Color = RGB(45, 62, 80)
    
    ' Ajout Boutons dans la Sidebar (Simulation visuelle)
    ' Tu devras insérer de vrais boutons (Forme ou ActiveX) par dessus
    ws.Range("B26").Value = "ACTIONS"
    ws.Range("B26").Font.Color = RGB(255, 192, 0)
    ws.Range("B26").Font.Bold = True
    
    MsgBox "Design mis à jour. Ajoute maintenant 2 boutons dans la zone bleue :\n1. 'Calculer & Tracer' -> Assigne la macro 'Port_BuildGrid' puis 'Port_PlotCharts_All'\n2. 'Voir les Grecs' -> Lien vers la feuille Greeks.", vbInformation
End Sub
