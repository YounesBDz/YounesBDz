VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_TRAITE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==========================================================================
' Code de la feuille TRAITE
'
' Quand l'utilisateur saisit un ID employe en colonne A,
' la ligne complete est remplie automatiquement :
'   - PT  -> colonnes 1..18
'   - GT  -> colonnes 19..36 (via PT col 8)
'   - AT  -> colonnes 37..54 (via PT col 10)
'
' INSTALLATION :
'   Coller ce code dans le module de la feuille "TRAITE"
'   (clic droit sur l'onglet TRAITE -> Visualiser le code).
'==========================================================================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Reagir uniquement a la saisie en colonne A, ligne >= 2, une seule cellule
    If Target.Cells.CountLarge > 1 Then Exit Sub
    If Target.Column <> 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    If Len(Trim(CStr(Target.Value))) = 0 Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo CleanExit
    Call FillTraiteRow(Target.Row)

CleanExit:
    Application.EnableEvents = True
End Sub
