Attribute VB_Name = "ExportCertificates"
'==============================================================================
' Module        : ExportCertificates
' Auteur        : YounesBDz
' Description   : Genere automatiquement des certificats Word a partir d'un
'                 tableau Excel d'employes, les sauvegarde et les imprime.
'
' Structure attendue de la feuille Excel (feuille active par defaut) :
'   Colonne A : ID employe       (commence en A2)
'   Colonne B : NOM
'   Colonne C : PRENOM
'   Colonne D : CIN
'   Colonne E : MOTIF
'
' Placeholders attendus dans le modele Word certificat.docx :
'   <<NOM>>  <<PRENOM>>  <<CIN>>  <<MOTIF>>
'
' Pre-requis : aucun (utilise late binding, pas de reference Word a ajouter).
'==============================================================================
Option Explicit

' ---- Constantes de configuration (a adapter au besoin) ----------------------
Private Const TEMPLATE_PATH As String = "C:\Template\certificat.docx"
Private Const OUTPUT_PATH   As String = "C:\Output\"
Private Const PRINT_AFTER_SAVE As Boolean = True   ' Mettre False pour ne pas imprimer
Private Const WD_REPLACE_ALL  As Long = 2          ' wdReplaceAll
Private Const WD_FIND_CONTINUE As Long = 1         ' wdFindContinue


'==============================================================================
' Procedure principale : ExportCertificates
'==============================================================================
Public Sub ExportCertificates()
    Dim wdApp     As Object
    Dim wdDoc     As Object
    Dim ws        As Worksheet
    Dim lastRow   As Long
    Dim i         As Long
    Dim generated As Long
    Dim outFile   As String
    Dim t0        As Single

    On Error GoTo ErrHandler

    ' --- Verifications prealables -------------------------------------------
    If Dir(TEMPLATE_PATH) = "" Then
        MsgBox "Modele Word introuvable :" & vbCrLf & TEMPLATE_PATH, _
               vbCritical, "ExportCertificates"
        Exit Sub
    End If

    EnsureFolder OUTPUT_PATH

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Aucun employe trouve dans la colonne A.", vbInformation
        Exit Sub
    End If

    ' --- Acceleration Excel --------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Initialisation de Word..."

    ' --- Lancement de Word en arriere-plan -----------------------------------
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0          ' wdAlertsNone

    t0 = Timer
    generated = 0

    ' --- Boucle sur les employes --------------------------------------------
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
            Application.StatusBar = "Generation certificat " & (i - 1) & _
                                    " / " & (lastRow - 1) & "..."

            Set wdDoc = wdApp.Documents.Open(Filename:=TEMPLATE_PATH, _
                                             ReadOnly:=False, _
                                             Visible:=False)

            ' Remplacement des placeholders (compatible tableaux, en-tetes,
            ' pieds de page, zones de texte, etc.)
            ReplaceEverywhere wdDoc, "<<NOM>>",    CStr(ws.Cells(i, 2).Value)
            ReplaceEverywhere wdDoc, "<<PRENOM>>", CStr(ws.Cells(i, 3).Value)
            ReplaceEverywhere wdDoc, "<<CIN>>",    CStr(ws.Cells(i, 4).Value)
            ReplaceEverywhere wdDoc, "<<MOTIF>>",  CStr(ws.Cells(i, 5).Value)

            ' Sauvegarde individuelle
            outFile = OUTPUT_PATH & "Certificat_" & _
                      SafeFileName(CStr(ws.Cells(i, 1).Value)) & ".docx"
            wdDoc.SaveAs2 Filename:=outFile, FileFormat:=16  ' wdFormatXMLDocument

            ' Impression automatique (optionnelle)
            If PRINT_AFTER_SAVE Then
                wdDoc.PrintOut Background:=False
            End If

            wdDoc.Close SaveChanges:=False
            Set wdDoc = Nothing
            generated = generated + 1
        End If
    Next i

    ' --- Fermeture propre de Word -------------------------------------------
    wdApp.Quit SaveChanges:=False
    Set wdApp = Nothing

    RestoreExcel
    MsgBox generated & " certificat(s) genere(s) en " & _
           Format(Timer - t0, "0.0") & " s." & vbCrLf & _
           "Dossier : " & OUTPUT_PATH, vbInformation, "Termine"
    Exit Sub

ErrHandler:
    Dim errMsg As String
    errMsg = "Erreur " & Err.Number & " : " & Err.Description & _
            vbCrLf & "Ligne Excel en cours : " & i

    ' Nettoyage en cas d'erreur
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit SaveChanges:=False
    Set wdDoc = Nothing
    Set wdApp = Nothing
    RestoreExcel
    MsgBox errMsg, vbCritical, "ExportCertificates"
End Sub


'==============================================================================
' ReplaceEverywhere : remplace findText par replaceText dans TOUS les
' StoryRanges du document (corps, tableaux, en-tetes, pieds de page,
' zones de texte, notes de bas de page, etc.).
'
' Cette version est plus robuste que doc.Content.Find seul, qui ne couvre
' pas les en-tetes / pieds de page / zones de texte.
'==============================================================================
Public Sub ReplaceEverywhere(doc As Object, _
                             ByVal findText As String, _
                             ByVal replaceText As String)
    Dim story As Object
    Dim rng   As Object

    For Each story In doc.StoryRanges
        Set rng = story
        Do
            ReplaceInRange rng, findText, replaceText
            ' Parcourir aussi les sous-stories liees (zones de texte chainees)
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next story
End Sub


'==============================================================================
' ReplaceInRange : execute Find/Replace sur un Range Word donne.
'==============================================================================
Private Sub ReplaceInRange(rng As Object, _
                           ByVal findText As String, _
                           ByVal replaceText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = WD_FIND_CONTINUE
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=WD_REPLACE_ALL
    End With
End Sub


'==============================================================================
' Utilitaires
'==============================================================================
Private Sub EnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Dim trimmed As String
        trimmed = folderPath
        If Right$(trimmed, 1) = "\" Then trimmed = Left$(trimmed, Len(trimmed) - 1)
        Dim parent As String
        parent = fso.GetParentFolderName(trimmed)
        If Len(parent) > 0 And Not fso.FolderExists(parent) Then
            EnsureFolder parent & "\"
        End If
        fso.CreateFolder trimmed
    End If
End Sub

Private Function SafeFileName(ByVal s As String) As String
    Dim invalid As Variant
    Dim ch      As Variant
    invalid = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In invalid
        s = Replace(s, CStr(ch), "_")
    Next ch
    SafeFileName = Trim(s)
End Function

Private Sub RestoreExcel()
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub
