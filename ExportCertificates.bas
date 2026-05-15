Attribute VB_Name = "ModExportCertificates"
'=============================================================================
' MODULE : ExportCertificates
' DESCRIPTION : Systeme VBA de generation automatique de certificats Word
'               a partir des donnees employes dans Excel.
'
' FONCTIONNALITES :
'   - Lecture automatique des IDs employes (colonne A a partir de A2)
'   - Ouverture d'un modele Word (.docx) en arriere-plan
'   - Remplacement des placeholders : <<NOM>>, <<PRENOM>>, <<CIN>>, <<MOTIF>>
'   - Sauvegarde automatique des certificats dans un dossier
'   - Impression automatique de chaque certificat
'   - Fermeture automatique de Word apres traitement
'   - Compatible avec les tableaux Word
'   - Preserve la mise en page du modele
'
' STRUCTURE EXCEL ATTENDUE :
'   Colonne A : ID Employe
'   Colonne B : Nom
'   Colonne C : Prenom
'   Colonne D : CIN
'   Colonne E : Motif
'
' UTILISATION :
'   1. Remplir les donnees employes dans la feuille "Employes"
'   2. Placer le modele Word dans le chemin defini (templatePath)
'   3. Executer la macro ExportCertificates
'=============================================================================

Option Explicit

'-----------------------------------------------------------------------------
' CONSTANTES DE CONFIGURATION
' Modifier ces valeurs selon votre environnement
'-----------------------------------------------------------------------------
Private Const TEMPLATE_PATH As String = "C:\Certificats\Template\certificat.docx"
Private Const OUTPUT_FOLDER As String = "C:\Certificats\Output\"
Private Const SHEET_NAME As String = "Employes"

'=============================================================================
' PROCEDURE PRINCIPALE : ExportCertificates
'=============================================================================
Public Sub ExportCertificates()
    
    ' --- Declaration des variables ---
    Dim wdApp As Object          ' Application Word
    Dim wdDoc As Object          ' Document Word actuel
    Dim ws As Worksheet          ' Feuille Excel des employes
    Dim lastRow As Long          ' Derniere ligne utilisee
    Dim i As Long                ' Compteur de boucle
    Dim employeeID As String     ' ID de l'employe
    Dim nom As String            ' Nom de l'employe
    Dim prenom As String         ' Prenom de l'employe
    Dim cin As String            ' CIN de l'employe
    Dim motif As String          ' Motif du certificat
    Dim outputFile As String     ' Chemin du fichier de sortie
    Dim countSuccess As Long     ' Compteur de certificats generes
    Dim countError As Long       ' Compteur d'erreurs
    
    ' --- Verification de la feuille ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La feuille '" & SHEET_NAME & "' n'existe pas." & vbCrLf & _
               "Veuillez creer une feuille nommee '" & SHEET_NAME & "' avec les colonnes :" & vbCrLf & _
               "A: ID | B: Nom | C: Prenom | D: CIN | E: Motif", _
               vbCritical, "Erreur - Feuille introuvable"
        Exit Sub
    End If
    
    ' --- Verification du modele Word ---
    If Dir(TEMPLATE_PATH) = "" Then
        MsgBox "Le modele Word est introuvable :" & vbCrLf & _
               TEMPLATE_PATH & vbCrLf & vbCrLf & _
               "Veuillez verifier le chemin du fichier template.", _
               vbCritical, "Erreur - Modele introuvable"
        Exit Sub
    End If
    
    ' --- Creation du dossier de sortie si inexistant ---
    If Dir(OUTPUT_FOLDER, vbDirectory) = "" Then
        MkDir OUTPUT_FOLDER
    End If
    
    ' --- Trouver la derniere ligne remplie ---
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Aucun employe trouve dans la feuille." & vbCrLf & _
               "Veuillez remplir les donnees a partir de la ligne 2.", _
               vbExclamation, "Aucune donnee"
        Exit Sub
    End If
    
    ' --- Ouvrir Word en arriere-plan (invisible) ---
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "Impossible de lancer Microsoft Word." & vbCrLf & _
               "Veuillez verifier que Word est installe.", _
               vbCritical, "Erreur - Word"
        Exit Sub
    End If
    
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0  ' wdAlertsNone = 0
    
    ' --- Initialisation des compteurs ---
    countSuccess = 0
    countError = 0
    
    ' --- Afficher la barre de statut ---
    Application.StatusBar = "Generation des certificats en cours..."
    Application.ScreenUpdating = False
    
    ' --- Boucle sur tous les employes ---
    For i = 2 To lastRow
        
        ' Lire l'ID de l'employe
        employeeID = Trim(CStr(ws.Cells(i, 1).Value))
        
        ' Verifier que l'ID n'est pas vide
        If employeeID <> "" Then
            
            ' Lire les informations de l'employe
            nom = Trim(CStr(ws.Cells(i, 2).Value))
            prenom = Trim(CStr(ws.Cells(i, 3).Value))
            cin = Trim(CStr(ws.Cells(i, 4).Value))
            motif = Trim(CStr(ws.Cells(i, 5).Value))
            
            ' Mettre a jour la barre de statut
            Application.StatusBar = "Traitement employe " & (i - 1) & "/" & (lastRow - 1) & _
                                    " - ID: " & employeeID
            
            ' Ouvrir le modele Word
            On Error Resume Next
            Set wdDoc = wdApp.Documents.Open(TEMPLATE_PATH, ReadOnly:=True)
            On Error GoTo 0
            
            If wdDoc Is Nothing Then
                countError = countError + 1
                GoTo NextEmployee
            End If
            
            ' --- Remplacement des placeholders ---
            ' Dans le corps du document
            Call ReplaceTextInDocument(wdDoc, "<<NOM>>", nom)
            Call ReplaceTextInDocument(wdDoc, "<<PRENOM>>", prenom)
            Call ReplaceTextInDocument(wdDoc, "<<CIN>>", cin)
            Call ReplaceTextInDocument(wdDoc, "<<MOTIF>>", motif)
            
            ' Dans les en-tetes et pieds de page
            Call ReplaceTextInHeadersFooters(wdDoc, "<<NOM>>", nom)
            Call ReplaceTextInHeadersFooters(wdDoc, "<<PRENOM>>", prenom)
            Call ReplaceTextInHeadersFooters(wdDoc, "<<CIN>>", cin)
            Call ReplaceTextInHeadersFooters(wdDoc, "<<MOTIF>>", motif)
            
            ' --- Sauvegarder le certificat ---
            outputFile = OUTPUT_FOLDER & "Certificat_" & employeeID & ".docx"
            
            On Error Resume Next
            wdDoc.SaveAs2 outputFile, 16  ' 16 = wdFormatDocumentDefault (.docx)
            
            If Err.Number <> 0 Then
                countError = countError + 1
                Err.Clear
                wdDoc.Close False
                Set wdDoc = Nothing
                GoTo NextEmployee
            End If
            On Error GoTo 0
            
            ' --- Impression automatique ---
            On Error Resume Next
            wdDoc.PrintOut Background:=False
            If Err.Number <> 0 Then
                ' L'impression a echoue mais le fichier est sauvegarde
                Err.Clear
            End If
            On Error GoTo 0
            
            ' --- Fermer le document sans sauvegarder a nouveau ---
            wdDoc.Close False
            Set wdDoc = Nothing
            
            countSuccess = countSuccess + 1
            
        End If
        
NextEmployee:
    Next i
    
    ' --- Fermer Word ---
    wdApp.Quit
    Set wdApp = Nothing
    
    ' --- Restaurer l'affichage ---
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' --- Message de fin ---
    MsgBox "Traitement termine !" & vbCrLf & vbCrLf & _
           "Certificats generes : " & countSuccess & vbCrLf & _
           "Erreurs : " & countError & vbCrLf & vbCrLf & _
           "Dossier de sortie :" & vbCrLf & OUTPUT_FOLDER, _
           vbInformation, "Generation terminee"
    
End Sub

'=============================================================================
' FONCTION : ReplaceTextInDocument
' Remplace un placeholder dans tout le corps du document Word
' (inclut le texte principal ET les tableaux)
'=============================================================================
Private Sub ReplaceTextInDocument(doc As Object, findText As String, replaceText As String)
    
    Dim oRange As Object
    Dim oTable As Object
    Dim cellRange As Object
    
    ' --- Remplacement dans le corps principal du document ---
    Set oRange = doc.Content
    With oRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 1  ' wdFindContinue = 1
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=2  ' wdReplaceAll = 2
    End With
    
    ' --- Remplacement dans les tableaux (securite supplementaire) ---
    Dim t As Long, r As Long, c As Long
    On Error Resume Next
    For t = 1 To doc.Tables.Count
        Set oTable = doc.Tables(t)
        For r = 1 To oTable.Rows.Count
            For c = 1 To oTable.Columns.Count
                Set cellRange = oTable.Cell(r, c).Range
                With cellRange.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = findText
                    .Replacement.Text = replaceText
                    .Forward = True
                    .Wrap = 1
                    .Format = False
                    .MatchCase = True
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .Execute Replace:=2
                End With
                Set cellRange = Nothing
            Next c
        Next r
        Set oTable = Nothing
    Next t
    On Error GoTo 0
    
End Sub

'=============================================================================
' FONCTION : ReplaceTextInHeadersFooters
' Remplace un placeholder dans les en-tetes et pieds de page
'=============================================================================
Private Sub ReplaceTextInHeadersFooters(doc As Object, findText As String, replaceText As String)
    
    Dim oSection As Object
    Dim oHeader As Object
    Dim oFooter As Object
    Dim oRange As Object
    
    On Error Resume Next
    
    Dim s As Long, h As Long
    For s = 1 To doc.Sections.Count
        Set oSection = doc.Sections(s)
        
        ' Parcourir les en-tetes (1=Primary, 2=FirstPage, 3=EvenPages)
        For h = 1 To 3
            Set oHeader = oSection.Headers(h)
            If Not oHeader Is Nothing Then
                Set oRange = oHeader.Range
                With oRange.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = findText
                    .Replacement.Text = replaceText
                    .Forward = True
                    .Wrap = 1
                    .Format = False
                    .MatchCase = True
                    .Execute Replace:=2
                End With
            End If
            
            ' Parcourir les pieds de page
            Set oFooter = oSection.Footers(h)
            If Not oFooter Is Nothing Then
                Set oRange = oFooter.Range
                With oRange.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = findText
                    .Replacement.Text = replaceText
                    .Forward = True
                    .Wrap = 1
                    .Format = False
                    .MatchCase = True
                    .Execute Replace:=2
                End With
            End If
        Next h
    Next s
    
    On Error GoTo 0
    
End Sub

'=============================================================================
' PROCEDURE BONUS : ExportCertificatesAsPDF
' Meme fonctionnalite mais exporte egalement en PDF
'=============================================================================
Public Sub ExportCertificatesAsPDF()
    
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim employeeID As String
    Dim nom As String
    Dim prenom As String
    Dim cin As String
    Dim motif As String
    Dim outputFile As String
    Dim pdfFile As String
    Dim countSuccess As Long
    
    ' --- Verification de la feuille ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La feuille '" & SHEET_NAME & "' n'existe pas.", vbCritical, "Erreur"
        Exit Sub
    End If
    
    ' --- Verification du modele ---
    If Dir(TEMPLATE_PATH) = "" Then
        MsgBox "Modele introuvable : " & TEMPLATE_PATH, vbCritical, "Erreur"
        Exit Sub
    End If
    
    ' --- Creation des dossiers ---
    If Dir(OUTPUT_FOLDER, vbDirectory) = "" Then MkDir OUTPUT_FOLDER
    If Dir(OUTPUT_FOLDER & "PDF\", vbDirectory) = "" Then MkDir OUTPUT_FOLDER & "PDF\"
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' --- Ouvrir Word ---
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0
    
    countSuccess = 0
    Application.StatusBar = "Export PDF en cours..."
    
    For i = 2 To lastRow
        employeeID = Trim(CStr(ws.Cells(i, 1).Value))
        
        If employeeID <> "" Then
            nom = Trim(CStr(ws.Cells(i, 2).Value))
            prenom = Trim(CStr(ws.Cells(i, 3).Value))
            cin = Trim(CStr(ws.Cells(i, 4).Value))
            motif = Trim(CStr(ws.Cells(i, 5).Value))
            
            Application.StatusBar = "Export PDF " & (i - 1) & "/" & (lastRow - 1)
            
            Set wdDoc = wdApp.Documents.Open(TEMPLATE_PATH, ReadOnly:=True)
            
            ' Remplacement
            Call ReplaceTextInDocument(wdDoc, "<<NOM>>", nom)
            Call ReplaceTextInDocument(wdDoc, "<<PRENOM>>", prenom)
            Call ReplaceTextInDocument(wdDoc, "<<CIN>>", cin)
            Call ReplaceTextInDocument(wdDoc, "<<MOTIF>>", motif)
            Call ReplaceTextInHeadersFooters(wdDoc, "<<NOM>>", nom)
            Call ReplaceTextInHeadersFooters(wdDoc, "<<PRENOM>>", prenom)
            Call ReplaceTextInHeadersFooters(wdDoc, "<<CIN>>", cin)
            Call ReplaceTextInHeadersFooters(wdDoc, "<<MOTIF>>", motif)
            
            ' Sauvegarder en DOCX
            outputFile = OUTPUT_FOLDER & "Certificat_" & employeeID & ".docx"
            wdDoc.SaveAs2 outputFile, 16
            
            ' Exporter en PDF
            pdfFile = OUTPUT_FOLDER & "PDF\Certificat_" & employeeID & ".pdf"
            wdDoc.ExportAsFixedFormat pdfFile, 17, False, 0  ' 17 = wdExportFormatPDF
            
            wdDoc.Close False
            Set wdDoc = Nothing
            countSuccess = countSuccess + 1
        End If
    Next i
    
    wdApp.Quit
    Set wdApp = Nothing
    
    Application.StatusBar = False
    
    MsgBox "Export termine ! " & countSuccess & " certificats generes (DOCX + PDF)." & vbCrLf & _
           "Dossier : " & OUTPUT_FOLDER, vbInformation, "Succes"
    
End Sub
