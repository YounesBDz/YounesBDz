Attribute VB_Name = "modAttestations"
'==========================================================================
' Module : modAttestations
' Projet : Fusion Tables PT/GT/AT -> TRAITE + Generation Attestations Word
' Auteur : YounesBDz
'
' Contenu :
'   - Reconstruire_TRAITE     : reconstruit toute la table TRAITE
'   - FillTraiteRow           : remplit une ligne TRAITE depuis un ID employe
'   - Generate_Attestations   : genere 1 fichier Word par ligne TRAITE
'   - ReplaceInDocument       : remplace un placeholder dans tout le document
'   - ColumnExists            : verifie l'existence d'une colonne
'   - SafeName                : nettoie un nom de fichier
'   - ConvertMotif            : convertit un code MOTIF (AT->Autorisation, ...)
'==========================================================================
Option Explicit

' === CONFIGURATION ===
Public Const SHEET_PT      As String = "PT"
Public Const SHEET_GT      As String = "GT"
Public Const SHEET_AT      As String = "AT"
Public Const SHEET_TRAITE  As String = "TRAITE"
Public Const SHEET_MOTIF   As String = "MOTIF"

' Cles de liaison : PT col 8 -> GT col 1, PT col 10 -> AT col 1
Public Const COL_PT_KEY_GT As Long = 8
Public Const COL_PT_KEY_AT As Long = 10

Public Const PT_COLS As Long = 18
Public Const GT_COLS As Long = 18
Public Const AT_COLS As Long = 18

' Sous-dossier de sortie pour les attestations Word (a cote du classeur)
Public Const OUTPUT_SUBFOLDER As String = "Attestations"


'==========================================================================
' BOUTON 1 : Reconstruire TRAITE
'   Lit PT, recherche les correspondances dans GT et AT,
'   et reecrit entierement la feuille TRAITE :
'     Colonnes 1..18  = PT
'     Colonnes 19..36 = GT
'     Colonnes 37..54 = AT
'==========================================================================
Public Sub Reconstruire_TRAITE()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsPT As Worksheet, wsGT As Worksheet, wsAT As Worksheet, wsTR As Worksheet

    On Error Resume Next
    Set wsPT = wb.Worksheets(SHEET_PT)
    Set wsGT = wb.Worksheets(SHEET_GT)
    Set wsAT = wb.Worksheets(SHEET_AT)
    Set wsTR = wb.Worksheets(SHEET_TRAITE)
    On Error GoTo 0

    If wsPT Is Nothing Then MsgBox "Feuille '" & SHEET_PT & "' introuvable.", vbCritical: Exit Sub
    If wsGT Is Nothing Then MsgBox "Feuille '" & SHEET_GT & "' introuvable.", vbCritical: Exit Sub
    If wsAT Is Nothing Then MsgBox "Feuille '" & SHEET_AT & "' introuvable.", vbCritical: Exit Sub

    ' Cree la feuille TRAITE si absente
    If wsTR Is Nothing Then
        Set wsTR = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsTR.Name = SHEET_TRAITE
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Vide le contenu de TRAITE (sans toucher aux objets / boutons)
    wsTR.Cells.Clear

    ' --- En-tetes TRAITE = en-tetes PT + GT + AT ---
    Dim c As Long
    For c = 1 To PT_COLS
        wsTR.Cells(1, c).Value = wsPT.Cells(1, c).Value
    Next c
    For c = 1 To GT_COLS
        wsTR.Cells(1, PT_COLS + c).Value = wsGT.Cells(1, c).Value
    Next c
    For c = 1 To AT_COLS
        wsTR.Cells(1, PT_COLS + GT_COLS + c).Value = wsAT.Cells(1, c).Value
    Next c

    ' Dictionnaires pour acceleration des lookups
    Dim dictGT As Object, dictAT As Object
    Set dictGT = BuildKeyDict(wsGT, 1)
    Set dictAT = BuildKeyDict(wsAT, 1)

    Dim lastPT As Long
    lastPT = wsPT.Cells(wsPT.Rows.Count, 1).End(xlUp).Row
    If lastPT < 2 Then GoTo Finalize

    Dim destRow As Long: destRow = 2
    Dim i As Long
    For i = 2 To lastPT
        ' --- Copie PT (colonnes 1..18) ---
        For c = 1 To PT_COLS
            wsTR.Cells(destRow, c).Value = wsPT.Cells(i, c).Value
        Next c

        ' --- Recherche GT via PT col 8 ---
        Dim keyGT As String
        keyGT = CStr(wsPT.Cells(i, COL_PT_KEY_GT).Value)
        If dictGT.Exists(keyGT) Then
            Dim rGT As Long: rGT = dictGT(keyGT)
            For c = 1 To GT_COLS
                wsTR.Cells(destRow, PT_COLS + c).Value = wsGT.Cells(rGT, c).Value
            Next c
        End If

        ' --- Recherche AT via PT col 10 ---
        Dim keyAT As String
        keyAT = CStr(wsPT.Cells(i, COL_PT_KEY_AT).Value)
        If dictAT.Exists(keyAT) Then
            Dim rAT As Long: rAT = dictAT(keyAT)
            For c = 1 To AT_COLS
                wsTR.Cells(destRow, PT_COLS + GT_COLS + c).Value = wsAT.Cells(rAT, c).Value
            Next c
        End If

        destRow = destRow + 1
    Next i

    ' --- Mise en forme legere de l'en-tete ---
    With wsTR.Range(wsTR.Cells(1, 1), wsTR.Cells(1, PT_COLS + GT_COLS + AT_COLS))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With
    wsTR.Columns.AutoFit

Finalize:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Table TRAITE reconstruite : " & (destRow - 2) & " lignes.", vbInformation
End Sub


'--------------------------------------------------------------------------
' BuildKeyDict : construit un dictionnaire {cle -> numero de ligne}
'--------------------------------------------------------------------------
Private Function BuildKeyDict(ws As Worksheet, keyCol As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim last As Long
    last = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    Dim r As Long, k As String
    For r = 2 To last
        k = CStr(ws.Cells(r, keyCol).Value)
        If Len(k) > 0 And Not dict.Exists(k) Then
            dict.Add k, r
        End If
    Next r
    Set BuildKeyDict = dict
End Function


'==========================================================================
' FillTraiteRow : remplit une ligne TRAITE a partir de l'ID employe
'                 saisi en colonne A (utilisee par Worksheet_Change)
'==========================================================================
Public Sub FillTraiteRow(ByVal targetRow As Long)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsPT As Worksheet, wsGT As Worksheet, wsAT As Worksheet, wsTR As Worksheet

    On Error Resume Next
    Set wsPT = wb.Worksheets(SHEET_PT)
    Set wsGT = wb.Worksheets(SHEET_GT)
    Set wsAT = wb.Worksheets(SHEET_AT)
    Set wsTR = wb.Worksheets(SHEET_TRAITE)
    On Error GoTo 0

    If wsPT Is Nothing Or wsTR Is Nothing Then Exit Sub

    Dim empID As String
    empID = CStr(wsTR.Cells(targetRow, 1).Value)
    If Len(empID) = 0 Then Exit Sub

    ' Recherche dans PT (colonne 1)
    Dim foundRow As Long
    foundRow = FindRow(wsPT, 1, empID)
    If foundRow = 0 Then
        MsgBox "Employe '" & empID & "' introuvable dans " & SHEET_PT & ".", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False

    ' Copie PT (1..18)
    Dim c As Long
    For c = 1 To PT_COLS
        wsTR.Cells(targetRow, c).Value = wsPT.Cells(foundRow, c).Value
    Next c

    ' GT (19..36) via PT col 8
    If Not wsGT Is Nothing Then
        Dim rGT As Long
        rGT = FindRow(wsGT, 1, CStr(wsPT.Cells(foundRow, COL_PT_KEY_GT).Value))
        If rGT > 0 Then
            For c = 1 To GT_COLS
                wsTR.Cells(targetRow, PT_COLS + c).Value = wsGT.Cells(rGT, c).Value
            Next c
        End If
    End If

    ' AT (37..54) via PT col 10
    If Not wsAT Is Nothing Then
        Dim rAT As Long
        rAT = FindRow(wsAT, 1, CStr(wsPT.Cells(foundRow, COL_PT_KEY_AT).Value))
        If rAT > 0 Then
            For c = 1 To AT_COLS
                wsTR.Cells(targetRow, PT_COLS + GT_COLS + c).Value = wsAT.Cells(rAT, c).Value
            Next c
        End If
    End If

    Application.EnableEvents = True
End Sub


'--------------------------------------------------------------------------
' FindRow : recherche lineaire d'une cle dans une colonne
'--------------------------------------------------------------------------
Private Function FindRow(ws As Worksheet, keyCol As Long, key As String) As Long
    FindRow = 0
    If Len(key) = 0 Then Exit Function
    Dim last As Long
    last = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    Dim r As Long
    For r = 2 To last
        If StrComp(CStr(ws.Cells(r, keyCol).Value), key, vbTextCompare) = 0 Then
            FindRow = r
            Exit Function
        End If
    Next r
End Function


'==========================================================================
' BOUTON 2 : Generate_Attestations
'   Pour chaque ligne de TRAITE, ouvre le modele Word, remplace les
'   placeholders <<NOM>>, <<PRENOM>>, <<CIN>>, <<MOTIF>>, <<DATE>>, ...
'   et enregistre un fichier .docx par employe.
'   La mise en page du modele (marges, logo, tableaux, header/footer)
'   est conservee : seuls les placeholders sont modifies.
'==========================================================================
Public Sub Generate_Attestations()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsTR As Worksheet
    On Error Resume Next
    Set wsTR = wb.Worksheets(SHEET_TRAITE)
    On Error GoTo 0
    If wsTR Is Nothing Then MsgBox "Feuille '" & SHEET_TRAITE & "' introuvable.", vbCritical: Exit Sub

    ' --- Choix du modele Word ---
    Dim templatePath As String
    templatePath = ChooseWordTemplate()
    If Len(templatePath) = 0 Then Exit Sub

    ' --- Dossier de sortie ---
    Dim outFolder As String
    outFolder = wb.Path & Application.PathSeparator & OUTPUT_SUBFOLDER
    If Len(Dir(outFolder, vbDirectory)) = 0 Then MkDir outFolder

    ' --- Map en-tetes : nom de colonne -> index ---
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    headers.CompareMode = vbTextCompare

    Dim totalCols As Long: totalCols = PT_COLS + GT_COLS + AT_COLS
    Dim c As Long, h As String
    For c = 1 To totalCols
        h = Trim(CStr(wsTR.Cells(1, c).Value))
        If Len(h) > 0 And Not headers.Exists(h) Then headers.Add h, c
    Next c

    ' --- Late binding Word ---
    Dim wordApp As Object
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then Set wordApp = CreateObject("Word.Application")
    On Error GoTo 0
    If wordApp Is Nothing Then MsgBox "Impossible d'ouvrir Microsoft Word.", vbCritical: Exit Sub
    wordApp.Visible = False

    Dim lastRow As Long
    lastRow = wsTR.Cells(wsTR.Rows.Count, 1).End(xlUp).Row
    Dim count As Long: count = 0

    Application.ScreenUpdating = False

    Dim r As Long
    For r = 2 To lastRow
        If Len(CStr(wsTR.Cells(r, 1).Value)) > 0 Then
            Dim doc As Object
            Set doc = wordApp.Documents.Open(templatePath, ReadOnly:=False)

            ' --- Remplacement de chaque placeholder <<HEADER>> ---
            Dim k As Variant, val As String
            For Each k In headers.Keys
                val = CStr(wsTR.Cells(r, headers(k)).Value)

                Select Case UCase(CStr(k))
                    Case "MOTIF"
                        val = ConvertMotif(val)
                    Case "DATE"
                        If Len(val) = 0 Then
                            val = Format(Date, "dd/mm/yyyy")
                        ElseIf IsDate(wsTR.Cells(r, headers(k)).Value) Then
                            val = Format(wsTR.Cells(r, headers(k)).Value, "dd/mm/yyyy")
                        End If
                End Select

                ReplaceInDocument doc, "<<" & CStr(k) & ">>", val
            Next k

            ' Placeholder universel pour la date du jour
            ReplaceInDocument doc, "<<DATE_AUTO>>", Format(Date, "dd/mm/yyyy")

            ' --- Sauvegarde ---
            Dim outPath As String
            outPath = outFolder & Application.PathSeparator & BuildFilename(wsTR, r, headers)

            ' 16 = wdFormatXMLDocument (.docx)
            doc.SaveAs2 outPath, 16
            doc.Close False
            count = count + 1
        End If
    Next r

    wordApp.Quit
    Set wordApp = Nothing
    Application.ScreenUpdating = True

    MsgBox count & " attestation(s) generee(s) dans :" & vbCrLf & outFolder, vbInformation
End Sub


'--------------------------------------------------------------------------
' ChooseWordTemplate : dialogue de selection du modele Word
'--------------------------------------------------------------------------
Private Function ChooseWordTemplate() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Choisir le modele Word d'attestation"
        .Filters.Clear
        .Filters.Add "Documents Word", "*.docx;*.docm;*.dotx;*.dotm;*.doc"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then
            ChooseWordTemplate = .SelectedItems(1)
        Else
            ChooseWordTemplate = ""
        End If
    End With
End Function


'--------------------------------------------------------------------------
' BuildFilename : Attestation_<NOM>_<PRENOM>.docx
'--------------------------------------------------------------------------
Private Function BuildFilename(ws As Worksheet, r As Long, headers As Object) As String
    Dim nom As String, prenom As String
    If headers.Exists("NOM") Then nom = CStr(ws.Cells(r, headers("NOM")).Value)
    If headers.Exists("PRENOM") Then prenom = CStr(ws.Cells(r, headers("PRENOM")).Value)

    Dim base As String
    If Len(nom & prenom) > 0 Then
        base = "Attestation_" & SafeName(nom)
        If Len(prenom) > 0 Then base = base & "_" & SafeName(prenom)
    Else
        base = "Attestation_Ligne" & r
    End If

    BuildFilename = Trim(base) & ".docx"
End Function


'==========================================================================
' ReplaceInDocument : remplace un texte dans tout le document
'   - corps principal
'   - tableaux (couverts par Content)
'   - en-tetes et pieds de page de toutes les sections
'==========================================================================
Public Sub ReplaceInDocument(doc As Object, findText As String, replaceText As String)
    ReplaceInRange doc.Content, findText, replaceText

    Dim sec As Object, hf As Object
    For Each sec In doc.Sections
        For Each hf In sec.Headers
            If Not hf.Range Is Nothing Then ReplaceInRange hf.Range, findText, replaceText
        Next hf
        For Each hf In sec.Footers
            If Not hf.Range Is Nothing Then ReplaceInRange hf.Range, findText, replaceText
        Next hf
    Next sec
End Sub

Private Sub ReplaceInRange(rng As Object, findText As String, replaceText As String)
    Const wdReplaceAll  As Long = 2
    Const wdFindContinue As Long = 1

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub


'==========================================================================
' ColumnExists : verifie qu'une colonne (par son en-tete) existe en ligne 1
'==========================================================================
Public Function ColumnExists(ws As Worksheet, colName As String) As Boolean
    ColumnExists = (GetColumnIndex(ws, colName) > 0)
End Function

Public Function GetColumnIndex(ws As Worksheet, colName As String) As Long
    GetColumnIndex = 0
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(CStr(ws.Cells(1, c).Value), colName, vbTextCompare) = 0 Then
            GetColumnIndex = c
            Exit Function
        End If
    Next c
End Function


'==========================================================================
' SafeName : retire les caracteres interdits dans un nom de fichier Windows
'==========================================================================
Public Function SafeName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "")
    Next i
    SafeName = Trim(s)
End Function


'==========================================================================
' ConvertMotif : convertit un code (AT, M, ...) via la feuille MOTIF
'   Feuille MOTIF :
'     Col A = Code   (ex : AT)
'     Col B = Texte  (ex : Autorisation)
'   Si la feuille ou le code n'existe pas, retourne la valeur d'origine.
'==========================================================================
Public Function ConvertMotif(ByVal code As String) As String
    ConvertMotif = code
    If Len(code) = 0 Then Exit Function

    Dim wsM As Worksheet
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets(SHEET_MOTIF)
    On Error GoTo 0
    If wsM Is Nothing Then Exit Function

    Dim last As Long
    last = wsM.Cells(wsM.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To last
        If StrComp(CStr(wsM.Cells(r, 1).Value), code, vbTextCompare) = 0 Then
            ConvertMotif = CStr(wsM.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function
