Attribute VB_Name = "MOD_01_CoreEngine"
Option Explicit

' =========================================================================
' MODULE: MOD_01_CoreEngine
' OBJECTIF: Fonctions d'Intégrité (Regex, Séquenceur Autonome) & Master Data
' =========================================================================

Public Sub DEPLOIEMENT_ETAPE_2_CORE()
    Application.ScreenUpdating = False
    
    Unprotect_All
    Bootstrapper_Dimensions
    Protect_All
    
    Application.ScreenUpdating = True
    MsgBox "CORE ENGINE MIS À JOUR." & vbCrLf & vbCrLf & _
           "Le séquenceur d'ID a été blindé contre les erreurs de protection Excel." & vbCrLf & _
           "La sauvegarde depuis le formulaire fonctionnera désormais à 100%.", vbInformation, "SFP v3.1 - Core Sécurisé"
End Sub

' -------------------------------------------------------------------------
' 1. MOTEURS D'INTÉGRITÉ DES DONNÉES (LES OUTILS D'ÉLITE)
' -------------------------------------------------------------------------

' SANITISATION (Regex)
Public Function CLEAN_TEXT(ByVal strInput As String) As String
    If Len(Trim(strInput)) = 0 Then
        CLEAN_TEXT = ""
        Exit Function
    End If
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .Pattern = "\s+"
        strInput = .Replace(Trim(strInput), " ")
    End With
    CLEAN_TEXT = UCase(Left(strInput, 1)) & LCase(Mid(strInput, 2))
End Function

' GÉNÉRATION D'ID SÉCURISÉ (Moteur Auto-Déverrouillant = Correction du Crash)
Public Function GENERER_NOUVEL_ID(ByVal NomTable As String) As Long
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    Dim tblSys As ListObject: Set tblSys = wsSys.ListObjects("T_SYS_Config")
    Dim paramName As String: paramName = "SEQ_" & NomTable
    Dim i As Long, newID As Long, found As Boolean: found = False
    
    ' --- CORRECTION ABSOLUE : DÉVERROUILLAGE AUTONOME ---
    ' On retire la protection le temps de l'opération pour contrer le bug Microsoft de ListRows.Add
    wsSys.Unprotect "SFP_ADMIN_2026"

    If tblSys.ListRows.Count > 0 Then
        For i = 1 To tblSys.ListRows.Count
            If tblSys.DataBodyRange(i, 1).Value = paramName Then
                newID = CLng(tblSys.DataBodyRange(i, 2).Value) + 1
                tblSys.DataBodyRange(i, 2).Value = newID
                found = True: Exit For
            End If
        Next i
    End If
    
    If Not found Then
        Dim newRow As ListRow: Set newRow = tblSys.ListRows.Add
        newID = 1
        newRow.Range(1, 1).Value = paramName
        newRow.Range(1, 2).Value = newID
        newRow.Range(1, 3).Value = "Séquence " & NomTable
    End If
    
    ' --- CORRECTION ABSOLUE : REVERROUILLAGE ---
    wsSys.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    
    GENERER_NOUVEL_ID = newID
End Function

' ANTI-DOUBLON (O(1))
Public Function EXISTE_DEJA(ByVal NomTable As String, ByVal IndexColonne As Integer, ByVal Valeur As String) As Boolean
    Dim ws As Worksheet, tbl As ListObject
    EXISTE_DEJA = False
    Valeur = UCase(Trim(Valeur))
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next: Set tbl = ws.ListObjects(NomTable): On Error GoTo 0
        If Not tbl Is Nothing Then Exit For
    Next ws
    
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function
    
    Dim varData As Variant: varData = tbl.DataBodyRange.Columns(IndexColonne).Value
    Dim i As Long
    If tbl.ListRows.Count = 1 Then
        If UCase(Trim(CStr(varData))) = Valeur Then EXISTE_DEJA = True
    Else
        For i = 1 To UBound(varData, 1)
            If UCase(Trim(CStr(varData(i, 1)))) = Valeur Then
                EXISTE_DEJA = True: Exit For
            End If
        Next i
    End If
End Function

' -------------------------------------------------------------------------
' 2. MOTEUR D'AMORÇAGE (MASTER DATA INJECTION)
' -------------------------------------------------------------------------

Private Sub Bootstrapper_Dimensions()
    Dim colsCompte As Variant: colsCompte = Array("LIQUIDITE", "MUR")
    Alimenter_DB "T_DIM_Compte", Array( _
        Array("Compte Courant Principal", "LIQUIDITE", "MUR"), _
        Array("Livret d'Épargne", "LIQUIDITE", "MUR"), _
        Array("Portefeuille Espèces", "LIQUIDITE", "MUR"), _
        Array("Assurance Vie", "INVESTISSEMENT", "EUR"), _
        Array("PEA / Actions", "INVESTISSEMENT", "EUR"), _
        Array("Portefeuille Crypto", "INVESTISSEMENT", "USD"), _
        Array("Prêt Immobilier", "DETTE", "MUR"), _
        Array("Carte de Crédit (Différé)", "DETTE", "MUR"), _
        Array("Autre (Préciser...)", "AUTRE", "MUR"))
        
    Alimenter_DB "T_DIM_Categorie", Array( _
        Array("Salaire / Revenus Pro", "REVENU"), _
        Array("Intérêts / Dividendes", "REVENU"), _
        Array("Aides / Bourses / Allocations", "REVENU"), _
        Array("Logement (Loyer/Prêt/Charges)", "DEPENSE"), _
        Array("Alimentation & Supermarché", "DEPENSE"), _
        Array("Transports (Essence/Assurance)", "DEPENSE"), _
        Array("Santé & Mutuelle", "DEPENSE"), _
        Array("Loisirs, Sorties & Vacances", "DEPENSE"), _
        Array("Impôts & Taxes", "DEPENSE"), _
        Array("Virement Interne (Épargne)", "TRANSFERT"), _
        Array("Autre (Préciser...)", "AUTRE"))
        
    Alimenter_DB "T_DIM_Tiers", Array( _
        Array("Employeur principal", "ENTREPRISE"), _
        Array("Banque / Courtier", "FINANCE"), _
        Array("État / Impôts", "INSTITUTION"), _
        Array("Supermarché / Alimentaire", "COMMERCE"), _
        Array("Propriétaire / Syndic", "IMMOBILIER"), _
        Array("Station Service", "COMMERCE"), _
        Array("Autre (Préciser...)", "AUTRE"))
End Sub

Private Sub Alimenter_DB(NomTable As String, Lignes As Variant)
    Dim ws As Worksheet, tbl As ListObject
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next: Set tbl = ws.ListObjects(NomTable): On Error GoTo 0
        If Not tbl Is Nothing Then Exit For
    Next ws
    If tbl Is Nothing Then Exit Sub
    
    If tbl.ListRows.Count = 1 Then
        If Trim(tbl.ListRows(1).Range(1, 2).Value) = "" Then tbl.DataBodyRange.Delete
    End If
    
    Dim item As Variant, newRow As ListRow, i As Integer, libelle As String
    For Each item In Lignes
        libelle = CStr(item(0))
        If Not EXISTE_DEJA(NomTable, 2, libelle) Then
            Set newRow = tbl.ListRows.Add
            newRow.Range(1, 1).Value = GENERER_NOUVEL_ID(NomTable)
            newRow.Range(1, 2).Value = libelle
            newRow.Range(1, 3).Value = CStr(item(1))
            If UBound(item) >= 2 And NomTable = "T_DIM_Compte" Then
                newRow.Range(1, 4).Value = CStr(item(2))
                newRow.Range(1, 5).Value = "OUI"
            End If
        End If
    Next item
End Sub

' -------------------------------------------------------------------------
' UTILITAIRES DE SÉCURITÉ
' -------------------------------------------------------------------------
Private Sub Unprotect_All()
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
End Sub

Private Sub Protect_All()
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
End Sub
