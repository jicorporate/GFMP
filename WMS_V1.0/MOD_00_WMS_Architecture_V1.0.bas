Attribute VB_Name = "MOD_00_WMS_Architecture"
Option Explicit

' =========================================================================
' MODULE: MOD_00_WMS_Architecture
' OBJECTIF: Amorçage du Backend WMS (Wealth Management System)
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_1()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' 1. CRÉATION DU SYSTÈME ET DICTIONNAIRE
    Creer_Table "SYS_Config", "T_SYS_Config", Array("Parametre", "Valeur", "Description")
    
    ' 2. CRÉATION DES DIMENSIONS
    Creer_Table "DIM_Portfolio", "T_DIM_Portfolio", Array("ID_Portfolio", "Nom_Compte", "Courtier", "Devise_Base", "Est_Actif")
    Creer_Table "DIM_Asset", "T_DIM_Asset", Array("ID_Asset", "Ticker_Symbole", "Nom_Actif", "Classe_Actif", "Devise_Cotation", "ISIN")
    
    ' 3. CRÉATION DU REGISTRE DE TRADING (FACT)
    ' La colonne Taux_FX_Historique fige le taux de change au moment du Trade pour la compta !
    Creer_Table "FACT_Trade", "T_FACT_Trade", _
        Array("ID_Trade", "Date_Trade", "ID_Portfolio", "ID_Asset", "Type_Ordre", "Quantite", "Prix_Unitaire", "Frais_Courtage", "Taux_FX_Historique", "SYS_Date")

    ' 4. INJECTION DES DONNÉES D'AMORÇAGE
    Bootstrapper_WMS

   ' --- DEBUT PATCH (Masquage Ciblé et Sécurisé) ---
    ' 5. VERROUILLAGE ET MASQUAGE CIBLÉ
    Dim ws As Worksheet
    Dim OngletsDB As Variant
    OngletsDB = Array("SYS_Config", "DIM_Portfolio", "DIM_Asset", "FACT_Trade")
    
    ' SÉCURITÉ ABSOLUE : On garantit qu'au moins la première feuille reste visible pour éviter le crash Excel
    ThisWorkbook.Sheets(1).Visible = xlSheetVisible
    
    Dim element As Variant
    For Each ws In ThisWorkbook.Worksheets
        For Each element In OngletsDB
            If ws.Name = CStr(element) Then
                ' On masque et protège UNIQUEMENT nos tables de base de données
                ws.Visible = xlSheetVeryHidden
                ws.Protect Password:="WMS_ADMIN_2026", UserInterfaceOnly:=True
                Exit For
            End If
        Next element
    Next ws
    ' --- FIN PATCH ---

    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "FONDATIONS WMS DÉPLOYÉES AVEC SUCCÈS." & vbCrLf & vbCrLf & _
           "La base de données financière a été structurée et scellée en arrière-plan." & vbCrLf & _
           "Passez à l'étape suivante.", vbInformation, "WMS v1.0 - Étape 1"
End Sub

' --- Moteur Autonome de Création de Tables ---
Private Sub Creer_Table(NomOnglet As String, NomTable As String, Headers As Variant)
    Dim ws As Worksheet, tbl As ListObject, i As Integer
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(NomOnglet): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = NomOnglet
    End If
    On Error Resume Next: Set tbl = ws.ListObjects(NomTable): On Error GoTo 0
    If tbl Is Nothing Then
        ws.Cells.Clear
        For i = LBound(Headers) To UBound(Headers)
            ws.Cells(1, i + 1).Value = Headers(i)
        Next i
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(Headers) + 1)), , xlYes)
        tbl.Name = NomTable
        tbl.TableStyle = "TableStyleMedium15"
        tbl.ListRows(1).Delete ' Supprime la ligne vide par défaut
    End If
End Sub

' --- Injection des Données d'Exemple ---
Private Sub Bootstrapper_WMS()
    ' Variables Globales du Système
    Alimenter_Config "LANGUE_UI", "FR", "Langue Globale"
    Alimenter_Config "SYS_DEVISE_BASE", "USD", "Devise Mère du Portfolio"
    
    ' Actifs de base (À connecter aux Tickers réels de Yahoo Finance)
    Alimenter_Table "T_DIM_Asset", Array( _
        Array(1, "AAPL", "Apple Inc.", "ACTION", "USD", "US0378331005"), _
        Array(2, "BTC-USD", "Bitcoin", "CRYPTO", "USD", "CRYPTO"), _
        Array(3, "CW8.PA", "Amundi MSCI World", "ETF", "EUR", "LU1681043599"), _
        Array(4, "MCB.MU", "MCB Group", "ACTION", "MUR", "MU0004N00004"))
        
    ' Portefeuilles de base
    Alimenter_Table "T_DIM_Portfolio", Array( _
        Array(1, "PEA International", "Bourse Direct", "EUR", "OUI"), _
        Array(2, "Hardware Wallet", "Ledger Nano", "USD", "OUI"))
End Sub

' --- DEBUT PATCH (Déverrouillage autonome des Bootstrappers) ---
Private Sub Alimenter_Config(Param As String, Valeur As String, Desc As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("SYS_Config")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("T_SYS_Config")
    
    ' DÉVERROUILLAGE OBLIGATOIRE POUR CONTOURNER LE BUG MICROSOFT
    ws.Unprotect "WMS_ADMIN_2026"
    
    Dim nr As ListRow: Set nr = tbl.ListRows.Add
    nr.Range(1, 1).Value = Param: nr.Range(1, 2).Value = Valeur: nr.Range(1, 3).Value = Desc
    
    ' REVERROUILLAGE IMMÉDIAT
    ws.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
End Sub

Private Sub Alimenter_Table(NomTable As String, Lignes As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(Split(NomTable, "_", 2)(1))
    Dim tbl As ListObject: Set tbl = ws.ListObjects(NomTable)
    
    ' DÉVERROUILLAGE OBLIGATOIRE
    ws.Unprotect "WMS_ADMIN_2026"
    
    Dim item As Variant, nr As ListRow, i As Integer
    For Each item In Lignes
        Set nr = tbl.ListRows.Add
        For i = LBound(item) To UBound(item)
            nr.Range(1, i + 1).Value = item(i)
        Next i
    Next item
    
    ' REVERROUILLAGE IMMÉDIAT
    ws.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
End Sub
' --- FIN PATCH ---

' --- Fonction Système Centrale : Générateur d'ID ---
Public Function GENERER_ID(ByVal NomTable As String) As Long
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    Dim tblSys As ListObject: Set tblSys = wsSys.ListObjects("T_SYS_Config")
    Dim paramName As String: paramName = "SEQ_" & NomTable
    Dim i As Long, newID As Long, found As Boolean: found = False
    
    wsSys.Unprotect "WMS_ADMIN_2026"
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
        newRow.Range(1, 3).Value = "Séquenceur " & NomTable
    End If
    wsSys.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    GENERER_ID = newID
End Function
' --- DEBUT PATCH (Le Cerveau des Paramètres WMS) ---
' =========================================================================
' FONCTION CENTRALE : Lecture/Écriture des Paramètres Globaux
' =========================================================================
Public Function Obtenir_Parametre(NomParam As String, ValeurDefaut As String) As String
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    Dim tblConf As ListObject, i As Long
    On Error Resume Next: Set tblConf = wsSys.ListObjects("T_SYS_Config"): On Error GoTo 0
    
    If tblConf Is Nothing Then
        Obtenir_Parametre = ValeurDefaut
        Exit Function
    End If
    
    ' Lecture
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = NomParam Then
            Obtenir_Parametre = tblConf.DataBodyRange(i, 2).Value
            Exit Function
        End If
    Next i
    
    ' Écriture (Si le paramètre n'existe pas, on le crée avec la valeur par défaut)
    wsSys.Unprotect "WMS_ADMIN_2026"
    Dim nr As ListRow: Set nr = tblConf.ListRows.Add
    nr.Range(1, 1).Value = NomParam
    nr.Range(1, 2).Value = ValeurDefaut
    nr.Range(1, 3).Value = "Filtre Actif"
    wsSys.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    
    Obtenir_Parametre = ValeurDefaut
End Function
' --- FIN PATCH ---
