Attribute VB_Name = "MOD_00_Architecture"
Option Explicit

' =========================================================================
' MODULE: MOD_00_Architecture
' OBJECTIF: Déploiement Idempotent du Star Schema & Sécurité RBAC
' =========================================================================

Public Sub DEPLOIEMENT_ETAPE_1_DB_CORE()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' 1. DÉPLOIEMENT DES DIMENSIONS (Référentiels)
    ' DIM_Compte intègre la "Classe Actif" pour le futur Bilan Patrimonial
    Creer_Table "DIM_Compte", "T_DIM_Compte", Array("ID_Compte", "Nom_Compte", "Classe_Actif", "Devise_Defaut", "Est_Actif")
    Creer_Table "DIM_Categorie", "T_DIM_Categorie", Array("ID_Categorie", "Nom_Categorie", "Type_Flux")
    Creer_Table "DIM_Tiers", "T_DIM_Tiers", Array("ID_Tiers", "Nom_Tiers", "Type_Tiers")
    
    ' 2. DÉPLOIEMENT DES FAITS (Transactions & Budgets)
    ' FACT_Transaction intègre nativement la Devise et les logs Système
    Creer_Table "FACT_Transaction", "T_FACT_Transaction", _
        Array("ID_Trans", "Date_Trans", "ID_Compte", "ID_Categorie", "ID_Tiers", "Montant", "Devise", "Notes", "SYS_User", "SYS_Date")
        
    ' FACT_Budget : La nouveauté pour le Zero-Based Budgeting
    Creer_Table "FACT_Budget", "T_FACT_Budget", _
        Array("ID_Budget", "Mois_Annee", "ID_Categorie", "Montant_Alloue", "SYS_User", "SYS_Date")
    
    ' 3. DÉPLOIEMENT SYSTÈME (Configuration & Auto-Incrémentation)
    Creer_Table "SYS_Config", "T_SYS_Config", Array("Parametre", "Valeur", "Description")
    
    ' 4. VERROUILLAGE SÉCURITÉ
    Apply_RBAC_Security
    
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "ARCHITECTURE DÉPLOYÉE AVEC SUCCÈS." & vbCrLf & vbCrLf & _
           "Les 6 tables (DIM, FACT, SYS) ont été créées en arrière-plan." & vbCrLf & _
           "La base de données est désormais sécurisée et invisible pour l'utilisateur.", vbInformation, "SFP v3.0 - Étape 1"
End Sub

' --- MOTEUR DE CRÉATION DE TABLES STRUCTUREES (IDEMPOTENT) ---
Private Sub Creer_Table(NomOnglet As String, NomTable As String, Headers As Variant)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    
    ' 1. Vérifier si l'onglet existe, sinon le créer
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NomOnglet)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = NomOnglet
    End If
    
    ' 2. Vérifier si la table existe, sinon la créer
    On Error Resume Next
    Set tbl = ws.ListObjects(NomTable)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        ws.Cells.Clear
        ' Injection des en-têtes
        For i = LBound(Headers) To UBound(Headers)
            ws.Cells(1, i + 1).Value = Headers(i)
        Next i
        
        ' Création du ListObject (Tableau Excel Officiel)
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(Headers) + 1)), , xlYes)
        tbl.Name = NomTable
        tbl.TableStyle = "TableStyleMedium15" ' Style neutre
    End If
End Sub

' --- MOTEUR DE SÉCURITÉ (RBAC) ---
Public Sub Apply_RBAC_Security()
    Dim ws As Worksheet
    Dim OngletsBaseDeDonnees As Variant
    Dim element As Variant
    Const SYS_PASS As String = "SFP_ADMIN_2026"
    
    ' Liste stricte des onglets Backend
    OngletsBaseDeDonnees = Array("DIM_Compte", "DIM_Categorie", "DIM_Tiers", "FACT_Transaction", "FACT_Budget", "SYS_Config")
    
    For Each ws In ThisWorkbook.Worksheets
        ' Verrouillage pour autoriser le VBA mais bloquer l'humain
        ws.Protect Password:=SYS_PASS, UserInterfaceOnly:=True
        
        ' Masquage de niveau "VeryHidden" si c'est une table de DB
        For Each element In OngletsBaseDeDonnees
            If ws.Name = element Then
                ws.Visible = xlSheetVeryHidden
                Exit For
            End If
        Next element
    Next ws
End Sub



