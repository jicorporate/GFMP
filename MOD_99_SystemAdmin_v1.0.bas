Attribute VB_Name = "MOD_99_SystemAdmin"
Option Explicit

' =========================================================================
' MODULE: MOD_99_SystemAdmin
' OBJECTIF: Protocoles de Réinitialisation (Soft & Hard Reset)
' =========================================================================

Public Sub RESET_SOFT_TRANSACTIONS()
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("ATTENTION : Vous êtes sur le point de SUPPRIMER DÉFINITIVEMENT TOUTES VOS TRANSACTIONS ET BUDGETS." & vbCrLf & vbCrLf & _
                     "Vos comptes et catégories seront conservés." & vbCrLf & _
                     "Voulez-vous vraiment continuer ?", vbYesNo + vbCritical + vbDefaultButton2, "SOFT RESET (Purge des Faits)")
                     
    If reponse = vbYes Then
        Application.ScreenUpdating = False
        Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
        
        ' 1. Purge des Tables de Faits
        Vider_Table "T_FACT_Transaction"
        Vider_Table "T_FACT_Budget"
        
        ' 2. Réinitialisation des Séquenceurs d'ID
        Reset_Sequence "SEQ_T_FACT_Transaction"
        Reset_Sequence "SEQ_T_FACT_Budget"
        
        ' 3. Nettoyage des Vues Dashboards si elles existent
        On Error Resume Next
        ThisWorkbook.Sheets("DASHBOARD").Delete
        ThisWorkbook.Sheets("BUDGET_ZBB").Delete
        ThisWorkbook.Sheets("NET_WORTH").Delete
        On Error GoTo 0
        
        For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
        Application.ScreenUpdating = True
        
        MsgBox "SOFT RESET TERMINÉ." & vbCrLf & "Vos données de transactions ont été effacées. L'historique est vierge.", vbInformation, "Succès"
        
        ' Retour au Hub Central
        On Error Resume Next: ThisWorkbook.Sheets("APP_HOME").Activate: On Error GoTo 0
    End If
End Sub

Public Sub RESET_HARD_FACTORY()
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("DANGER CRITIQUE : Vous êtes sur le point de faire un FACTORY RESET." & vbCrLf & vbCrLf & _
                     "Absolument TOUTES les données (Transactions, Budgets, mais aussi vos Comptes et Catégories personnalisés) vont être détruites." & vbCrLf & _
                     "Le système sera réinitialisé à son état d'usine initial." & vbCrLf & _
                     "Êtes-vous absolument sûr ?", vbYesNo + vbCritical + vbDefaultButton2, "HARD RESET (Remise à Zéro Usine)")
                     
    If reponse = vbYes Then
        Application.ScreenUpdating = False
        Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
        
        ' 1. Atomisation de TOUTES les tables
        Vider_Table "T_FACT_Transaction"
        Vider_Table "T_FACT_Budget"
        Vider_Table "T_DIM_Compte"
        Vider_Table "T_DIM_Categorie"
        Vider_Table "T_DIM_Tiers"
        
        ' 2. Destruction des Séquenceurs
        Vider_Table "T_SYS_Config"
        
        ' 3. Suppression des Vues
        On Error Resume Next
        ThisWorkbook.Sheets("DASHBOARD").Delete
        ThisWorkbook.Sheets("BUDGET_ZBB").Delete
        ThisWorkbook.Sheets("NET_WORTH").Delete
        On Error GoTo 0
        
        ' 4. Reconstruction du Cœur (Ré-injection des Master Data de l'Étape 2)
        MOD_01_CoreEngine.DEPLOIEMENT_ETAPE_2_CORE
        
        For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
        Application.ScreenUpdating = True
        
        MsgBox "HARD RESET TERMINÉ." & vbCrLf & "Le système est totalement neuf. Il ne contient que la taxonomie d'usine.", vbInformation, "Système Réinitialisé"
        
        On Error Resume Next: ThisWorkbook.Sheets("APP_HOME").Activate: On Error GoTo 0
    End If
End Sub

' -------------------------------------------------------------------------
' UTILITAIRES DE PURGE (Zéro Bug de Ligne Fantôme)
' -------------------------------------------------------------------------
Private Sub Vider_Table(NomTable As String)
    Dim ws As Worksheet, tbl As ListObject
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next: Set tbl = ws.ListObjects(NomTable): On Error GoTo 0
        If Not tbl Is Nothing Then Exit For
    Next ws
    
    If Not tbl Is Nothing Then
        If tbl.ListRows.Count > 0 Then
            ' Supprime toutes les lignes proprement
            tbl.DataBodyRange.Rows.Delete
        End If
    End If
End Sub

Private Sub Reset_Sequence(NomSequence As String)
    Dim tblSys As ListObject, i As Long
    On Error Resume Next: Set tblSys = ThisWorkbook.Sheets("SYS_Config").ListObjects("T_SYS_Config"): On Error GoTo 0
    
    If Not tblSys Is Nothing Then
        If tblSys.ListRows.Count > 0 Then
            For i = 1 To tblSys.ListRows.Count
                If tblSys.DataBodyRange(i, 1).Value = NomSequence Then
                    ' On remet le compteur à 0 (le prochain ID généré sera 1)
                    tblSys.DataBodyRange(i, 2).Value = 0
                    Exit Sub
                End If
            Next i
        End If
    End If
End Sub
Option Explicit

' =========================================================================
' MODULE: MOD_99_SystemAdmin
' OBJECTIF: Passerelle ETL pour Power BI (Ouverture/Fermeture du Backend)
' =========================================================================

Public Sub POWER_BI_ACTIVER_CONNEXION()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim OngletsDB As Variant
    Dim element As Variant
    
    ' Liste des onglets de notre Star Schema
    OngletsDB = Array("DIM_Compte", "DIM_Categorie", "DIM_Tiers", "FACT_Transaction", "FACT_Budget", "SYS_Config")
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect "SFP_ADMIN_2026"
        For Each element In OngletsDB
            If ws.Name = element Then
                ws.Visible = xlSheetVisible ' Rend les tables visibles pour le Radar Power BI
                Exit For
            End If
        Next element
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "?? MODE POWER BI : ACTIVÉ" & vbCrLf & vbCrLf & _
           "Les tables Backend sont temporairement visibles." & vbCrLf & _
           "ÉTAPE OBLIGATOIRE : Sauvegardez maintenant votre fichier Excel (Ctrl+S) !" & vbCrLf & _
           "Ensuite, allez dans Power BI et actualisez le Navigateur.", vbInformation, "ETL Bridge Ouvert"
End Sub

Public Sub POWER_BI_SECURISER_BACKEND()
    ' Referme la forteresse une fois Power BI connecté
    On Error Resume Next
    MOD_00_Architecture.Apply_RBAC_Security
    On Error GoTo 0
    
    MsgBox "?? MODE POWER BI : DÉSACTIVÉ" & vbCrLf & vbCrLf & _
           "La forteresse est verrouillée. Les tables sont à nouveau 'VeryHidden'." & vbCrLf & _
           "Power BI continuera à se mettre à jour silencieusement en arrière-plan.", vbInformation, "ETL Bridge Fermé"
End Sub
