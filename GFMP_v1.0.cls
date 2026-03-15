VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GFMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' --- DEBUT PATCH 2 (Hook de Démarrage Strict) ---
Private Sub Workbook_Open()
    On Error Resume Next
    ' 1. Verrouillage du Backend
    MOD_00_Architecture.Apply_RBAC_Security
    
    ' 2. VÉRIFICATION DE LA LICENCE (PRIORITÉ ABSOLUE)
    MOD_02_AppHome_Global.VERIFIER_LICENCE
    
    ' 3. Si la licence est valide, on passe à l'Onboarding de la Devise
    MOD_02_AppHome_Global.VERIFIER_PREMIER_LANCEMENT
    On Error GoTo 0
End Sub
' --- FIN PATCH 2 ---

' =========================================================================
' MOTEUR U.C.R : Routeur d'Événements avec Architecture "Anti-Freeze"
' =========================================================================
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Sh.Name <> "APP_HOME" Then Exit Sub
    If Target.Count > 1 Then Exit Sub
    
    If Target.Column = 1 And Target.Row >= 11 And Target.Row <= 26 Then
        Dim TargetRow As Integer
        TargetRow = Target.Row
        
        ' 1. Désactivation et remise à zéro UX
        Application.EnableEvents = False
        Sh.Range("A1").Select
        
        ' 2. RÉACTIVATION IMMÉDIATE (Anti-Freeze)
        Application.EnableEvents = True
        
        ' 3. Routage Explicite
        On Error GoTo CrashRouteur
        Select Case TargetRow
            Case 11: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "FR"
            Case 12: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "EN"
            Case 13: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "ES"
            Case 14: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "PT"
            Case 15: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "DE"
            Case 16: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "IT"
            Case 17: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "NL"
            Case 18: MOD_02_AppHome_Global.EXECUTER_CHANGER_LANGUE "SV"
            
            Case 21: MOD_02_AppHome_Global.EXECUTER_ROUTER_SAISIE
            Case 22: MOD_02_AppHome_Global.EXECUTER_ROUTER_DASHBOARD
            
            ' CORRECTION : Séparation des modules futurs
            Case 23: MOD_02_AppHome_Global.EXECUTER_ROUTER_BUDGET
            Case 24: MOD_02_AppHome_Global.EXECUTER_ROUTER_NETWORTH
            ' --- DEBUT PATCH 3 (Routage UCR du Bouton FX) ---
            Case 25: MOD_99_SystemAdmin.ACTUALISER_DEVISES_WEB
            ' --- FIN PATCH 3 ---
            Case 26: MOD_02_AppHome_Global.EXECUTER_CONFIG_SYSTEME
        End Select
        Exit Sub
        
CrashRouteur:
        MsgBox "Erreur d'exécution : " & Err.Description, vbCritical, "Alerte Routeur"
    End If
End Sub

