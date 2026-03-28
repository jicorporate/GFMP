Attribute VB_Name = "MOD_01_TradeGatekeeper"
Option Explicit

' =========================================================================
' MODULE: MOD_01_TradeGatekeeper
' OBJECTIF: Générateur du Formulaire de Trading (Achat/Vente/Dividendes)
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_2_GATEKEEPER()
    Application.ScreenUpdating = False
    
    Dim VBP As Object: On Error Resume Next: Set VBP = ThisWorkbook.VBProject: On Error GoTo 0
    If VBP Is Nothing Then
        MsgBox "ERREUR DE SÉCURITÉ :" & vbCrLf & "Veuillez activer l'accès au modèle d'objet VBA." & vbCrLf & _
               "(Fichier > Options > Centre de gestion de la confidentialité > Paramètres > Paramètres des macros > Cocher 'Accès approuvé au modèle d'objet du projet VBA').", vbCritical
        Exit Sub
    End If
    
    ' 1. Destruction de l'ancien formulaire s'il existe
    Dim VBComp As Object
    For Each VBComp In VBP.VBComponents
        If VBComp.Name = "USF_Trade" Then VBP.VBComponents.Remove VBComp
    Next VBComp
    
    ' 2. Création du Designer UI
    Set VBComp = VBP.VBComponents.Add(3) ' vbext_ct_MSForm
    VBComp.Properties("Name") = "USF_Trade"
    
    Dim myForm As Object: Set myForm = VBComp.Designer
    VBComp.Properties("Width") = 300
    VBComp.Properties("Height") = 420
    VBComp.Properties("Caption") = "Ordre de Bourse (Trade Entry)"
    
    Dim t As Integer: t = 10
    Creer_Controle myForm, "txt_Date", "Date d'exécution (MM/JJ/AAAA) :", "TextBox", t
    Creer_Controle myForm, "cmb_Portfolio", "Portefeuille (Compte) :", "ComboBox", t
    Creer_Controle myForm, "cmb_Asset", "Actif (Ticker) :", "ComboBox", t
    Creer_Controle myForm, "cmb_Type", "Sens de l'opération :", "ComboBox", t
    Creer_Controle myForm, "txt_Qty", "Quantité (Parts/Jetons) :", "TextBox", t
    Creer_Controle myForm, "txt_Price", "Prix Unitaire d'Exécution :", "TextBox", t
    Creer_Controle myForm, "txt_Fees", "Frais de Courtage :", "TextBox", t
    Creer_Controle myForm, "txt_FXRate", "Taux de Change (Ex: 1 si devise identique) :", "TextBox", t
    
    ' Boutons
    Dim ctrl As Object
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Save")
    ctrl.Top = t + 10: ctrl.Left = 30: ctrl.Width = 100: ctrl.Height = 25
    ctrl.Caption = "EXÉCUTER": ctrl.BackColor = RGB(46, 204, 113): ctrl.Font.Bold = True: ctrl.ForeColor = vbWhite
    
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Cancel")
    ctrl.Top = t + 10: ctrl.Left = 150: ctrl.Width = 100: ctrl.Height = 25
    ctrl.Caption = "ANNULER": ctrl.BackColor = RGB(231, 76, 60): ctrl.Font.Bold = True: ctrl.ForeColor = vbWhite
    
    ' 3. Injection du Cerveau Financier (Code Behind)
    VBComp.CodeModule.AddFromString Code_VBA_USF_Trade()
    
    Application.ScreenUpdating = True
    MsgBox "LE TRADE GATEKEEPER EST OPÉRATIONNEL." & vbCrLf & vbCrLf & _
           "Le formulaire d'ordres de bourse a été généré avec succès.", vbInformation, "WMS v1.0 - Étape 2"
End Sub

Private Sub Creer_Controle(myForm As Object, Nom As String, Titre As String, TypeCtrl As String, ByRef t As Integer)
    Dim lbl As Object, c As Object
    Set lbl = myForm.Controls.Add("Forms.Label.1", "lbl_" & Nom)
    lbl.Caption = Titre: lbl.Top = t: lbl.Left = 20: lbl.Width = 240: lbl.Height = 12
    Set c = myForm.Controls.Add("Forms." & TypeCtrl & ".1", Nom)
    c.Top = t + 12: c.Left = 20: c.Width = 240: c.Height = 18
    t = t + 35
End Sub

' --- Le Cerveau du Formulaire (Code Injecté) ---
Private Function Code_VBA_USF_Trade() As String
    Dim L() As String: ReDim L(1 To 100): Dim i As Integer: i = 1
    
    L(i) = "Option Explicit": i = i + 1
    L(i) = "Private Sub UserForm_Initialize()": i = i + 1
    L(i) = "    Me.txt_Date.Value = Format(Date, ""mm/dd/yyyy"")": i = i + 1
    L(i) = "    Me.txt_FXRate.Value = ""1""": i = i + 1
    L(i) = "    Me.txt_Fees.Value = ""0""": i = i + 1
    L(i) = "    Me.cmb_Type.List = Array(""ACHAT"", ""VENTE"", ""DIVIDENDE"", ""SPLIT"")": i = i + 1
    L(i) = "    Me.cmb_Type.ListIndex = 0": i = i + 1
    
    L(i) = "    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets(""DIM_Portfolio"")": i = i + 1
    L(i) = "    Dim r As Long: For r = 1 To wsP.ListObjects(""T_DIM_Portfolio"").ListRows.Count": i = i + 1
    L(i) = "        Me.cmb_Portfolio.AddItem wsP.ListObjects(""T_DIM_Portfolio"").DataBodyRange(r, 1).Value": i = i + 1
    L(i) = "        Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListCount - 1, 1) = wsP.ListObjects(""T_DIM_Portfolio"").DataBodyRange(r, 2).Value & "" ("" & wsP.ListObjects(""T_DIM_Portfolio"").DataBodyRange(r, 4).Value & "")"": i = i + 1"
    L(i) = "    Next r": i = i + 1
    
    L(i) = "    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets(""DIM_Asset"")": i = i + 1
    L(i) = "    For r = 1 To wsA.ListObjects(""T_DIM_Asset"").ListRows.Count": i = i + 1
    L(i) = "        Me.cmb_Asset.AddItem wsA.ListObjects(""T_DIM_Asset"").DataBodyRange(r, 1).Value": i = i + 1
    L(i) = "        Me.cmb_Asset.List(Me.cmb_Asset.ListCount - 1, 1) = wsA.ListObjects(""T_DIM_Asset"").DataBodyRange(r, 2).Value & "" - "" & wsA.ListObjects(""T_DIM_Asset"").DataBodyRange(r, 3).Value": i = i + 1
    L(i) = "    Next r": i = i + 1
    
    L(i) = "    Me.cmb_Portfolio.ColumnCount = 2: Me.cmb_Portfolio.ColumnWidths = ""0 pt;200 pt""": i = i + 1
    L(i) = "    Me.cmb_Asset.ColumnCount = 2: Me.cmb_Asset.ColumnWidths = ""0 pt;200 pt""": i = i + 1
    L(i) = "End Sub": i = i + 1
    
    L(i) = "Private Sub btn_Save_Click()": i = i + 1
    L(i) = "    If Me.cmb_Portfolio.ListIndex = -1 Or Me.cmb_Asset.ListIndex = -1 Then MsgBox ""Sélection Incomplète."", vbCritical: Exit Sub": i = i + 1
    L(i) = "    Dim dStr As String: dStr = Replace(Replace(Replace(Me.txt_Date.Value, ""-"", ""/""), ""."", ""/""), "" "", """")": i = i + 1
    L(i) = "    Dim dParts() As String: dParts = Split(dStr, ""/"")": i = i + 1
    L(i) = "    If UBound(dParts) <> 2 Then MsgBox ""Date invalide."", vbCritical: Exit Sub": i = i + 1
    L(i) = "    If Not IsNumeric(dParts(0)) Or Not IsNumeric(dParts(1)) Or Not IsNumeric(dParts(2)) Then MsgBox ""Date invalide."", vbCritical: Exit Sub": i = i + 1
    
    L(i) = "    Dim qty As Double, px As Double, fees As Double, fx As Double": i = i + 1
    L(i) = "    qty = Val(Replace(Me.txt_Qty.Value, "","", ""."")): px = Val(Replace(Me.txt_Price.Value, "","", "".""))": i = i + 1
    L(i) = "    fees = Val(Replace(Me.txt_Fees.Value, "","", ""."")): fx = Val(Replace(Me.txt_FXRate.Value, "","", "".""))": i = i + 1
    L(i) = "    If qty <= 0 Or px <= 0 Or fx <= 0 Then MsgBox ""Quantité, Prix et Taux doivent être supérieurs à 0."", vbCritical: Exit Sub": i = i + 1
    
    L(i) = "    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets(""FACT_Trade"")": i = i + 1
    L(i) = "    wsF.Unprotect ""WMS_ADMIN_2026""": i = i + 1
    L(i) = "    Dim nR As ListRow: Set nR = wsF.ListObjects(""T_FACT_Trade"").ListRows.Add": i = i + 1
    L(i) = "    nR.Range(1, 1).Value = MOD_00_WMS_Architecture.GENERER_ID(""T_FACT_Trade"")": i = i + 1
    L(i) = "    nR.Range(1, 2).Value = DateSerial(CInt(dParts(2)), CInt(dParts(0)), CInt(dParts(1)))": i = i + 1
    L(i) = "    nR.Range(1, 3).Value = Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListIndex, 0)": i = i + 1
    L(i) = "    nR.Range(1, 4).Value = Me.cmb_Asset.List(Me.cmb_Asset.ListIndex, 0)": i = i + 1
    L(i) = "    nR.Range(1, 5).Value = Me.cmb_Type.Value": i = i + 1
    L(i) = "    nR.Range(1, 6).Value = qty: nR.Range(1, 7).Value = px: nR.Range(1, 8).Value = fees: nR.Range(1, 9).Value = fx": i = i + 1
    L(i) = "    nR.Range(1, 10).Value = Now": i = i + 1
    
    L(i) = "    wsF.Protect ""WMS_ADMIN_2026"", UserInterfaceOnly:=True": i = i + 1
    L(i) = "    MsgBox ""Ordre de Bourse enregistré avec succès !"", vbInformation: Unload Me": i = i + 1
    L(i) = "End Sub": i = i + 1
    L(i) = "Private Sub btn_Cancel_Click(): Unload Me: End Sub": i = i + 1
    
    ReDim Preserve L(1 To i - 1)
    Code_VBA_USF_Trade = Join(L, vbCrLf)
End Function

