Attribute VB_Name = "MOD_01_TradeGatekeeper"
Option Explicit

' =========================================================================
' MODULE: MOD_01_TradeGatekeeper
' OBJECTIF: Générateur du Formulaire de Trading (Avec Auto-Apprentissage)
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_2_GATEKEEPER()
    Application.ScreenUpdating = False
    
    Dim VBP As Object: On Error Resume Next: Set VBP = ThisWorkbook.VBProject: On Error GoTo 0
    If VBP Is Nothing Then
        MsgBox "Activez l'accès au modèle d'objet VBA.", vbCritical
        Exit Sub
    End If
    
    Dim VBComp As Object
    For Each VBComp In VBP.VBComponents
        If VBComp.Name = "USF_Trade" Then VBP.VBComponents.Remove VBComp
    Next VBComp
    
    Set VBComp = VBP.VBComponents.Add(3)
    VBComp.Properties("Name") = "USF_Trade"
    
    Dim myForm As Object: Set myForm = VBComp.Designer
    VBComp.Properties("Width") = 300
    VBComp.Properties("Height") = 550
    VBComp.Properties("Caption") = "Ordre de Bourse (Trade Entry)"
    
    Dim t As Integer: t = 10
    Creer_Controle myForm, "txt_Date", "Date d'exécution (MM/JJ/AAAA) :", "TextBox", t
    
    Creer_Controle myForm, "cmb_Portfolio", "Portefeuille (Compte) :", "ComboBox", t
    Creer_Controle_Double myForm, "txt_New_Port", "txt_New_Broker", "Précisez : Nom du Portefeuille", "Courtier", t
    
    Creer_Controle myForm, "cmb_Asset", "Actif (Ticker) :", "ComboBox", t
    Creer_Controle_Double_Cmb myForm, "txt_New_Asset", "cmb_New_Class", "Précisez : Ticker (Ex: TSLA)", "Classe d'actif", t
    
    Creer_Controle myForm, "cmb_Type", "Sens de l'opération :", "ComboBox", t
    Creer_Controle myForm, "txt_Qty", "Quantité (Parts/Jetons) :", "TextBox", t
    Creer_Controle myForm, "txt_Price", "Prix Unitaire d'Exécution :", "TextBox", t
    Creer_Controle myForm, "txt_Fees", "Frais de Courtage :", "TextBox", t
    Creer_Controle myForm, "txt_FXRate", "Taux de Change (Ex: 1 si devise identique) :", "TextBox", t
    
    Dim ctrl As Object
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Save")
    ctrl.Top = t + 10: ctrl.Left = 30: ctrl.Width = 100: ctrl.Height = 25
    ctrl.Caption = "EXÉCUTER": ctrl.BackColor = RGB(46, 204, 113): ctrl.Font.Bold = True: ctrl.ForeColor = vbWhite
    
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Cancel")
    ctrl.Top = t + 10: ctrl.Left = 150: ctrl.Width = 100: ctrl.Height = 25
    ctrl.Caption = "ANNULER": ctrl.BackColor = RGB(231, 76, 60): ctrl.Font.Bold = True: ctrl.ForeColor = vbWhite
    
    VBComp.CodeModule.AddFromString Code_VBA_USF_Trade()
    
    Application.ScreenUpdating = True
    MsgBox "TRADE GATEKEEPER MIS À JOUR." & vbCrLf & "L'Auto-Apprentissage et l'affichage des comptes sont restaurés.", vbInformation, "WMS v1.0 - Étape 2"
End Sub

Private Sub Creer_Controle(myForm As Object, Nom As String, Titre As String, TypeCtrl As String, ByRef t As Integer)
    Dim lbl As Object, c As Object
    Set lbl = myForm.Controls.Add("Forms.Label.1", "lbl_" & Nom)
    lbl.Caption = Titre: lbl.Top = t: lbl.Left = 20: lbl.Width = 240: lbl.Height = 12
    Set c = myForm.Controls.Add("Forms." & TypeCtrl & ".1", Nom)
    c.Top = t + 12: c.Left = 20: c.Width = 240: c.Height = 18
    t = t + 35
End Sub

Private Sub Creer_Controle_Double(myForm As Object, NomT1 As String, NomT2 As String, Tip1 As String, Tip2 As String, ByRef t As Integer)
    Dim lbl As Object, c1 As Object, c2 As Object
    Set lbl = myForm.Controls.Add("Forms.Label.1", "lbl_" & NomT1)
    lbl.Top = t: lbl.Left = 20: lbl.Width = 220: lbl.Height = 12: lbl.Visible = False: lbl.Caption = "Nouveau Compte (Nom & Courtier) :"
    Set c1 = myForm.Controls.Add("Forms.TextBox.1", NomT1)
    c1.Top = t + 12: c1.Left = 20: c1.Width = 125: c1.Height = 18: c1.Visible = False: c1.ControlTipText = Tip1
    Set c2 = myForm.Controls.Add("Forms.TextBox.1", NomT2)
    c2.Top = t + 12: c2.Left = 150: c2.Width = 110: c2.Height = 18: c2.Visible = False: c2.ControlTipText = Tip2
    t = t + 35
End Sub

Private Sub Creer_Controle_Double_Cmb(myForm As Object, NomT1 As String, NomCmb As String, Tip1 As String, Tip2 As String, ByRef t As Integer)
    Dim lbl As Object, c1 As Object, c2 As Object
    Set lbl = myForm.Controls.Add("Forms.Label.1", "lbl_" & NomT1)
    lbl.Top = t: lbl.Left = 20: lbl.Width = 220: lbl.Height = 12: lbl.Visible = False: lbl.Caption = "Nouvel Actif (Ticker & Classe) :"
    Set c1 = myForm.Controls.Add("Forms.TextBox.1", NomT1)
    c1.Top = t + 12: c1.Left = 20: c1.Width = 125: c1.Height = 18: c1.Visible = False: c1.ControlTipText = Tip1
    Set c2 = myForm.Controls.Add("Forms.ComboBox.1", NomCmb)
    c2.Top = t + 12: c2.Left = 150: c2.Width = 110: c2.Height = 18: c2.Visible = False: c2.ControlTipText = Tip2
    t = t + 35
End Sub

Private Function Code_VBA_USF_Trade() As String
    Dim L() As String: ReDim L(1 To 150): Dim i As Integer: i = 1
    
    ' --- DEBUT PATCH (Générateur Infaillible des Listes Déroulantes) ---
    L(i) = "Option Explicit": i = i + 1
    L(i) = "Private Function TR(Cle As String) As String: TR = MOD_02_WMS_Hub.TR(Cle): End Function": i = i + 1
    
    L(i) = "Private Sub UserForm_Initialize()": i = i + 1
    L(i) = "    Me.Caption = TR(""CARD_T_T"")": i = i + 1
    L(i) = "    Me.txt_Date.Value = Format(Date, ""mm/dd/yyyy"")": i = i + 1
    L(i) = "    Me.txt_FXRate.Value = ""1""": i = i + 1
    L(i) = "    Me.txt_Fees.Value = ""0""": i = i + 1
    
    L(i) = "    Me.cmb_Type.List = Array(""ACHAT"", ""VENTE"", ""DIVIDENDE"", ""SPLIT"")": i = i + 1
    L(i) = "    Me.cmb_Type.ListIndex = 0": i = i + 1
    
    L(i) = "    Me.cmb_New_Class.List = Array(""ACTION"", ""ETF"", ""CRYPTO"", ""OBLIGATION"")": i = i + 1
    L(i) = "    Me.cmb_New_Class.ListIndex = 0": i = i + 1
    
    L(i) = "    Me.cmb_Portfolio.ColumnCount = 2: Me.cmb_Portfolio.ColumnWidths = ""0 pt;200 pt""": i = i + 1
    L(i) = "    Me.cmb_Asset.ColumnCount = 2: Me.cmb_Asset.ColumnWidths = ""0 pt;200 pt""": i = i + 1
    
    ' --- INJECTION SÉCURISÉE : PORTEFEUILLES ---
    L(i) = "    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets(""DIM_Portfolio"")": i = i + 1
    L(i) = "    Dim r As Long: For r = 1 To wsP.ListObjects(""T_DIM_Portfolio"").ListRows.Count": i = i + 1
    L(i) = "        Me.cmb_Portfolio.AddItem wsP.ListObjects(""T_DIM_Portfolio"").DataBodyRange(r, 1).Value": i = i + 1
    L(i) = "        Dim pName As String: pName = CStr(wsP.ListObjects(""T_DIM_Portfolio"").DataBodyRange(r, 2).Value)": i = i + 1
    L(i) = "        Dim pBroker As String: pBroker = CStr(wsP.ListObjects(""T_DIM_Portfolio"").DataBodyRange(r, 3).Value)": i = i + 1
    L(i) = "        Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListCount - 1, 1) = pName & "" ("" & pBroker & "")""": i = i + 1
    L(i) = "    Next r": i = i + 1
    L(i) = "    Me.cmb_Portfolio.AddItem ""AUTRE""": i = i + 1
    L(i) = "    Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListCount - 1, 1) = TR(""Autre (Préciser...)"")": i = i + 1
    
    ' --- INJECTION SÉCURISÉE : ACTIFS ---
    L(i) = "    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets(""DIM_Asset"")": i = i + 1
    L(i) = "    For r = 1 To wsA.ListObjects(""T_DIM_Asset"").ListRows.Count": i = i + 1
    L(i) = "        Me.cmb_Asset.AddItem wsA.ListObjects(""T_DIM_Asset"").DataBodyRange(r, 1).Value": i = i + 1
    L(i) = "        Dim aTicker As String: aTicker = CStr(wsA.ListObjects(""T_DIM_Asset"").DataBodyRange(r, 2).Value)": i = i + 1
    L(i) = "        Dim aName As String: aName = CStr(wsA.ListObjects(""T_DIM_Asset"").DataBodyRange(r, 3).Value)": i = i + 1
    L(i) = "        Me.cmb_Asset.List(Me.cmb_Asset.ListCount - 1, 1) = aTicker & "" - "" & aName": i = i + 1
    L(i) = "    Next r": i = i + 1
    L(i) = "    Me.cmb_Asset.AddItem ""AUTRE""": i = i + 1
    L(i) = "    Me.cmb_Asset.List(Me.cmb_Asset.ListCount - 1, 1) = TR(""Autre (Préciser...)"")": i = i + 1
    
    L(i) = "End Sub": i = i + 1
    ' --- FIN PATCH ---
    
    ' --- DEBUT PATCH (Correction Erreur 94 - Anti Null) ---
    L(i) = "Private Sub cmb_Portfolio_Change()": i = i + 1
    L(i) = "    Dim isOther As Boolean: isOther = False": i = i + 1
    L(i) = "    If Me.cmb_Portfolio.ListIndex <> -1 Then isOther = (CStr(Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListIndex, 0)) = ""AUTRE"")": i = i + 1
    L(i) = "    Me.txt_New_Port.Visible = isOther: Me.txt_New_Broker.Visible = isOther: Me.lbl_txt_New_Port.Visible = isOther": i = i + 1
    L(i) = "End Sub": i = i + 1

    L(i) = "Private Sub cmb_Asset_Change()": i = i + 1
    L(i) = "    Dim isOther As Boolean: isOther = False": i = i + 1
    L(i) = "    If Me.cmb_Asset.ListIndex <> -1 Then isOther = (CStr(Me.cmb_Asset.List(Me.cmb_Asset.ListIndex, 0)) = ""AUTRE"")": i = i + 1
    L(i) = "    Me.txt_New_Asset.Visible = isOther: Me.cmb_New_Class.Visible = isOther: Me.lbl_txt_New_Asset.Visible = isOther": i = i + 1
    L(i) = "End Sub": i = i + 1
    ' --- FIN PATCH ---
    
    L(i) = "Private Sub btn_Save_Click()": i = i + 1
    L(i) = "    If Me.cmb_Portfolio.ListIndex = -1 Or Me.cmb_Asset.ListIndex = -1 Then MsgBox ""Sélection Incomplète."", vbCritical: Exit Sub": i = i + 1
    L(i) = "    Dim dStr As String: dStr = Replace(Replace(Replace(Me.txt_Date.Value, ""-"", ""/""), ""."", ""/""), "" "", """")": i = i + 1
    L(i) = "    Dim dParts() As String: dParts = Split(dStr, ""/"")": i = i + 1
    L(i) = "    If UBound(dParts) <> 2 Then MsgBox ""Date invalide."", vbCritical: Exit Sub": i = i + 1
    L(i) = "    If Not IsNumeric(dParts(0)) Or Not IsNumeric(dParts(1)) Or Not IsNumeric(dParts(2)) Then MsgBox ""Date invalide."", vbCritical: Exit Sub": i = i + 1
    
    L(i) = "    Dim qty As Double, px As Double, fees As Double, fx As Double": i = i + 1
    L(i) = "    qty = Val(Replace(Me.txt_Qty.Value, "","", ""."")): px = Val(Replace(Me.txt_Price.Value, "","", "".""))": i = i + 1
    L(i) = "    fees = Val(Replace(Me.txt_Fees.Value, "","", ""."")): fx = Val(Replace(Me.txt_FXRate.Value, "","", "".""))": i = i + 1
    L(i) = "    If qty <= 0 Or px <= 0 Or fx <= 0 Then MsgBox ""Quantité, Prix et Taux doivent être > 0."", vbCritical: Exit Sub": i = i + 1
    
    L(i) = "    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets(""DIM_Portfolio"")": i = i + 1
    L(i) = "    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets(""DIM_Asset"")": i = i + 1
    L(i) = "    Dim idP As Variant: idP = Me.cmb_Portfolio.Value": i = i + 1
    L(i) = "    Dim idA As Variant: idA = Me.cmb_Asset.Value": i = i + 1
    
    ' Auto-Learning Portfolio
    L(i) = "    If idP = ""AUTRE"" Then": i = i + 1
    L(i) = "        wsP.Unprotect ""WMS_ADMIN_2026""": i = i + 1
    L(i) = "        Dim nrP As ListRow: Set nrP = wsP.ListObjects(""T_DIM_Portfolio"").ListRows.Add": i = i + 1
    L(i) = "        idP = MOD_00_WMS_Architecture.GENERER_ID(""T_DIM_Portfolio"")": i = i + 1
    L(i) = "        nrP.Range(1, 1).Value = idP: nrP.Range(1, 2).Value = Me.txt_New_Port.Value: nrP.Range(1, 3).Value = Me.txt_New_Broker.Value: nrP.Range(1, 4).Value = MOD_00_WMS_Architecture.Obtenir_Parametre(""SYS_DEVISE_BASE"", ""USD""): nrP.Range(1, 5).Value = ""OUI""": i = i + 1
    L(i) = "        wsP.Protect ""WMS_ADMIN_2026"", UserInterfaceOnly:=True": i = i + 1
    L(i) = "    End If": i = i + 1

    ' Auto-Learning Asset
    L(i) = "    If idA = ""AUTRE"" Then": i = i + 1
    L(i) = "        wsA.Unprotect ""WMS_ADMIN_2026""": i = i + 1
    L(i) = "        Dim nrA As ListRow: Set nrA = wsA.ListObjects(""T_DIM_Asset"").ListRows.Add": i = i + 1
    L(i) = "        idA = MOD_00_WMS_Architecture.GENERER_ID(""T_DIM_Asset"")": i = i + 1
    L(i) = "        nrA.Range(1, 1).Value = idA: nrA.Range(1, 2).Value = UCase(Me.txt_New_Asset.Value): nrA.Range(1, 3).Value = UCase(Me.txt_New_Asset.Value): nrA.Range(1, 4).Value = Me.cmb_New_Class.Value: nrA.Range(1, 5).Value = MOD_00_WMS_Architecture.Obtenir_Parametre(""SYS_DEVISE_BASE"", ""USD""): nrA.Range(1, 6).Value = ""-""": i = i + 1
    L(i) = "        wsA.Protect ""WMS_ADMIN_2026"", UserInterfaceOnly:=True": i = i + 1
    L(i) = "    End If": i = i + 1

    L(i) = "    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets(""FACT_Trade"")": i = i + 1
    L(i) = "    wsF.Unprotect ""WMS_ADMIN_2026""": i = i + 1
    L(i) = "    Dim nR As ListRow: Set nR = wsF.ListObjects(""T_FACT_Trade"").ListRows.Add": i = i + 1
    L(i) = "    nR.Range(1, 1).Value = MOD_00_WMS_Architecture.GENERER_ID(""T_FACT_Trade"")": i = i + 1
    L(i) = "    nR.Range(1, 2).Value = DateSerial(CInt(dParts(2)), CInt(dParts(0)), CInt(dParts(1)))": i = i + 1
    L(i) = "    nR.Range(1, 3).Value = idP": i = i + 1
    L(i) = "    nR.Range(1, 4).Value = idA": i = i + 1
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
