VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Transaction 
   Caption         =   "UserForm1"
   ClientHeight    =   9425.001
   ClientLeft      =   117
   ClientTop       =   455
   ClientWidth     =   5369
   OleObjectBlob   =   "USF_Transaction_V1.0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function TR(Clé As String) As String
    TR = MOD_02_AppHome_Global.TR(Clé)
End Function
Private Sub UserForm_Initialize()
    MOD_03_Gatekeeper.Garantir_Lexique_Formulaire
    Me.Caption = TR("FRM_TITLE")
    Me.lbl_txt_Date.Caption = TR("FRM_DATE")
    Me.lbl_cmb_Compte.Caption = TR("FRM_COMPTE")
    Me.lbl_cmb_Categorie.Caption = TR("FRM_CAT")
    Me.lbl_cmb_Tiers.Caption = TR("FRM_TIERS")
    Me.lbl_txt_Montant.Caption = TR("FRM_MONTANT")
    Me.lbl_cmb_Devise.Caption = TR("FRM_DEVISE")
    Me.lbl_txt_Description.Caption = TR("FRM_DESC")
    Me.lbl_txt_New_Compte.Caption = TR("FRM_NEW")
    Me.lbl_txt_New_Categorie.Caption = TR("FRM_NEW")
    Me.lbl_txt_New_Tiers.Caption = TR("FRM_NEW")
    Me.btn_Save.Caption = TR("FRM_SAVE")
    Me.btn_Cancel.Caption = TR("FRM_CANCEL")
    Me.txt_Date.ControlTipText = TR("TT_F_DATE")
    Me.cmb_Compte.ControlTipText = TR("TT_F_COMPTE")
    Me.cmb_Categorie.ControlTipText = TR("TT_F_CAT")
    Me.cmb_Tiers.ControlTipText = TR("TT_F_TIERS")
    Me.txt_Montant.ControlTipText = TR("TT_F_MONTANT")
    Me.cmb_Devise.ControlTipText = TR("TT_F_DEVISE")
    Me.txt_Description.ControlTipText = TR("TT_F_DESC")
    Me.btn_Save.ControlTipText = TR("TT_F_SAVE")
    Me.btn_Cancel.ControlTipText = TR("TT_F_CANCEL")
    Me.cmb_New_Cpt_Type.ControlTipText = TR("TT_F_TYPE_CPT")
    Me.cmb_New_Cat_Type.ControlTipText = TR("TT_F_TYPE_CAT")
    Me.txt_Date.Value = Format(Date, "mm/dd/yyyy")
    Me.cmb_Devise.List = MOD_01_CoreEngine.GET_TAUX_CHANGE().keys()
    Me.cmb_Devise.Value = MOD_06_Budget_ZBB.Obtenir_Parametre("SYS_DEVISE_BASE", "MUR")
    Me.cmb_New_Cpt_Type.List = Array(TR("OPT_LIQ"), TR("OPT_INV"), TR("OPT_DET"))
    Me.cmb_New_Cpt_Type.ListIndex = 0
    Me.cmb_New_Cat_Type.List = Array(TR("OPT_DEP"), TR("OPT_REV"), TR("OPT_TRA"))
    Me.cmb_New_Cat_Type.ListIndex = 0
    Charger_Combo Me.cmb_Compte, "T_DIM_Compte"
    Charger_Combo Me.cmb_Categorie, "T_DIM_Categorie"
    Charger_Combo Me.cmb_Tiers, "T_DIM_Tiers"
End Sub
Private Sub Charger_Combo(cmb As MSForms.ComboBox, NomTable As String)
    Dim tbl As ListObject, k As Long
    On Error Resume Next: Set tbl = ThisWorkbook.Sheets(Split(NomTable, "_", 2)(1)).ListObjects(NomTable): On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    cmb.Clear: cmb.ColumnCount = 2: cmb.ColumnWidths = "0 pt;200 pt"
    If tbl.ListRows.Count > 0 Then
        For k = 1 To tbl.ListRows.Count
            If Trim(tbl.ListRows(k).Range(1, 2).Value) <> "" Then
                cmb.AddItem tbl.ListRows(k).Range(1, 1).Value
                cmb.List(cmb.ListCount - 1, 1) = TR(CStr(tbl.ListRows(k).Range(1, 2).Value))
            End If
        Next k
    End If
End Sub
Private Sub cmb_Compte_Change(): Gerer_Visibilite_Double Me.cmb_Compte, Me.txt_New_Compte, Me.cmb_New_Cpt_Type, Me.lbl_txt_New_Compte: End Sub
Private Sub cmb_Categorie_Change(): Gerer_Visibilite_Double Me.cmb_Categorie, Me.txt_New_Categorie, Me.cmb_New_Cat_Type, Me.lbl_txt_New_Categorie: Update_Tiers_Mode: End Sub
Private Sub cmb_New_Cat_Type_Change(): Update_Tiers_Mode: End Sub
Private Sub cmb_Tiers_Change(): Gerer_Visibilite_Simple Me.cmb_Tiers, Me.txt_New_Tiers, Me.lbl_txt_New_Tiers: End Sub
Private Sub Update_Tiers_Mode()
    Dim t As String: t = ""
    If Me.txt_New_Categorie.Visible Then
        Select Case Me.cmb_New_Cat_Type.ListIndex: Case 0: t = "DEPENSE": Case 1: t = "REVENU": Case 2: t = "TRANSFERT": End Select
    ElseIf Me.cmb_Categorie.ListIndex <> -1 Then
        Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("DIM_Categorie")
        Dim r As Long: For r = 1 To wsC.ListObjects("T_DIM_Categorie").ListRows.Count
            If CStr(wsC.ListObjects("T_DIM_Categorie").DataBodyRange(r, 1).Value) = CStr(Me.cmb_Categorie.List(Me.cmb_Categorie.ListIndex, 0)) Then t = UCase(Trim(wsC.ListObjects("T_DIM_Categorie").DataBodyRange(r, 3).Value)): Exit For
        Next r
    End If
    If t = "TRANSFERT" Then
        Me.lbl_cmb_Tiers.Caption = TR("FRM_DEST")
        Charger_Combo Me.cmb_Tiers, "T_DIM_Compte"
        Me.lbl_txt_New_Tiers.Visible = False: Me.txt_New_Tiers.Visible = False: Me.txt_New_Tiers.Value = ""
    Else
        Me.lbl_cmb_Tiers.Caption = TR("FRM_TIERS")
        Charger_Combo Me.cmb_Tiers, "T_DIM_Tiers"
        Gerer_Visibilite_Simple Me.cmb_Tiers, Me.txt_New_Tiers, Me.lbl_txt_New_Tiers
    End If
End Sub
Private Sub Gerer_Visibilite_Double(cmb As MSForms.ComboBox, txt As MSForms.TextBox, cmbType As MSForms.ComboBox, lbl As MSForms.Label)
    Dim estAutre As Boolean: estAutre = (cmb.Text = TR("Autre (Préciser...)") Or InStr(1, cmb.Text, "Autre", vbTextCompare) > 0 Or InStr(1, cmb.Text, "Other", vbTextCompare) > 0)
    txt.Visible = estAutre: lbl.Visible = estAutre: cmbType.Visible = estAutre
    If Not estAutre Then txt.Value = ""
End Sub
Private Sub Gerer_Visibilite_Simple(cmb As MSForms.ComboBox, txt As MSForms.TextBox, lbl As MSForms.Label)
    Dim estAutre As Boolean: estAutre = (cmb.Text = TR("Autre (Préciser...)") Or InStr(1, cmb.Text, "Autre", vbTextCompare) > 0 Or InStr(1, cmb.Text, "Other", vbTextCompare) > 0)
    txt.Visible = estAutre: lbl.Visible = estAutre
    If Not estAutre Then txt.Value = ""
End Sub
Private Function Obtenir_ID(cmb As MSForms.ComboBox, txt As MSForms.TextBox, NomTable As String, TypeSelect As String) As Long
    If txt.Visible = False Then Obtenir_ID = CLng(cmb.List(cmb.ListIndex, 0)): Exit Function
    Dim valClean As String: valClean = MOD_01_CoreEngine.CLEAN_TEXT(txt.Value)
    Dim tbl As ListObject, ws As Worksheet, k As Long
    Set ws = ThisWorkbook.Sheets(Split(NomTable, "_", 2)(1)): Set tbl = ws.ListObjects(NomTable)
    For k = 1 To tbl.ListRows.Count
        If UCase(Trim(tbl.ListRows(k).Range(1, 2).Value)) = UCase(valClean) Then Obtenir_ID = tbl.ListRows(k).Range(1, 1).Value: Exit Function
    Next k
    ws.Unprotect "SFP_ADMIN_2026"
    Dim newRow As ListRow: Set newRow = tbl.ListRows.Add
    Dim newID As Long: newID = MOD_01_CoreEngine.GENERER_NOUVEL_ID(NomTable)
    newRow.Range(1, 1).Value = newID
    newRow.Range(1, 2).Value = valClean
    newRow.Range(1, 3).Value = TypeSelect
    If NomTable = "T_DIM_Compte" Then
        newRow.Range(1, 4).Value = Me.cmb_Devise.Value
        newRow.Range(1, 5).Value = "OUI"
    End If
    ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    Obtenir_ID = newID
End Function
Private Sub btn_Save_Click()
    If Me.cmb_Compte.ListIndex = -1 Or Me.cmb_Categorie.ListIndex = -1 Or Me.cmb_Tiers.ListIndex = -1 Then MsgBox TR("MSG_ERR_MISSING"), vbCritical: Exit Sub
    Dim strMontant As String, dblMontant As Double
    strMontant = Replace(Me.txt_Montant.Value, ",", ".")
    dblMontant = Val(strMontant)
    If dblMontant = 0 Then MsgBox TR("MSG_ERR_AMT"), vbCritical: Exit Sub
    Dim dParts() As String: dParts = Split(Replace(Me.txt_Date.Value, "-", "/"), "/")
    If UBound(dParts) <> 2 Then MsgBox TR("MSG_ERR_MISSING"), vbCritical: Exit Sub
    If Not IsNumeric(dParts(0)) Or Not IsNumeric(dParts(1)) Or Not IsNumeric(dParts(2)) Then MsgBox TR("MSG_ERR_MISSING"), vbCritical: Exit Sub
    Dim idC As Long, idCat As Long, idT As Long
    Dim rawCpt As String: Select Case Me.cmb_New_Cpt_Type.ListIndex: Case 0: rawCpt = "LIQUIDITE": Case 1: rawCpt = "INVESTISSEMENT": Case 2: rawCpt = "DETTE": End Select
    Dim rawCat As String: Select Case Me.cmb_New_Cat_Type.ListIndex: Case 0: rawCat = "DEPENSE": Case 1: rawCat = "REVENU": Case 2: rawCat = "TRANSFERT": End Select
    idC = Obtenir_ID(Me.cmb_Compte, Me.txt_New_Compte, "T_DIM_Compte", rawCpt)
    idCat = Obtenir_ID(Me.cmb_Categorie, Me.txt_New_Categorie, "T_DIM_Categorie", rawCat)
    idT = Obtenir_ID(Me.cmb_Tiers, Me.txt_New_Tiers, "T_DIM_Tiers", "AUTRE")
    Dim wsFact As Worksheet: Set wsFact = ThisWorkbook.Sheets("FACT_Transaction")
    Dim tblFact As ListObject: Set tblFact = wsFact.ListObjects("T_FACT_Transaction")
    Dim initRows As Long: initRows = tblFact.ListRows.Count
    On Error GoTo ROLLBACK_TRAN
    Dim typeF As String, rC As Long
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("DIM_Categorie")
    typeF = "AUTRE"
    For rC = 1 To wsC.ListObjects("T_DIM_Categorie").ListRows.Count
        If CStr(wsC.ListObjects("T_DIM_Categorie").DataBodyRange(rC, 1).Value) = CStr(idCat) Then typeF = UCase(Trim(wsC.ListObjects("T_DIM_Categorie").DataBodyRange(rC, 3).Value)): Exit For
    Next rC
    wsFact.Unprotect "SFP_ADMIN_2026"
    Dim nr As ListRow: Set nr = tblFact.ListRows.Add
    nr.Range(1, 1).Value = MOD_01_CoreEngine.GENERER_NOUVEL_ID("T_FACT_Transaction")
    nr.Range(1, 2).Value = DateSerial(CInt(dParts(2)), CInt(dParts(0)), CInt(dParts(1)))
    nr.Range(1, 3).Value = idC: nr.Range(1, 4).Value = idCat: nr.Range(1, 5).Value = idT
    nr.Range(1, 6).Value = dblMontant
    nr.Range(1, 7).Value = Me.cmb_Devise.Value
    nr.Range(1, 8).Value = MOD_01_CoreEngine.CLEAN_TEXT(Me.txt_Description.Value)
    nr.Range(1, 9).Value = Application.UserName: nr.Range(1, 10).Value = Now
    If typeF = "TRANSFERT" Then
        If idC = idT Then Err.Raise vbObjectError + 1, "", "Le compte source et destination doivent être différents.": Exit Sub
        nr.Range(1, 3).Value = idC: nr.Range(1, 4).Value = idCat: nr.Range(1, 5).Value = idT
        nr.Range(1, 6).Value = -dblMontant
        nr.Range(1, 7).Value = Me.cmb_Devise.Value: nr.Range(1, 8).Value = MOD_01_CoreEngine.CLEAN_TEXT(Me.txt_Description.Value): nr.Range(1, 9).Value = Application.UserName: nr.Range(1, 10).Value = Now
        Dim nR2 As ListRow: Set nR2 = tblFact.ListRows.Add
        nR2.Range(1, 1).Value = MOD_01_CoreEngine.GENERER_NOUVEL_ID("T_FACT_Transaction")
        nR2.Range(1, 2).Value = DateSerial(CInt(dParts(2)), CInt(dParts(0)), CInt(dParts(1)))
        nR2.Range(1, 3).Value = idT: nR2.Range(1, 4).Value = idCat: nR2.Range(1, 5).Value = idC
        nR2.Range(1, 6).Value = dblMontant
        nR2.Range(1, 7).Value = Me.cmb_Devise.Value: nR2.Range(1, 8).Value = MOD_01_CoreEngine.CLEAN_TEXT(Me.txt_Description.Value): nR2.Range(1, 9).Value = Application.UserName: nR2.Range(1, 10).Value = Now
    Else
        nr.Range(1, 3).Value = idC: nr.Range(1, 4).Value = idCat: nr.Range(1, 5).Value = idT
        nr.Range(1, 6).Value = dblMontant
        nr.Range(1, 7).Value = Me.cmb_Devise.Value: nr.Range(1, 8).Value = MOD_01_CoreEngine.CLEAN_TEXT(Me.txt_Description.Value): nr.Range(1, 9).Value = Application.UserName: nr.Range(1, 10).Value = Now
    End If
    wsFact.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    MsgBox TR("MSG_OK"), vbInformation: Unload Me
    Exit Sub
ROLLBACK_TRAN:
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    Dim currRows As Long: currRows = tblFact.ListRows.Count
    Dim rIdx As Long
    For rIdx = currRows To initRows + 1 Step -1
        tblFact.ListRows(rIdx).Delete
    Next rIdx
    wsFact.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    MsgBox "ERREUR CRITIQUE : TRANSACTION ANNULÉE (ROLLBACK)." & vbCrLf & vbCrLf & "Détail : " & errMsg, vbCritical, "ACID Rollback"
    Unload Me
End Sub
Private Sub btn_Cancel_Click(): Unload Me: End Sub
