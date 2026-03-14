VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Budget 
   Caption         =   "Allocation Budgétaire"
   ClientHeight    =   5226
   ClientLeft      =   117
   ClientTop       =   455
   ClientWidth     =   4771
   OleObjectBlob   =   "USF_Budget_V1.0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Budget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    Me.txt_Mois.Value = MOD_06_Budget_ZBB.Obtenir_Parametre("BUDG_FILTRE_MOIS", Format(Date, "yyyy-mm"))
    Me.Caption = MOD_02_AppHome_Global.TR("BTN_ALLOC")
    Me.lbl_Mois.Caption = MOD_02_AppHome_Global.TR("FRM_B_MOIS")
    Me.lbl_Cat.Caption = MOD_02_AppHome_Global.TR("FRM_B_CAT")
    Me.lbl_Devise.Caption = MOD_02_AppHome_Global.TR("FRM_B_DEV")
    Me.lbl_Montant.Caption = MOD_02_AppHome_Global.TR("FRM_B_AMT")
    Me.btn_Save.Caption = MOD_02_AppHome_Global.TR("FRM_B_SAVE")
    Me.btn_Cancel.Caption = MOD_02_AppHome_Global.TR("FRM_B_CANCEL")
    Me.cmb_Devise.List = MOD_01_CoreEngine.GET_TAUX_CHANGE().keys()
    Me.cmb_Devise.Value = MOD_06_Budget_ZBB.Obtenir_Parametre("SYS_DEVISE_BASE", "MUR")
    Dim tbl As ListObject, i As Long
    On Error Resume Next: Set tbl = ThisWorkbook.Sheets("DIM_Categorie").ListObjects("T_DIM_Categorie"): On Error GoTo 0
    If Not tbl Is Nothing Then
        Me.cmb_Cat.ColumnCount = 2: Me.cmb_Cat.ColumnWidths = "0 pt;150 pt"
        For i = 1 To tbl.ListRows.Count
            If UCase(Trim(tbl.DataBodyRange(i, 3).Value)) = "DEPENSE" Then
                Me.cmb_Cat.AddItem tbl.DataBodyRange(i, 1).Value
                Me.cmb_Cat.List(Me.cmb_Cat.ListCount - 1, 1) = MOD_02_AppHome_Global.TR(CStr(tbl.DataBodyRange(i, 2).Value))
            End If
        Next i
    End If
End Sub
Private Sub btn_Save_Click()
    If Me.cmb_Cat.ListIndex = -1 Then MsgBox "Sélectionnez une catégorie.", vbCritical: Exit Sub
    Dim m As String: m = Replace(Me.txt_Montant.Value, ",", ".")
    If Val(m) <= 0 Then MsgBox "Montant invalide.", vbCritical: Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("FACT_Budget")
    ws.Unprotect "SFP_ADMIN_2026"
    Dim tbl As ListObject: Set tbl = ws.ListObjects("T_FACT_Budget")
    If tbl.ListColumns.Count < 7 Then tbl.ListColumns.Add.Name = "Devise"
    Dim i As Long, found As Boolean, idCat As String, targetMois As String
    found = False: idCat = Me.cmb_Cat.List(Me.cmb_Cat.ListIndex, 0): targetMois = Trim(Me.txt_Mois.Value)
    If tbl.ListRows.Count > 0 Then
        For i = 1 To tbl.ListRows.Count
            If CStr(tbl.DataBodyRange(i, 2).Value) = targetMois And CStr(tbl.DataBodyRange(i, 3).Value) = idCat Then
                tbl.DataBodyRange(i, 4).Value = Val(m): tbl.DataBodyRange(i, 6).Value = Now: tbl.DataBodyRange(i, 7).Value = Me.cmb_Devise.Value
                found = True: Exit For
            End If
        Next i
    End If
    If Not found Then
        Dim nr As ListRow: Set nr = tbl.ListRows.Add
        nr.Range(1, 1).Value = MOD_01_CoreEngine.GENERER_NOUVEL_ID("T_FACT_Budget")
        nr.Range(1, 2).Value = targetMois: nr.Range(1, 3).Value = idCat
        nr.Range(1, 4).Value = Val(m): nr.Range(1, 5).Value = Application.UserName: nr.Range(1, 6).Value = Now: nr.Range(1, 7).Value = Me.cmb_Devise.Value
    End If
    ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    Unload Me: MOD_06_Budget_ZBB.Rafraichir_Budget
End Sub
    Me.cmb_Devise.List = MOD_01_CoreEngine.GET_TAUX_CHANGE().keys()
    Me.cmb_Devise.Value = MOD_06_Budget_ZBB.Obtenir_Parametre("SYS_DEVISE_BASE", "MUR")
    Me.lbl_Devise.Caption = MOD_02_AppHome_Global.TR("FRM_B_DEV")
End Sub
Private Sub btn_Cancel_Click(): Unload Me: End Sub
