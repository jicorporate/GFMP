VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_Trade 
   Caption         =   "Ordre de Bourse (Trade Entry)"
   ClientHeight    =   10426
   ClientLeft      =   117
   ClientTop       =   455
   ClientWidth     =   5772
   OleObjectBlob   =   "USF_Trade_V1.0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_Trade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function TR(Cle As String) As String: TR = MOD_02_WMS_Hub.TR(Cle): End Function
Private Sub UserForm_Initialize()
    Me.Caption = TR("CARD_T_T")
    Me.txt_Date.Value = Format(Date, "mm/dd/yyyy")
    Me.txt_FXRate.Value = "1"
    Me.txt_Fees.Value = "0"
    Me.cmb_Type.List = Array("ACHAT", "VENTE", "DIVIDENDE", "SPLIT")
    Me.cmb_Type.ListIndex = 0
    Me.cmb_New_Class.List = Array("ACTION", "ETF", "CRYPTO", "OBLIGATION")
    Me.cmb_New_Class.ListIndex = 0
    Me.cmb_Portfolio.ColumnCount = 2: Me.cmb_Portfolio.ColumnWidths = "0 pt;200 pt"
    Me.cmb_Asset.ColumnCount = 2: Me.cmb_Asset.ColumnWidths = "0 pt;200 pt"
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("DIM_Portfolio")
    Dim r As Long: For r = 1 To wsP.ListObjects("T_DIM_Portfolio").ListRows.Count
        Me.cmb_Portfolio.AddItem wsP.ListObjects("T_DIM_Portfolio").DataBodyRange(r, 1).Value
        Dim pName As String: pName = CStr(wsP.ListObjects("T_DIM_Portfolio").DataBodyRange(r, 2).Value)
        Dim pBroker As String: pBroker = CStr(wsP.ListObjects("T_DIM_Portfolio").DataBodyRange(r, 3).Value)
        Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListCount - 1, 1) = pName & " (" & pBroker & ")"
    Next r
    Me.cmb_Portfolio.AddItem "AUTRE"
    Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListCount - 1, 1) = TR("Autre (Préciser...)")
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("DIM_Asset")
    For r = 1 To wsA.ListObjects("T_DIM_Asset").ListRows.Count
        Me.cmb_Asset.AddItem wsA.ListObjects("T_DIM_Asset").DataBodyRange(r, 1).Value
        Dim aTicker As String: aTicker = CStr(wsA.ListObjects("T_DIM_Asset").DataBodyRange(r, 2).Value)
        Dim aName As String: aName = CStr(wsA.ListObjects("T_DIM_Asset").DataBodyRange(r, 3).Value)
        Me.cmb_Asset.List(Me.cmb_Asset.ListCount - 1, 1) = aTicker & " - " & aName
    Next r
    Me.cmb_Asset.AddItem "AUTRE"
    Me.cmb_Asset.List(Me.cmb_Asset.ListCount - 1, 1) = TR("Autre (Préciser...)")
End Sub
Private Sub cmb_Portfolio_Change()
    Dim isOther As Boolean: isOther = False
    If Me.cmb_Portfolio.ListIndex <> -1 Then isOther = (CStr(Me.cmb_Portfolio.List(Me.cmb_Portfolio.ListIndex, 0)) = "AUTRE")
    Me.txt_New_Port.Visible = isOther: Me.txt_New_Broker.Visible = isOther: Me.lbl_txt_New_Port.Visible = isOther
End Sub
Private Sub cmb_Asset_Change()
    Dim isOther As Boolean: isOther = False
    If Me.cmb_Asset.ListIndex <> -1 Then isOther = (CStr(Me.cmb_Asset.List(Me.cmb_Asset.ListIndex, 0)) = "AUTRE")
    Me.txt_New_Asset.Visible = isOther: Me.cmb_New_Class.Visible = isOther: Me.lbl_txt_New_Asset.Visible = isOther
End Sub
Private Sub btn_Save_Click()
    If Me.cmb_Portfolio.ListIndex = -1 Or Me.cmb_Asset.ListIndex = -1 Then MsgBox "Sélection Incomplète.", vbCritical: Exit Sub
    Dim dStr As String: dStr = Replace(Replace(Replace(Me.txt_Date.Value, "-", "/"), ".", "/"), " ", "")
    Dim dParts() As String: dParts = Split(dStr, "/")
    If UBound(dParts) <> 2 Then MsgBox "Date invalide.", vbCritical: Exit Sub
    If Not IsNumeric(dParts(0)) Or Not IsNumeric(dParts(1)) Or Not IsNumeric(dParts(2)) Then MsgBox "Date invalide.", vbCritical: Exit Sub
    Dim qty As Double, px As Double, fees As Double, fx As Double
    qty = Val(Replace(Me.txt_Qty.Value, ",", ".")): px = Val(Replace(Me.txt_Price.Value, ",", "."))
    fees = Val(Replace(Me.txt_Fees.Value, ",", ".")): fx = Val(Replace(Me.txt_FXRate.Value, ",", "."))
    If qty <= 0 Or px <= 0 Or fx <= 0 Then MsgBox "Quantité, Prix et Taux doivent être > 0.", vbCritical: Exit Sub
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("DIM_Portfolio")
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("DIM_Asset")
    Dim idP As Variant: idP = Me.cmb_Portfolio.Value
    Dim idA As Variant: idA = Me.cmb_Asset.Value
    If idP = "AUTRE" Then
        wsP.Unprotect "WMS_ADMIN_2026"
        Dim nrP As ListRow: Set nrP = wsP.ListObjects("T_DIM_Portfolio").ListRows.Add
        idP = MOD_00_WMS_Architecture.GENERER_ID("T_DIM_Portfolio")
        nrP.Range(1, 1).Value = idP: nrP.Range(1, 2).Value = Me.txt_New_Port.Value: nrP.Range(1, 3).Value = Me.txt_New_Broker.Value: nrP.Range(1, 4).Value = MOD_00_WMS_Architecture.Obtenir_Parametre("SYS_DEVISE_BASE", "USD"): nrP.Range(1, 5).Value = "OUI"
        wsP.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    End If
    If idA = "AUTRE" Then
        wsA.Unprotect "WMS_ADMIN_2026"
        Dim nrA As ListRow: Set nrA = wsA.ListObjects("T_DIM_Asset").ListRows.Add
        idA = MOD_00_WMS_Architecture.GENERER_ID("T_DIM_Asset")
        nrA.Range(1, 1).Value = idA: nrA.Range(1, 2).Value = UCase(Me.txt_New_Asset.Value): nrA.Range(1, 3).Value = UCase(Me.txt_New_Asset.Value): nrA.Range(1, 4).Value = Me.cmb_New_Class.Value: nrA.Range(1, 5).Value = MOD_00_WMS_Architecture.Obtenir_Parametre("SYS_DEVISE_BASE", "USD"): nrA.Range(1, 6).Value = "-"
        wsA.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    End If
    Dim wsF As Worksheet: Set wsF = ThisWorkbook.Sheets("FACT_Trade")
    wsF.Unprotect "WMS_ADMIN_2026"
    Dim nr As ListRow: Set nr = wsF.ListObjects("T_FACT_Trade").ListRows.Add
    nr.Range(1, 1).Value = MOD_00_WMS_Architecture.GENERER_ID("T_FACT_Trade")
    nr.Range(1, 2).Value = DateSerial(CInt(dParts(2)), CInt(dParts(0)), CInt(dParts(1)))
    nr.Range(1, 3).Value = idP
    nr.Range(1, 4).Value = idA
    nr.Range(1, 5).Value = Me.cmb_Type.Value
    nr.Range(1, 6).Value = qty: nr.Range(1, 7).Value = px: nr.Range(1, 8).Value = fees: nr.Range(1, 9).Value = fx
    nr.Range(1, 10).Value = Now
    wsF.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    MsgBox "Ordre de Bourse enregistré avec succès !", vbInformation: Unload Me
End Sub
Private Sub btn_Cancel_Click(): Unload Me: End Sub
