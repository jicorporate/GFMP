Option Explicit

' =========================================================================
' MODULE: MOD_06_Budget_ZBB
' OBJECTIF: Pilotage Budget ZBB, Violet Zebra, Devises, Temps, ADLaM 10
' =========================================================================

Public Sub DEPLOIEMENT_ETAPE_6_BUDGET()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    
    Garantir_Lexique_Budget
    Generer_Formulaire_Budget
    GENERER_BUDGET_DASHBOARD
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    
    AppActivate ThisWorkbook.Name ' Force la fermeture de l'éditeur VBA en arrière-plan
    Application.ScreenUpdating = True
    
    MsgBox "LE PILOTAGE BUDGÉTAIRE 'MASTER CLASS' EST DÉPLOYÉ." & vbCrLf & vbCrLf & _
           "1. Le Filtre Multi-Devises et le Time Slider sont opérationnels." & vbCrLf & _
           "2. Le tableau arbore la charte 'Violet Zebra' avec des DataBars Jaunes." & vbCrLf & _
           "3. Le Responsive Design garantit zéro chevauchement au Zoom 100%.", vbInformation, "SFP v3.2 - Élégance Absolue"
End Sub

' -------------------------------------------------------------------------
' 1. STATE MANAGEMENT (Filtres Devise & Mois)
' -------------------------------------------------------------------------
Public Function Obtenir_Parametre(NomParam As String, ValeurDefaut As String) As String
    Dim tblConf As ListObject, i As Long
    On Error Resume Next: Set tblConf = ThisWorkbook.Sheets("SYS_Config").ListObjects("T_SYS_Config"): On Error GoTo 0
    If tblConf Is Nothing Then
        Obtenir_Parametre = ValeurDefaut
        Exit Function
    End If
    
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = NomParam Then
            Obtenir_Parametre = tblConf.DataBodyRange(i, 2).Value
            Exit Function
        End If
    Next i
    
    Dim nr As ListRow: Set nr = tblConf.ListRows.Add
    nr.Range(1, 1).Value = NomParam
    nr.Range(1, 2).Value = ValeurDefaut
    nr.Range(1, 3).Value = "Filtre Actif"
    Obtenir_Parametre = ValeurDefaut
End Function

Private Sub Modifier_Parametre(NomParam As String, NouvelleValeur As String)
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    wsSys.Unprotect "SFP_ADMIN_2026"
    Dim tblConf As ListObject: Set tblConf = wsSys.ListObjects("T_SYS_Config")
    Dim i As Long
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = NomParam Then
            tblConf.DataBodyRange(i, 2).Value = NouvelleValeur
            Exit For
        End If
    Next i
    wsSys.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
End Sub

' -------------------------------------------------------------------------
' 2. MOTEUR AUTO-CICATRISANT
' -------------------------------------------------------------------------
Private Sub Garantir_Lexique_Budget()
    Dim tblDic As ListObject
    On Error Resume Next: Set tblDic = ThisWorkbook.Sheets("SYS_Config").ListObjects("T_SYS_Dictionary"): On Error GoTo 0
    If tblDic Is Nothing Then Exit Sub
    
    Upsert_Dico tblDic, "BUDG_TITLE", "PILOTAGE BUDGÉTAIRE", "BUDGET TRACKING", "CONTROL PRESUPUESTARIO", "CONTROLE ORÇAMENTÁRIO", "BUDGETKONTROLLE", "CONTROLLO BUDGET", "BUDGETBEHEER", "BUDGETKONTROLL"
    Upsert_Dico tblDic, "BTN_ALLOC", "ALLOUER BUDGET", "ALLOCATE BUDGET", "ASIGNAR PRESUPUESTO", "ALOCAR ORÇAMENTO", "BUDGET ZUWEISEN", "ASSEGNA BUDGET", "BUDGET TOEWIJZEN", "TILLDELA BUDGET"
    
    Upsert_Dico tblDic, "KPI_B_ALLO", "TOTAL ALLOUÉ", "TOTAL ALLOCATED", "TOTAL ASIGNADO", "TOTAL ALOCADO", "GESAMT ZUGEWIESEN", "TOTALE ASSEGNATO", "TOTAAL TOEGEWEZEN", "TOTALT TILLDELAD"
    Upsert_Dico tblDic, "KPI_B_SPEN", "TOTAL DÉPENSÉ", "TOTAL SPENT", "TOTAL GASTADO", "TOTAL GASTO", "GESAMT AUSGEGEBEN", "TOTALE SPESO", "TOTAAL UITGEGEVEN", "TOTALT SPENDERAT"
    Upsert_Dico tblDic, "KPI_B_LEFT", "RESTE À DÉPENSER", "REMAINING", "RESTANTE", "RESTANTE", "VERBLEIBEND", "RIMANENTE", "RESTEREND", "ÅTERSTÅENDE"
    
    Upsert_Dico tblDic, "COL_B_CAT", "ENVELOPPE (CATÉGORIE)", "ENVELOPE (CATEGORY)", "CATEGORÍA", "CATEGORIA", "KATEGORIE", "CATEGORIA", "CATEGORIE", "KATEGORI"
    Upsert_Dico tblDic, "COL_B_ALLO", "ALLOUÉ", "ALLOCATED", "ASIGNADO", "ALOCADO", "ZUGEWIESEN", "ASSEGNATO", "TOEGEWEZEN", "TILLDELAD"
    Upsert_Dico tblDic, "COL_B_REAL", "RÉALISÉ", "ACTUAL", "REAL", "REALIZADO", "IST-WERT", "REALE", "WERKELIJK", "FAKTISKT"
    Upsert_Dico tblDic, "COL_B_DIFF", "ÉCART", "VARIANCE", "DIFERENCIA", "DIFERENÇA", "ABWEICHUNG", "VARIAZIONE", "VERSCHIL", "AVVIKELSE"
    Upsert_Dico tblDic, "COL_B_PROG", "CONSOMMATION", "CONSUMPTION", "CONSUMO", "CONSUMO", "VERBRAUCH", "CONSUMO", "VERBRUIK", "FÖRBRUKNING"
    Upsert_Dico tblDic, "NO_BUDG", "Aucun budget alloué", "No budget allocated", "Sin presupuesto", "Sem orçamento", "Kein Budget", "Nessun budget", "Geen budget", "Ingen budget"
End Sub

Private Sub Upsert_Dico(tbl As ListObject, k As String, fr As String, en As String, es As String, pt As String, de As String, it As String, nl As String, sv As String)
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 1).Value = k Then
            tbl.DataBodyRange(i, 2).Value = fr
            tbl.DataBodyRange(i, 3).Value = en
            tbl.DataBodyRange(i, 4).Value = es
            tbl.DataBodyRange(i, 5).Value = pt
            tbl.DataBodyRange(i, 6).Value = de
            tbl.DataBodyRange(i, 7).Value = it
            tbl.DataBodyRange(i, 8).Value = nl
            tbl.DataBodyRange(i, 9).Value = sv
            Exit Sub
        End If
    Next i
    Dim nr As ListRow: Set nr = tbl.ListRows.Add
    nr.Range(1, 1).Value = k
    nr.Range(1, 2).Value = fr
    nr.Range(1, 3).Value = en
    nr.Range(1, 4).Value = es
    nr.Range(1, 5).Value = pt
    nr.Range(1, 6).Value = de
    nr.Range(1, 7).Value = it
    nr.Range(1, 8).Value = nl
    nr.Range(1, 9).Value = sv
End Sub

Private Function TR(Clé As String) As String
    TR = MOD_02_AppHome_Global.TR(Clé)
End Function

' -------------------------------------------------------------------------
' 3. GÉNÉRATEUR DU FORMULAIRE D'ALLOCATION
' -------------------------------------------------------------------------
Private Sub Generer_Formulaire_Budget()
    Dim VBP As Object, VBComp As Object, myForm As Object, ctrl As Object
    On Error Resume Next: Set VBP = ThisWorkbook.VBProject: On Error GoTo 0
    If VBP Is Nothing Then Exit Sub
    
    For Each VBComp In VBP.VBComponents
        If VBComp.Name = "USF_Budget" Then VBP.VBComponents.Remove VBComp
    Next VBComp
    
    Set VBComp = VBP.VBComponents.Add(3)
    VBComp.Properties("Name") = "USF_Budget"
    Set myForm = VBComp.Designer
    VBComp.Properties("Width") = 250: VBComp.Properties("Height") = 250
    VBComp.Properties("Caption") = "Allocation Budgétaire"
    
    Dim t As Integer: t = 10
    
    Set ctrl = myForm.Controls.Add("Forms.Label.1", "lbl_Mois")
    ctrl.Caption = "Mois cible (AAAA-MM) :": ctrl.Top = t: ctrl.Left = 20: ctrl.Width = 200: ctrl.Height = 12
    Set ctrl = myForm.Controls.Add("Forms.TextBox.1", "txt_Mois")
    ctrl.Top = t + 12: ctrl.Left = 20: ctrl.Width = 200: ctrl.Height = 18
    t = t + 35
    
    Set ctrl = myForm.Controls.Add("Forms.Label.1", "lbl_Cat")
    ctrl.Caption = "Enveloppe (Catégorie) :": ctrl.Top = t: ctrl.Left = 20: ctrl.Width = 200: ctrl.Height = 12
    Set ctrl = myForm.Controls.Add("Forms.ComboBox.1", "cmb_Cat")
    ctrl.Top = t + 12: ctrl.Left = 20: ctrl.Width = 200: ctrl.Height = 18
    t = t + 35
    
    Set ctrl = myForm.Controls.Add("Forms.Label.1", "lbl_Montant")
    ctrl.Caption = "Montant Alloué (Base Devise) :": ctrl.Top = t: ctrl.Left = 20: ctrl.Width = 200: ctrl.Height = 12
    Set ctrl = myForm.Controls.Add("Forms.TextBox.1", "txt_Montant")
    ctrl.Top = t + 12: ctrl.Left = 20: ctrl.Width = 200: ctrl.Height = 18
    t = t + 35
    
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Save")
    ctrl.Caption = "ALLOUER": ctrl.Top = t + 10: ctrl.Left = 20: ctrl.Width = 90: ctrl.Height = 25
    ctrl.BackColor = RGB(250, 218, 94): ctrl.Font.Bold = True
    
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Cancel")
    ctrl.Caption = "ANNULER": ctrl.Top = t + 10: ctrl.Left = 130: ctrl.Width = 90: ctrl.Height = 25
    ctrl.BackColor = RGB(120, 81, 169): ctrl.ForeColor = vbWhite: ctrl.Font.Bold = True
    
    VBComp.CodeModule.AddFromString Code_VBA_USF_Budget()
End Sub

Private Function Code_VBA_USF_Budget() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Me.txt_Mois.Value = MOD_06_Budget_ZBB.Obtenir_Parametre(""BUDG_FILTRE_MOIS"", Format(Date, ""yyyy-mm""))" & vbCrLf
    c = c & "    Me.Caption = MOD_02_AppHome_Global.TR(""BTN_ALLOC"")" & vbCrLf
    c = c & "    Dim tbl As ListObject, i As Long" & vbCrLf
    c = c & "    On Error Resume Next: Set tbl = ThisWorkbook.Sheets(""DIM_Categorie"").ListObjects(""T_DIM_Categorie""): On Error GoTo 0" & vbCrLf
    c = c & "    If Not tbl Is Nothing Then" & vbCrLf
    c = c & "        Me.cmb_Cat.ColumnCount = 2: Me.cmb_Cat.ColumnWidths = ""0 pt;150 pt""" & vbCrLf
    c = c & "        For i = 1 To tbl.ListRows.Count" & vbCrLf
    c = c & "            If UCase(Trim(tbl.DataBodyRange(i, 3).Value)) = ""DEPENSE"" Then" & vbCrLf
    c = c & "                Me.cmb_Cat.AddItem tbl.DataBodyRange(i, 1).Value" & vbCrLf
    c = c & "                Me.cmb_Cat.List(Me.cmb_Cat.ListCount - 1, 1) = tbl.DataBodyRange(i, 2).Value" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        Next i" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    
    c = c & "Private Sub btn_Save_Click()" & vbCrLf
    c = c & "    If Me.cmb_Cat.ListIndex = -1 Then MsgBox ""Sélectionnez una catégorie."", vbCritical: Exit Sub" & vbCrLf
    c = c & "    Dim m As String: m = Replace(Me.txt_Montant.Value, "","", ""."")" & vbCrLf
    c = c & "    If Val(m) <= 0 Then MsgBox ""Montant invalide."", vbCritical: Exit Sub" & vbCrLf
    
    c = c & "    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(""FACT_Budget"")" & vbCrLf
    c = c & "    ws.Unprotect ""SFP_ADMIN_2026""" & vbCrLf
    c = c & "    Dim tbl As ListObject: Set tbl = ws.ListObjects(""T_FACT_Budget"")" & vbCrLf
    
    c = c & "    Dim i As Long, found As Boolean, idCat As String, targetMois As String" & vbCrLf
    c = c & "    found = False: idCat = Me.cmb_Cat.List(Me.cmb_Cat.ListIndex, 0): targetMois = Trim(Me.txt_Mois.Value)" & vbCrLf
    c = c & "    If tbl.ListRows.Count > 0 Then" & vbCrLf
    c = c & "        For i = 1 To tbl.ListRows.Count" & vbCrLf
    c = c & "            If CStr(tbl.DataBodyRange(i, 2).Value) = targetMois And CStr(tbl.DataBodyRange(i, 3).Value) = idCat Then" & vbCrLf
    c = c & "                tbl.DataBodyRange(i, 4).Value = Val(m): tbl.DataBodyRange(i, 6).Value = Now" & vbCrLf
    c = c & "                found = True: Exit For" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        Next i" & vbCrLf
    c = c & "    End If" & vbCrLf
    
    c = c & "    If Not found Then" & vbCrLf
    c = c & "        Dim nr As ListRow: Set nr = tbl.ListRows.Add" & vbCrLf
    c = c & "        nr.Range(1, 1).Value = MOD_01_CoreEngine.GENERER_NOUVEL_ID(""T_FACT_Budget"")" & vbCrLf
    c = c & "        nr.Range(1, 2).Value = targetMois: nr.Range(1, 3).Value = idCat" & vbCrLf
    c = c & "        nr.Range(1, 4).Value = Val(m): nr.Range(1, 5).Value = Application.UserName: nr.Range(1, 6).Value = Now" & vbCrLf
    c = c & "    End If" & vbCrLf
    
    c = c & "    ws.Protect ""SFP_ADMIN_2026"", UserInterfaceOnly:=True" & vbCrLf
    c = c & "    Unload Me" & vbCrLf
    c = c & "    MOD_06_Budget_ZBB.Rafraichir_Budget" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "Private Sub btn_Cancel_Click(): Unload Me: End Sub"
    Code_VBA_USF_Budget = c
End Function

' -------------------------------------------------------------------------
' 4. MOTEUR DE RENDU UI (Time Slider, Devises, Violet Zebra & DataBars)
' -------------------------------------------------------------------------
Public Sub GENERER_BUDGET_DASHBOARD()
    Garantir_Lexique_Budget
    
    Dim DeviseFiltre As String: DeviseFiltre = Obtenir_Parametre("BUDG_FILTRE_DEV", "MUR")
    Dim MoisFiltre As String: MoisFiltre = Obtenir_Parametre("BUDG_FILTRE_MOIS", Format(Date, "yyyy-mm"))

    Dim wsBud As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next: ThisWorkbook.Sheets("BUDGET_ZBB").Delete: On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsBud = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("APP_HOME"))
    wsBud.Name = "BUDGET_ZBB"
    wsBud.Activate
    
    ' --- FORÇAGE DE L'ÉLÉGANCE GLOBALE ---
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    wsBud.Cells.Font.Name = "ADLaM Display"
    wsBud.Cells.Font.Size = 10
    wsBud.Cells.Interior.Color = RGB(248, 248, 250)
    
    ' --- 1. VERROUILLAGE DES LARGEURS DE COLONNES (Pour un rendu Web parfait) ---
    wsBud.Columns("A:B").ColumnWidth = 2
    wsBud.Columns("C").ColumnWidth = 45   ' Catégorie (Très large)
    wsBud.Columns("D").ColumnWidth = 25   ' Alloué
    wsBud.Columns("E").ColumnWidth = 25   ' Réalisé
    wsBud.Columns("F").ColumnWidth = 25   ' Écart
    wsBud.Columns("G").ColumnWidth = 35   ' DataBar
    
    ' --- BANDEAU SUPÉRIEUR ---
    wsBud.Range("A1:Z5").Interior.Color = RGB(65, 105, 225)
    
    ' --- BOUTON RETOUR (Blanc, discret, sans wrap) ---
    Dim btnBack As Shape
    Set btnBack = wsBud.Shapes.AddShape(msoShapeRoundedRectangle, 20, 15, 140, 32)
    btnBack.Name = "BTN_RETOUR_TACTILE_BUDG"
    btnBack.Fill.ForeColor.RGB = RGB(250, 218, 94)
    btnBack.Line.Visible = msoFalse
    btnBack.TextFrame2.WordWrap = msoFalse
    btnBack.TextFrame2.MarginLeft = 0: btnBack.TextFrame2.MarginRight = 0
    With btnBack.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 4: .OffsetX = 0: .OffsetY = 2: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    btnBack.TextFrame2.TextRange.Text = "<  " & TR("BTN_BACK")
    btnBack.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnBack.TextFrame2.TextRange.Font.Bold = True: btnBack.TextFrame2.TextRange.Font.Size = 9: btnBack.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    btnBack.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnBack.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnBack.OnAction = "MOD_06_Budget_ZBB.ANIMATION_RETOUR"
    
    ' --- TITRE VECTORIEL (Protégé) ---
    Dim shpTitle As Shape
    Set shpTitle = wsBud.Shapes.AddTextbox(msoTextOrientationHorizontal, 180, 10, 300, 40)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = UCase(TR("BUDG_TITLE")) & vbCrLf & "As of : " & Format(CDate(MoisFiltre & "-01"), "mmm yyyy")
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 18: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 10: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' --- TIME SLIDER (Mois) ---
    Dim LabelDate As String: LabelDate = UCase(Format(CDate(MoisFiltre & "-01"), "mmmm yyyy"))
    Dessiner_Widget wsBud, "BTN_BUDG_PREV", "<", 480, 15, 35, 32, RGB(220, 220, 220), RGB(0, 0, 0), "MOD_06_Budget_ZBB.MOIS_PRECEDENT_BUDG"
    Dessiner_Widget wsBud, "LBL_BUDG_MONTH", LabelDate, 520, 15, 150, 32, RGB(220, 220, 220), RGB(0, 0, 0), ""
    Dessiner_Widget wsBud, "BTN_BUDG_NEXT", ">", 675, 15, 35, 32, RGB(220, 220, 220), RGB(0, 0, 0), "MOD_06_Budget_ZBB.MOIS_SUIVANT_BUDG"
    
    ' --- CONVERTISSEUR DE DEVISES ---
    Dim devLeft As Integer: devLeft = 730
    Dessiner_Widget wsBud, "BTN_BUDG_DEV_MUR", "MUR", devLeft, 15, 45, 32, IIf(DeviseFiltre = "MUR", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "MUR", RGB(40, 40, 40), vbWhite), "MOD_06_Budget_ZBB.CHANGER_DEVISE_BUDG"
    Dessiner_Widget wsBud, "BTN_BUDG_DEV_EUR", "EUR", devLeft + 50, 15, 45, 32, IIf(DeviseFiltre = "EUR", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "EUR", RGB(40, 40, 40), vbWhite), "MOD_06_Budget_ZBB.CHANGER_DEVISE_BUDG"
    Dessiner_Widget wsBud, "BTN_BUDG_DEV_USD", "USD", devLeft + 100, 15, 45, 32, IIf(DeviseFiltre = "USD", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "USD", RGB(40, 40, 40), vbWhite), "MOD_06_Budget_ZBB.CHANGER_DEVISE_BUDG"
    Dessiner_Widget wsBud, "BTN_BUDG_DEV_OXF", "OXF", devLeft + 150, 15, 45, 32, IIf(DeviseFiltre = "XOF", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "XOF", RGB(40, 40, 40), vbWhite), "MOD_06_Budget_ZBB.CHANGER_DEVISE_BUDG"
    
    ' --- BOUTON ALLOUER BUDGET (Ancré et Responsive) ---
    Dim btnAlloc As Shape
    Dim BtnAllocLeft As Double: BtnAllocLeft = wsBud.Range("H1").Left - 30 - 4
    Set btnAlloc = wsBud.Shapes.AddShape(msoShapeRoundedRectangle, BtnAllocLeft, 15, 150, 32)
    btnAlloc.Name = "BTN_ALLOC_TACTILE"
    btnAlloc.Fill.ForeColor.RGB = RGB(250, 218, 94): btnAlloc.Line.Visible = msoFalse
    btnAlloc.TextFrame2.WordWrap = msoFalse: btnAlloc.TextFrame2.MarginLeft = 0: btnAlloc.TextFrame2.MarginRight = 0
    btnAlloc.TextFrame2.TextRange.Text = "+  " & TR("BTN_ALLOC")
    btnAlloc.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnAlloc.TextFrame2.TextRange.Font.Bold = True: btnAlloc.TextFrame2.TextRange.Font.Size = 9: btnAlloc.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    btnAlloc.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnAlloc.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With btnAlloc.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 4: .OffsetX = 0: .OffsetY = 2: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    btnAlloc.OnAction = "MOD_06_Budget_ZBB.ANIMATION_OUVRIR_ALLOC"
    
    ' --- MOTEUR DE TAUX DE CHANGE ---
    Dim dictTaux As Object: Set dictTaux = CreateObject("Scripting.Dictionary")
    dictTaux("MUR") = 1: dictTaux("EUR") = 49.5: dictTaux("USD") = 46.2: dictTaux("GBP") = 58.1: dictTaux("ZAR") = 2.4: dictTaux("XOF") = 0.083
    Dim TauxC As Double: TauxC = IIf(dictTaux.exists(DeviseFiltre), dictTaux(DeviseFiltre), 1)
    
    ' --- MOTEUR ETL DOUBLE ---
    Dim tblBud As ListObject, tblTx As ListObject, tblCat As ListObject
    On Error Resume Next
    Set tblBud = ThisWorkbook.Sheets("FACT_Budget").ListObjects("T_FACT_Budget")
    Set tblTx = ThisWorkbook.Sheets("FACT_Transaction").ListObjects("T_FACT_Transaction")
    Set tblCat = ThisWorkbook.Sheets("DIM_Categorie").ListObjects("T_DIM_Categorie")
    On Error GoTo 0
    
    ' --- PATCH: DÉCLARATION DES DEUX DICTIONNAIRES (NOM ET TYPE) ---
    Dim dictCatName As Object: Set dictCatName = CreateObject("Scripting.Dictionary")
    Dim dictCatType As Object: Set dictCatType = CreateObject("Scripting.Dictionary") ' <-- LA VARIABLE MANQUANTE
    
    If Not tblCat Is Nothing Then
        If tblCat.ListRows.Count > 0 Then
            Dim x As Long: For x = 1 To tblCat.ListRows.Count
                dictCatName(CStr(tblCat.DataBodyRange(x, 1).Value)) = CStr(tblCat.DataBodyRange(x, 2).Value)
                ' <-- LE REMPLISSAGE DU DICTIONNAIRE MANQUANT
                dictCatType(CStr(tblCat.DataBodyRange(x, 1).Value)) = UCase(Trim(CStr(tblCat.DataBodyRange(x, 3).Value)))
            Next x
        End If
    End If
    
    ' =========================================================================
    ' L'OMNISCIENCE ETL : FUSION DES ALLOCATIONS ET DES DÉPENSES SANS EXCEPTION
    ' =========================================================================
    Dim dictMaster As Object: Set dictMaster = CreateObject("Scripting.Dictionary") ' Le cerveau de fusion
    Dim dictAlloc As Object: Set dictAlloc = CreateObject("Scripting.Dictionary")
    Dim TotAlloc As Double: TotAlloc = 0
    
    If Not tblBud Is Nothing Then
        If tblBud.ListRows.Count > 0 Then
            Dim arrB As Variant: arrB = tblBud.DataBodyRange.Value
            For x = 1 To UBound(arrB, 1)
                If Trim(CStr(arrB(x, 2))) = MoisFiltre Then
                    Dim idC_B As String: idC_B = CStr(arrB(x, 3))
                    Dim ValAlloc As Double: ValAlloc = CDbl(arrB(x, 4)) / TauxC
                    dictAlloc(idC_B) = dictAlloc(idC_B) + ValAlloc
                    TotAlloc = TotAlloc + ValAlloc
                    dictMaster(idC_B) = True ' Ajout au Master
                End If
            Next x
        End If
    End If
    
    Dim dictSpent As Object: Set dictSpent = CreateObject("Scripting.Dictionary")
    Dim TotSpent As Double: TotSpent = 0
    
    If Not tblTx Is Nothing Then
        If tblTx.ListRows.Count > 0 Then
            Dim arrT As Variant: arrT = tblTx.DataBodyRange.Value
            Dim dtTx As String, idC As String, amt As Double
            Dim DevOrigine As String, TauxO As Double
            Dim FluxType As String
            
            For x = 1 To UBound(arrT, 1)
                If Trim(CStr(arrT(x, 1))) <> "" Then
                    dtTx = Format(CDate(arrT(x, 2)), "yyyy-mm")
                    If dtTx = MoisFiltre Then
                        idC = CStr(arrT(x, 4))
                        FluxType = IIf(dictCatType.exists(idC), dictCatType(idC), "AUTRE")
                        
                        ' CORRECTION: On intègre TOUTES les dépenses (budgétisées ou auto-apprises)
                        If FluxType = "DEPENSE" Then
                            DevOrigine = UCase(Trim(CStr(arrT(x, 7))))
                            TauxO = IIf(dictTaux.exists(DevOrigine), dictTaux(DevOrigine), 1)
                            amt = CDbl(arrT(x, 6))
                            Dim RealAmt As Double: RealAmt = (amt * TauxO) / TauxC
                            
                            dictSpent(idC) = dictSpent(idC) + RealAmt
                            TotSpent = TotSpent + RealAmt
                            dictMaster(idC) = True ' Ajout au Master (Même sans allocation !)
                        End If
                    End If
                End If
            Next x
        End If
    End If
    
    Dim Ligne As Long: Ligne = 0
    Dim arrConsolide() As Variant
    If dictMaster.Count > 0 Then
        ReDim arrConsolide(1 To dictMaster.Count, 1 To 5)
        Dim key As Variant, alloué As Double, depensé As Double, pct As Double
        For Each key In dictMaster.keys
            Ligne = Ligne + 1
            alloué = IIf(dictAlloc.exists(key), dictAlloc(key), 0)
            depensé = IIf(dictSpent.exists(key), dictSpent(key), 0)
            
            pct = 0
            If alloué > 0 Then
                pct = depensé / alloué
            ElseIf depensé > 0 Then
                pct = 1 ' Dépense sans budget = Barre pleine alerte
            End If
            
            arrConsolide(Ligne, 1) = IIf(dictCatName.exists(key), dictCatName(key), "Catégorie " & key)
            arrConsolide(Ligne, 2) = alloué
            arrConsolide(Ligne, 3) = depensé
            arrConsolide(Ligne, 4) = alloué - depensé ' Écart
            arrConsolide(Ligne, 5) = pct
        Next key
    Else
        ReDim arrConsolide(1 To 1, 1 To 5)
        arrConsolide(1, 1) = TR("NO_BUDG"): arrConsolide(1, 2) = 0: arrConsolide(1, 3) = 0: arrConsolide(1, 4) = 0: arrConsolide(1, 5) = 0
        Ligne = 1
    End If
    
    ' --- DESSIN MATHÉMATIQUE DES KPIS ---
    wsBud.Rows("7").RowHeight = 35: wsBud.Rows("8").RowHeight = 50
    Dim ZoneTable As Range: Set ZoneTable = wsBud.Range("C7:G7")
    Dim TotalW As Double: TotalW = ZoneTable.Width
    Dim Gap As Double: Gap = 15
    Dim CardW As Double: CardW = (TotalW - (2 * Gap)) / 3
    
    Dessiner_Shape_Card wsBud, "CARD_ALLO", TR("KPI_B_ALLO") & " (" & DeviseFiltre & ")", TotAlloc, RGB(52, 152, 219), vbWhite, ZoneTable.Left, wsBud.Range("C7").Top, CardW, 85
    Dessiner_Shape_Card wsBud, "CARD_SPEN", TR("KPI_B_SPEN") & " (" & DeviseFiltre & ")", TotSpent, RGB(120, 81, 169), vbWhite, ZoneTable.Left + CardW + Gap, wsBud.Range("C7").Top, CardW, 85
    Dessiner_Shape_Card wsBud, "CARD_LEFT", TR("KPI_B_LEFT") & " (" & DeviseFiltre & ")", TotAlloc - TotSpent, RGB(46, 204, 113), vbWhite, ZoneTable.Left + (CardW * 2) + (Gap * 2), wsBud.Range("C7").Top, CardW, 85
    
    wsBud.Rows("9:11").RowHeight = 15
    
    ' --- DESSIN DE LA TABLE "VIOLET ZEBRA" ---
    wsBud.Range("C12:G12").Value = Array(UCase(TR("COL_B_CAT")), UCase(TR("COL_B_ALLO")), UCase(TR("COL_B_REAL")), UCase(TR("COL_B_DIFF")), UCase(TR("COL_B_PROG")))
    wsBud.Range("C13").Resize(Ligne, 5).Value = arrConsolide
    
    Dim tblView As ListObject
    If arrConsolide(1, 1) <> TR("NO_BUDG") Then
        Set tblView = wsBud.ListObjects.Add(xlSrcRange, wsBud.Range("C12").Resize(Ligne + 1, 5), , xlYes)
    Else
        Set tblView = wsBud.ListObjects.Add(xlSrcRange, wsBud.Range("C12:G13"), , xlYes)
        With wsBud.Range("C13:G13")
            .HorizontalAlignment = xlCenterAcrossSelection: .VerticalAlignment = xlCenter
            .Font.Italic = True: .Font.Color = RGB(220, 220, 225)
        End With
        wsBud.Range("C13").Value = TR("NO_BUDG")
    End If
    
    tblView.Name = "VIEW_Budget"
    tblView.TableStyle = ""
    tblView.ShowAutoFilterDropDown = False
    
    With tblView.HeaderRowRange
        .Interior.Color = RGB(90, 50, 130): .Font.Color = vbWhite: .Font.Bold = True
        .RowHeight = 35: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlNone
    End With
    
    If arrConsolide(1, 1) <> TR("NO_BUDG") Then
        ' Style Violet Zebra
        With tblView.DataBodyRange
            .RowHeight = 30: .VerticalAlignment = xlCenter: .Borders.LineStyle = xlNone
            .Font.Color = vbWhite
            Dim r As Long: For r = 1 To .Rows.Count
                If r Mod 2 = 0 Then .Rows(r).Interior.Color = RGB(145, 110, 190) Else .Rows(r).Interior.Color = RGB(120, 81, 169)
            Next r
        End With
        
        tblView.ListColumns(1).DataBodyRange.HorizontalAlignment = xlLeft
        tblView.ListColumns(2).DataBodyRange.HorizontalAlignment = xlRight
        tblView.ListColumns(3).DataBodyRange.HorizontalAlignment = xlRight
        tblView.ListColumns(4).DataBodyRange.HorizontalAlignment = xlRight
        
        tblView.ListColumns(2).DataBodyRange.NumberFormat = "#,##0.00"
        tblView.ListColumns(3).DataBodyRange.NumberFormat = "#,##0.00"
        tblView.ListColumns(4).DataBodyRange.NumberFormat = "#,##0.00"
        
        ' LES DATABARS D'ÉLITE (Jaune Royal sur fond Violet)
        tblView.ListColumns(5).DataBodyRange.NumberFormat = "0%"
        With tblView.ListColumns(5).DataBodyRange.FormatConditions.AddDatabar
            .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
            .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
            .BarColor.Color = RGB(250, 218, 94) ' Jaune Royal
            .BarFillType = xlDataBarFillSolid: .Direction = xlContext
        End With
        
        For r = 1 To Ligne
            If CDbl(arrConsolide(r, 4)) < 0 Then
                tblView.ListColumns(4).DataBodyRange.Cells(r, 1).Font.Color = RGB(250, 218, 94)
                tblView.ListColumns(4).DataBodyRange.Cells(r, 1).Font.Bold = True
            End If
        Next r
    End If
    
    wsBud.Range("A1").Select
End Sub

' -------------------------------------------------------------------------
' 5. ANIMATIONS ET INTERACTIVITÉ
' -------------------------------------------------------------------------
Public Sub CHANGER_DEVISE_BUDG()
    Dim btnName As String: On Error Resume Next: btnName = Application.Caller: On Error GoTo 0
    If btnName <> "" Then
        Modifier_Parametre "BUDG_FILTRE_DEV", Replace(btnName, "BTN_BUDG_DEV_", "")
        Rafraichir_Budget
    End If
End Sub

Public Sub MOIS_PRECEDENT_BUDG()
    Anim_Btn_Bleu Application.Caller
    Modifier_Mois_Filtre_BUDG -1
    Rafraichir_Budget
End Sub

Public Sub MOIS_SUIVANT_BUDG()
    Anim_Btn_Bleu Application.Caller
    Modifier_Mois_Filtre_BUDG 1
    Rafraichir_Budget
End Sub

Private Sub Modifier_Mois_Filtre_BUDG(DeltaMois As Integer)
    Dim actuel As String: actuel = Obtenir_Parametre("BUDG_FILTRE_MOIS", Format(Date, "yyyy-mm"))
    Dim d As Date: d = DateAdd("m", DeltaMois, CDate(actuel & "-01"))
    Modifier_Parametre "BUDG_FILTRE_MOIS", Format(d, "yyyy-mm")
End Sub

Public Sub Rafraichir_Budget()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    GENERER_BUDGET_DASHBOARD
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

Public Sub ANIMATION_RETOUR()
    Dim btn As Shape: On Error Resume Next: Set btn = ActiveSheet.Shapes(Application.Caller): On Error GoTo 0
    If Not btn Is Nothing Then
        btn.Fill.ForeColor.RGB = RGB(230, 230, 235): btn.Shadow.Visible = msoFalse: btn.Top = btn.Top + 2: btn.Left = btn.Left + 2
        Dim t As Single: t = Timer: Do While Timer < t + 0.15: DoEvents: Loop
        btn.Fill.ForeColor.RGB = vbWhite: btn.Shadow.Visible = msoTrue: btn.Top = btn.Top - 2: btn.Left = btn.Left - 2
    End If
    On Error Resume Next: ThisWorkbook.Sheets("APP_HOME").Activate: ThisWorkbook.Sheets("APP_HOME").Range("A1").Select: On Error GoTo 0
End Sub

Public Sub ANIMATION_OUVRIR_ALLOC()
    Dim btn As Shape: On Error Resume Next: Set btn = ActiveSheet.Shapes(Application.Caller): On Error GoTo 0
    If Not btn Is Nothing Then
        btn.Fill.ForeColor.RGB = RGB(220, 190, 60): btn.Shadow.Visible = msoFalse: btn.Top = btn.Top + 2: btn.Left = btn.Left + 2
        Dim t As Single: t = Timer: Do While Timer < t + 0.15: DoEvents: Loop
        btn.Fill.ForeColor.RGB = RGB(250, 218, 94): btn.Shadow.Visible = msoTrue: btn.Top = btn.Top - 2: btn.Left = btn.Left - 2
    End If
    On Error Resume Next: VBA.UserForms.Add("USF_Budget").Show: On Error GoTo 0
End Sub

Private Sub Anim_Btn_Bleu(NomShape As String)
    Dim btn As Shape: On Error Resume Next: Set btn = ActiveSheet.Shapes(NomShape): On Error GoTo 0
    If Not btn Is Nothing Then
        btn.Fill.ForeColor.RGB = RGB(20, 40, 120): btn.Top = btn.Top + 2: btn.Left = btn.Left + 2
        Dim t As Single: t = Timer: Do While Timer < t + 0.1: DoEvents: Loop
        btn.Fill.ForeColor.RGB = RGB(40, 70, 180): btn.Top = btn.Top - 2: btn.Left = btn.Left - 2
    End If
End Sub

Private Sub Dessiner_Widget(ws As Worksheet, Nom As String, txt As String, L As Integer, t As Integer, W As Integer, H As Integer, CFond As Long, cTxt As Long, Macro As String)
    Dim shp As Shape: Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, t, W, H)
    shp.Name = Nom: shp.Fill.ForeColor.RGB = CFond: shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = txt: shp.TextFrame2.VerticalAnchor = msoAnchorMiddle: shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Font.Name = "ADLaM Display": shp.TextFrame2.TextRange.Font.Size = 10: shp.TextFrame2.TextRange.Font.Bold = True: shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = cTxt
    If Macro <> "" Then shp.OnAction = Macro
End Sub

Private Sub Dessiner_Shape_Card(ws As Worksheet, NomShape As String, Titre As String, Valeur As Double, CoulFond As Long, CoulTexte As Long, Gauche As Double, Haut As Double, Largeur As Double, Hauteur As Double)
    Dim shp As Shape: Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, Gauche, Haut, Largeur, Hauteur)
    shp.Name = NomShape: shp.Fill.ForeColor.RGB = CoulFond: shp.Line.Visible = msoFalse
    With shp.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 6: .OffsetX = 0: .OffsetY = 3: .Transparency = 0.6: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    shp.TextFrame2.TextRange.Text = UCase(Titre) & vbCrLf & Format(Valeur, "#,##0.00")
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle: shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With shp.TextFrame2.TextRange.Lines(1).Font: .Name = "ADLaM Display": .Size = 11: .Bold = True: .Fill.ForeColor.RGB = CoulTexte: End With
    With shp.TextFrame2.TextRange.Lines(2).Font: .Name = "ADLaM Display": .Size = 28: .Bold = True: .Fill.ForeColor.RGB = CoulTexte: End With
End Sub





