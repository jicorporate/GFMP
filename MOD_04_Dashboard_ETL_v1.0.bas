Option Explicit

' =========================================================================
' MODULE: MOD_04_Dashboard_ETL
' OBJECTIF: ETL Temporel, Convertisseur de Devises, Violet Zebra, ADLaM 10
' =========================================================================

Public Sub DEPLOIEMENT_ETAPE_5_DASHBOARD()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    
    Garantir_Lexique_Dashboard
    GENERER_DASHBOARD
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

' -------------------------------------------------------------------------
' 1. STATE MANAGEMENT (Mois et Devise Actifs)
' -------------------------------------------------------------------------
Private Function Obtenir_Parametre(NomParam As String, ValeurDefaut As String) As String
    Dim tblConf As ListObject, i As Long
    On Error Resume Next: Set tblConf = ThisWorkbook.Sheets("SYS_Config").ListObjects("T_SYS_Config"): On Error GoTo 0
    If tblConf Is Nothing Then Obtenir_Parametre = ValeurDefaut: Exit Function
    
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = NomParam Then
            Obtenir_Parametre = tblConf.DataBodyRange(i, 2).Value
            Exit Function
        End If
    Next i
    
    Dim nr As ListRow: Set nr = tblConf.ListRows.Add
    nr.Range(1, 1).Value = NomParam: nr.Range(1, 2).Value = ValeurDefaut: nr.Range(1, 3).Value = "Filtre Actif"
    Obtenir_Parametre = ValeurDefaut
End Function

Private Sub Modifier_Parametre(NomParam As String, NouvelleValeur As String)
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config"): wsSys.Unprotect "SFP_ADMIN_2026"
    Dim tblConf As ListObject: Set tblConf = wsSys.ListObjects("T_SYS_Config")
    Dim i As Long
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = NomParam Then
            tblConf.DataBodyRange(i, 2).Value = NouvelleValeur: Exit For
        End If
    Next i
    wsSys.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
End Sub

' -------------------------------------------------------------------------
' 2. MOTEUR AUTO-CICATRISANT
' -------------------------------------------------------------------------
Public Sub Garantir_Lexique_Dashboard()
    Dim tblDic As ListObject
    On Error Resume Next: Set tblDic = ThisWorkbook.Sheets("SYS_Config").ListObjects("T_SYS_Dictionary"): On Error GoTo 0
    If tblDic Is Nothing Then Exit Sub
    
    Upsert_Dico tblDic, "KPI_INC", "TOTAL REVENUS", "TOTAL INCOME", "TOTAL INGRESOS", "RENDA TOTAL", "GESAMTEINKOMMEN", "TOTALE ENTRATE", "TOTALE INKOMSTEN", "TOTAL INKOMST"
    Upsert_Dico tblDic, "KPI_EXP", "TOTAL DÉPENSES", "TOTAL EXPENSES", "TOTAL GASTOS", "DESPESAS TOTAIS", "GESAMTAUSGABEN", "TOTALE USCITE", "TOTALE UITGAVEN", "TOTALA UTGIFTER"
    Upsert_Dico tblDic, "KPI_NET", "CASHFLOW NET", "NET CASHFLOW", "FLUJO NETO", "FLUXO LÍQUIDO", "NETTO-CASHFLOW", "CASHFLOW NETTO", "NETTO CASHFLOW", "NETTO KASSAFLÖDE"
    Upsert_Dico tblDic, "BTN_BACK", "RETOUR AU MENU", "BACK TO MENU", "VOLVER AL MENÚ", "VOLTAR AO MENU", "ZURÜCK ZUM MENÜ", "TORNA AL MENU", "TERUG NAAR MENU", "TILLBAKA TILL MENY"
    Upsert_Dico tblDic, "COL_TYPE", "TYPE DE FLUX", "FLOW TYPE", "TIPO DE FLUJO", "TIPO DE FLUXO", "FLUSSTYP", "TIPO DI FLUSSO", "STROOMTYPE", "FLÖDESTYP"
    Upsert_Dico tblDic, "NO_DATA", "Aucune transaction ce mois-ci", "No transactions this month", "Sin transacciones este mes", "Nenhuma transação este mês", "Keine Transaktionen in diesem Monat", "Nessuna transazione questo mese", "Geen transacties deze maand", "Inga transaktioner denna månad"
    Upsert_Dico tblDic, "COL_YEAR", "Année", "Year", "Año", "Ano", "Jahr", "Anno", "Jaar", "År"
    Upsert_Dico tblDic, "COL_MONTH", "Mois", "Month", "Mes", "Mês", "Monat", "Mese", "Maand", "Månad"
End Sub

Private Sub Upsert_Dico(tbl As ListObject, k As String, fr As String, en As String, es As String, pt As String, de As String, it As String, nl As String, sv As String)
    Dim i As Long: For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 1).Value = k Then Exit Sub
    Next i
    Dim nr As ListRow: Set nr = tbl.ListRows.Add
    nr.Range(1, 1).Value = k: nr.Range(1, 2).Value = fr: nr.Range(1, 3).Value = en: nr.Range(1, 4).Value = es
    nr.Range(1, 5).Value = pt: nr.Range(1, 6).Value = de: nr.Range(1, 7).Value = it: nr.Range(1, 8).Value = nl: nr.Range(1, 9).Value = sv
End Sub

Private Function TR(Clé As String) As String
    TR = MOD_02_AppHome_Global.TR(Clé)
End Function

' -------------------------------------------------------------------------
' 3. LE MOTEUR UI (ETL Filtré, Devises, Violet Zebra)
' -------------------------------------------------------------------------
Public Sub GENERER_DASHBOARD()
    Garantir_Lexique_Dashboard
    
    Dim MoisFiltre As String: MoisFiltre = Obtenir_Parametre("DASH_FILTRE_MOIS", Format(Date, "yyyy-mm"))
    Dim DeviseFiltre As String: DeviseFiltre = Obtenir_Parametre("DASH_FILTRE_DEV", "MUR")

    Dim wsDash As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next: ThisWorkbook.Sheets("DASHBOARD").Delete: On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("APP_HOME"))
    wsDash.Name = "DASHBOARD"
    wsDash.Activate
    
    ' --- FORÇAGE DE LA POLICE GLOBALE ET DU ZOOM 100% ---
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    wsDash.Cells.Font.Name = "ADLaM Display"
    wsDash.Cells.Font.Size = 10
    wsDash.Cells.Interior.Color = RGB(248, 248, 250)
    
    ' --- BANDEAU SUPÉRIEUR ---
    wsDash.Range("A1:Z5").Interior.Color = RGB(65, 105, 225) ' Bleu Royal (légèrement plus haut)
    
    ' --- BOUTON RETOUR TACTILE (Charte Jaune Royal, 1 Ligne, Anti-Wrap) ---
    Dim btnBack As Shape
    Set btnBack = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, 20, 15, 200, 32) ' Élargi pour tenir sur 1 ligne
    btnBack.Name = "BTN_RETOUR_TACTILE"
    btnBack.Fill.ForeColor.RGB = RGB(250, 218, 94) ' Jaune Royal
    btnBack.Line.Visible = msoFalse
    btnBack.TextFrame2.WordWrap = msoFalse ' EMPÊCHE LE TEXTE DE S'ÉCRASER
    btnBack.TextFrame2.MarginLeft = 0: btnBack.TextFrame2.MarginRight = 0
    With btnBack.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow
        .Blur = 4: .OffsetX = 1: .OffsetY = 2: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    btnBack.TextFrame2.TextRange.Text = "<  " & TR("BTN_BACK") ' Flèche ANSI incassable
    btnBack.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnBack.TextFrame2.TextRange.Font.Bold = True: btnBack.TextFrame2.TextRange.Font.Size = 10: btnBack.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    btnBack.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnBack.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnBack.OnAction = "MOD_04_Dashboard_ETL.ANIMATION_RETOUR"
    
    ' --- TITRE VECTORIEL (Séparé des cellules pour éviter tout chevauchement) ---
    Dim shpTitle As Shape
    Set shpTitle = wsDash.Shapes.AddTextbox(msoTextOrientationHorizontal, 240, 10, 300, 40)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = UCase(TR("DASH_T")) & vbCrLf & Format(Date, "dd mmm yyyy")
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 18: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 10: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' --- LE TIME SLIDER (Flèches ANSI) ---
    Dim LabelDate As String: LabelDate = UCase(Format(CDate(MoisFiltre & "-01"), "mmmm yyyy"))
    Dessiner_Widget wsDash, "BTN_PREV_MONTH", "<", 440, 15, 20, 32, RGB(220, 220, 220), RGB(0, 0, 0), "MOD_04_Dashboard_ETL.MOIS_PRECEDENT"
    Dessiner_Widget wsDash, "LBL_MONTH", LabelDate, 465, 15, 145, 32, RGB(220, 220, 220), RGB(0, 0, 0), ""
    Dessiner_Widget wsDash, "BTN_NEXT_MONTH", ">", 615, 15, 25, 32, RGB(220, 220, 220), RGB(0, 0, 0), "MOD_04_Dashboard_ETL.MOIS_SUIVANT"
    
    ' --- LE CONVERTISSEUR DE DEVISES (Les 3 boutons tactiles) ---
    Dim devLeft As Integer: devLeft = 710
    Dessiner_Widget wsDash, "BTN_DEV_MUR", "MUR", devLeft, 15, 50, 35, IIf(DeviseFiltre = "MUR", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "MUR", RGB(40, 40, 40), vbWhite), "MOD_04_Dashboard_ETL.CHANGER_DEVISE"
    Dessiner_Widget wsDash, "BTN_DEV_EUR", "EUR", devLeft + 55, 15, 50, 35, IIf(DeviseFiltre = "EUR", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "EUR", RGB(40, 40, 40), vbWhite), "MOD_04_Dashboard_ETL.CHANGER_DEVISE"
    Dessiner_Widget wsDash, "BTN_DEV_USD", "USD", devLeft + 110, 15, 50, 35, IIf(DeviseFiltre = "USD", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "USD", RGB(40, 40, 40), vbWhite), "MOD_04_Dashboard_ETL.CHANGER_DEVISE"
    Dessiner_Widget wsDash, "BTN_DEV_XOF", "XOF", devLeft + 165, 15, 50, 35, IIf(DeviseFiltre = "XOF", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "XOF", RGB(40, 40, 40), vbWhite), "MOD_04_Dashboard_ETL.CHANGER_DEVISE"
    
    ' --- MOTEUR DE TAUX DE CHANGE (Taux basés sur 1 MUR) ---
    Dim dictTaux As Object: Set dictTaux = CreateObject("Scripting.Dictionary")
    dictTaux("MUR") = 1
    dictTaux("EUR") = 49.5  ' 1 EUR = 49.5 MUR
    dictTaux("USD") = 46.2  ' 1 USD = 46.2 MUR
    dictTaux("GBP") = 58.1
    dictTaux("ZAR") = 2.4
    dictTaux("XOF") = 0.083
    
    ' --- PHASE D'EXTRACTION (ETL FILTRÉ EN MÉMOIRE) ---
    Dim tblFact As ListObject
    On Error Resume Next: Set tblFact = ThisWorkbook.Sheets("FACT_Transaction").ListObjects("T_FACT_Transaction"): On Error GoTo 0
    
    Dim TotRev As Double: TotRev = 0
    Dim TotDep As Double: TotDep = 0
    Dim arrConsolide() As Variant
    Dim Ligne As Long: Ligne = 0
    
    If Not tblFact Is Nothing Then
        If tblFact.ListRows.Count > 0 Then
            Dim dictCompte As Object: Set dictCompte = Charger_Dico("T_DIM_Compte", 1, 2)
            Dim dictCat As Object: Set dictCat = Charger_Dico("T_DIM_Categorie", 1, 2)
            Dim dictCatType As Object: Set dictCatType = Charger_Dico("T_DIM_Categorie", 1, 3)
            Dim dictTiers As Object: Set dictTiers = Charger_Dico("T_DIM_Tiers", 1, 2)
            
            Dim arrFact As Variant: arrFact = tblFact.DataBodyRange.Value
            ReDim arrConsolide(1 To UBound(arrFact, 1), 1 To 11)
            
            Dim i As Long, IDCompte As String, idCat As String, IDTiers As String
            Dim TypeFlux As String, Montant As Double, dTx As Date
            Dim DevOrigine As String, TauxO As Double, TauxC As Double, MontantConverti As Double
            
            ' Taux de la devise cible
            TauxC = IIf(dictTaux.exists(DeviseFiltre), dictTaux(DeviseFiltre), 1)
            
            For i = 1 To UBound(arrFact, 1)
                If Trim(CStr(arrFact(i, 1))) <> "" And IsDate(arrFact(i, 2)) Then
                    dTx = CDate(arrFact(i, 2))
                    
                    ' FILTRE TEMPOREL
                    If Format(dTx, "yyyy-mm") = MoisFiltre Then
                        Ligne = Ligne + 1
                        IDCompte = Trim(CStr(arrFact(i, 3))): idCat = Trim(CStr(arrFact(i, 4))): IDTiers = Trim(CStr(arrFact(i, 5)))
                        TypeFlux = IIf(dictCatType.exists(idCat), dictCatType(idCat), "AUTRE")
                        
                        ' CONVERSION DE DEVISE O(1)
                        DevOrigine = UCase(Trim(CStr(arrFact(i, 7))))
                        TauxO = IIf(dictTaux.exists(DevOrigine), dictTaux(DevOrigine), 1)
                        Montant = CDbl(arrFact(i, 6))
                        
                        ' Mathématique de conversion : (Montant * TauxOrigineEnMur) / TauxCibleEnMur
                        MontantConverti = (Montant * TauxO) / TauxC
                        
                        If UCase(TypeFlux) = "REVENU" Then TotRev = TotRev + MontantConverti
                        If UCase(TypeFlux) = "DEPENSE" Then TotDep = TotDep + MontantConverti
                        
                        arrConsolide(Ligne, 1) = arrFact(i, 1): arrConsolide(Ligne, 2) = arrFact(i, 2)
                        arrConsolide(Ligne, 3) = Year(dTx): arrConsolide(Ligne, 4) = Month(dTx)
                        arrConsolide(Ligne, 5) = IIf(dictCompte.exists(IDCompte), dictCompte(IDCompte), "-")
                        arrConsolide(Ligne, 6) = IIf(dictCat.exists(idCat), dictCat(idCat), "-")
                        arrConsolide(Ligne, 7) = IIf(dictTiers.exists(IDTiers), dictTiers(IDTiers), "-")
                        arrConsolide(Ligne, 8) = MontantConverti
                        arrConsolide(Ligne, 9) = DeviseFiltre ' Affiche la devise convertie
                        arrConsolide(Ligne, 10) = TypeFlux: arrConsolide(Ligne, 11) = arrFact(i, 8)
                    End If
                End If
            Next i
        End If
    End If
    
    If Ligne = 0 Then
        ReDim arrConsolide(1 To 1, 1 To 11)
        arrConsolide(1, 1) = "-": arrConsolide(1, 2) = "-": arrConsolide(1, 3) = "-": arrConsolide(1, 4) = "-"
        arrConsolide(1, 5) = "-": arrConsolide(1, 6) = "-": arrConsolide(1, 7) = "-": arrConsolide(1, 8) = 0
        arrConsolide(1, 9) = "-": arrConsolide(1, 10) = "-": arrConsolide(1, 11) = "-"
        Ligne = 1
    End If
    
    ' --- CALIBRAGE DE LA GRILLE (11 Colonnes, de C à M) ---
    wsDash.Columns("A:B").ColumnWidth = 2
    wsDash.Columns("C").ColumnWidth = 4   ' ID
    wsDash.Columns("D").ColumnWidth = 8  ' Date
    wsDash.Columns("E").ColumnWidth = 6   ' Année
    wsDash.Columns("F").ColumnWidth = 4   ' Mois
    wsDash.Columns("G").ColumnWidth = 20  ' Compte
    wsDash.Columns("H").ColumnWidth = 30  ' Catégorie
    wsDash.Columns("I").ColumnWidth = 18  ' Tiers
    wsDash.Columns("J").ColumnWidth = 15  ' Montant
    wsDash.Columns("K").ColumnWidth = 6   ' Devise
    wsDash.Columns("L").ColumnWidth = 12  ' Type
    wsDash.Columns("M").ColumnWidth = 30  ' Notes
    
    ' --- DESSIN DES SOLID CARDS (Alignées sur la grille) ---
    wsDash.Rows("7").RowHeight = 35: wsDash.Rows("8").RowHeight = 50
    
    Dim ZoneTable As Range: Set ZoneTable = wsDash.Range("C7:M7")
    Dim TotalW As Double: TotalW = ZoneTable.Width
    Dim Gap As Double: Gap = 15
    Dim CardW As Double: CardW = (TotalW - (2 * Gap)) / 3
    
    Dessiner_Solid_Card wsDash, "CARD_INC", TR("KPI_INC") & " (" & DeviseFiltre & ")", TotRev, RGB(65, 105, 225), vbWhite, ZoneTable.Left, wsDash.Range("C7").Top, CardW, 85
    Dessiner_Solid_Card wsDash, "CARD_EXP", TR("KPI_EXP") & " (" & DeviseFiltre & ")", TotDep, RGB(120, 81, 169), vbWhite, ZoneTable.Left + CardW + Gap, wsDash.Range("C7").Top, CardW, 85
    Dessiner_Solid_Card wsDash, "CARD_NET", TR("KPI_NET") & " (" & DeviseFiltre & ")", TotRev - TotDep, RGB(250, 218, 94), RGB(40, 40, 40), ZoneTable.Left + (CardW * 2) + (Gap * 2), wsDash.Range("C7").Top, CardW, 85
    
    wsDash.Rows("9:11").RowHeight = 15
    
    ' --- DESSIN DU TABLEAU "VIOLET ZEBRA" ---
    Dim Headers As Variant
    Headers = Array("ID", UCase(TR("FRM_DATE")), UCase(TR("COL_YEAR")), UCase(TR("COL_MONTH")), UCase(TR("FRM_COMPTE")), UCase(TR("FRM_CAT")), UCase(TR("FRM_TIERS")), UCase(TR("FRM_MONTANT")), UCase(TR("FRM_DEVISE")), UCase(TR("COL_TYPE")), UCase(TR("FRM_DESC")))
    
    wsDash.Range("C12:M12").Value = Headers
    wsDash.Range("C13").Resize(Ligne, 11).Value = arrConsolide
    
    Dim tblView As ListObject
    Set tblView = wsDash.ListObjects.Add(xlSrcRange, wsDash.Range("C12").Resize(Ligne + 1, 11), , xlYes)
    tblView.Name = "VIEW_Transactions"
    tblView.TableStyle = ""
    tblView.ShowAutoFilterDropDown = False
    
    ' En-tête Violet Sombre
    With tblView.HeaderRowRange
        .Interior.Color = RGB(90, 50, 130)
        .Font.Color = vbWhite: .Font.Bold = True
        .RowHeight = 35: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlNone
    End With
    
    ' Corps Violet Zebra
    With tblView.DataBodyRange
        .RowHeight = 28: .VerticalAlignment = xlCenter: .Borders.LineStyle = xlNone: .Font.Color = vbWhite
        Dim r As Long: For r = 1 To .Rows.Count
            If r Mod 2 = 0 Then .Rows(r).Interior.Color = RGB(145, 110, 190) Else .Rows(r).Interior.Color = RGB(120, 81, 169)
        Next r
    End With
    
    If arrConsolide(1, 8) <> 0 Then ' S'il y a des données réelles
        tblView.ListColumns(1).DataBodyRange.HorizontalAlignment = xlCenter
        tblView.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter
        tblView.ListColumns(3).DataBodyRange.HorizontalAlignment = xlCenter
        tblView.ListColumns(4).DataBodyRange.HorizontalAlignment = xlCenter
        tblView.ListColumns(8).DataBodyRange.HorizontalAlignment = xlRight
        tblView.ListColumns(9).DataBodyRange.HorizontalAlignment = xlCenter
        tblView.ListColumns(10).DataBodyRange.HorizontalAlignment = xlCenter
        
        On Error Resume Next
        tblView.ListColumns(8).DataBodyRange.NumberFormat = "#,##0.00"
        On Error GoTo 0
    Else
        ' FIX 1004 : CenterAcrossSelection pour l'Empty State
        With tblView.DataBodyRange
            .RowHeight = 35: .ClearContents
        End With
        With wsDash.Range("C13:M13")
            .HorizontalAlignment = xlCenterAcrossSelection: .VerticalAlignment = xlCenter
            .Font.Italic = True: .Font.Color = RGB(220, 220, 225)
        End With
        wsDash.Range("C13").Value = TR("NO_DATA")
    End If
    
    wsDash.Range("A1").Select
End Sub

' -------------------------------------------------------------------------
' 4. MOTEUR INTERACTIF & ANIMATIONS TACTILES
' -------------------------------------------------------------------------
Public Sub CHANGER_DEVISE()
    Dim btnName As String: On Error Resume Next: btnName = Application.Caller: On Error GoTo 0
    If btnName <> "" Then
        Modifier_Parametre "DASH_FILTRE_DEV", Replace(btnName, "BTN_DEV_", "")
        Rafraichir_Dashboard
    End If
End Sub

Public Sub MOIS_PRECEDENT()
    Anim_Btn_Bleu Application.Caller
    Modifier_Mois_Filtre -1
    Rafraichir_Dashboard
End Sub

Public Sub MOIS_SUIVANT()
    Anim_Btn_Bleu Application.Caller
    Modifier_Mois_Filtre 1
    Rafraichir_Dashboard
End Sub

Private Sub Modifier_Mois_Filtre(DeltaMois As Integer)
    Dim actuel As String: actuel = Obtenir_Parametre("DASH_FILTRE_MOIS", Format(Date, "yyyy-mm"))
    Dim d As Date: d = DateAdd("m", DeltaMois, CDate(actuel & "-01"))
    Modifier_Parametre "DASH_FILTRE_MOIS", Format(d, "yyyy-mm")
End Sub

Private Sub Rafraichir_Dashboard()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    GENERER_DASHBOARD
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

Public Sub ANIMATION_RETOUR()
    Dim btn As Shape: On Error Resume Next: Set btn = ActiveSheet.Shapes(Application.Caller): On Error GoTo 0
    If Not btn Is Nothing Then
        btn.Fill.ForeColor.RGB = RGB(220, 190, 60): btn.Shadow.Visible = msoFalse: btn.Top = btn.Top + 2: btn.Left = btn.Left + 2
        Dim t As Single: t = Timer: Do While Timer < t + 0.15: DoEvents: Loop
        btn.Fill.ForeColor.RGB = RGB(250, 218, 94): btn.Shadow.Visible = msoTrue: btn.Top = btn.Top - 2: btn.Left = btn.Left - 2
    End If
    On Error Resume Next: ThisWorkbook.Sheets("APP_HOME").Activate: ThisWorkbook.Sheets("APP_HOME").Range("A1").Select: On Error GoTo 0
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

Private Function Charger_Dico(NomTable As String, ColKey As Integer, ColVal As Integer) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim tbl As ListObject: On Error Resume Next: Set tbl = ThisWorkbook.Sheets(Split(NomTable, "_", 2)(1)).ListObjects(NomTable): On Error GoTo 0
    If Not tbl Is Nothing Then
        Dim i As Long: For i = 1 To tbl.ListRows.Count
            If Trim(CStr(tbl.DataBodyRange(i, ColKey).Value)) <> "" Then dict(Trim(CStr(tbl.DataBodyRange(i, ColKey).Value))) = Trim(CStr(tbl.DataBodyRange(i, ColVal).Value))
        Next i
    End If
    Set Charger_Dico = dict
End Function

Private Sub Dessiner_Solid_Card(ws As Worksheet, NomShape As String, Titre As String, Valeur As Double, CoulFond As Long, CoulTexte As Long, Gauche As Double, Haut As Double, Largeur As Double, Hauteur As Double)
    Dim shp As Shape: Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, Gauche, Haut, Largeur, Hauteur)
    shp.Name = NomShape: shp.Fill.ForeColor.RGB = CoulFond: shp.Line.Visible = msoFalse
    With shp.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 6: .OffsetX = 0: .OffsetY = 3: .Transparency = 0.6: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    shp.TextFrame2.TextRange.Text = UCase(Titre) & vbCrLf & Format(Valeur, "#,##0.00")
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle: shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With shp.TextFrame2.TextRange.Lines(1).Font: .Name = "ADLaM Display": .Size = 12: .Bold = True: .Fill.ForeColor.RGB = CoulTexte: End With
    With shp.TextFrame2.TextRange.Lines(2).Font: .Name = "ADLaM Display": .Size = 30: .Bold = True: .Fill.ForeColor.RGB = CoulTexte: End With
End Sub
