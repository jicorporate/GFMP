Option Explicit

' =========================================================================
' MODULE: MOD_05_Advanced_Modules
' OBJECTIF: Bilan Patrimonial, Time Slider (Snapshot), Multi-Devises, Violet Zebra
' =========================================================================

Public Sub DEPLOIEMENT_ETAPE_6_NETWORTH()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    
    Garantir_Lexique_NetWorth
    GENERER_NET_WORTH_DASHBOARD
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
    
    MsgBox "LE MOTEUR TEMPOREL PATRIMONIAL EST DÉPLOYÉ." & vbCrLf & vbCrLf & _
           "1. Le Time Slider a été ajouté au bandeau." & vbCrLf & _
           "2. Logique Snapshot : Le système calcule votre patrimoine 'À date' (toutes les transactions jusqu'au mois choisi)." & vbCrLf & _
           "3. Zéro Régression : Violet Zebra, Bouton Jaune, et Devises sont intacts.", vbInformation, "SFP v3.2 - Time Engine"
End Sub

' -------------------------------------------------------------------------
' 1. STATE MANAGEMENT (Devise & Mois Actifs)
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
' 2. MOTEUR AUTO-CICATRISANT DU PATRIMOINE
' -------------------------------------------------------------------------
Private Sub Garantir_Lexique_NetWorth()
    Dim tblDic As ListObject
    On Error Resume Next: Set tblDic = ThisWorkbook.Sheets("SYS_Config").ListObjects("T_SYS_Dictionary"): On Error GoTo 0
    If tblDic Is Nothing Then Exit Sub
    
    Upsert_Dico tblDic, "KPI_ASSETS", "TOTAL ACTIFS", "TOTAL ASSETS", "TOTAL ACTIVOS", "TOTAL DE ATIVOS", "GESAMTVERMÖGEN", "TOTALE ATTIVITÀ", "TOTALE ACTIVA", "TOTALA TILLGÅNGAR"
    Upsert_Dico tblDic, "KPI_LIAB", "TOTAL PASSIFS", "TOTAL LIABILITIES", "TOTAL PASIVOS", "TOTAL DE PASSIVOS", "GESAMTVERBINDLICHKEITEN", "TOTALE PASSIVITÀ", "TOTALE PASSIVA", "TOTALA SKULDER"
    Upsert_Dico tblDic, "KPI_NW", "VALEUR NETTE", "NET WORTH", "PATRIMONIO NETO", "PATRIMÔNIO LÍQUIDO", "NETTOVERMÖGEN", "PATRIMONIO NETTO", "NETTOWAARDE", "NETTOFÖRMÖGENHET"
    
    Upsert_Dico tblDic, "NW_COMPTE", "COMPTE FINANCIER", "FINANCIAL ACCOUNT", "CUENTA FINANCIERA", "CONTA FINANCEIRA", "FINANZKONTO", "CONTO FINANZIARIO", "FINANCIËLE REKENING", "FINANSIELLT KONTO"
    Upsert_Dico tblDic, "NW_CLASS", "CLASSE D'ACTIF", "ASSET CLASS", "CLASE DE ACTIVO", "CLASSE DE ATIVO", "ANLAGEKLASSE", "CLASSE DI INVESTIMENTO", "ACTIVA KLASSE", "TILLGÅNGSKLASS"
    Upsert_Dico tblDic, "NW_BALANCE", "SOLDE", "BALANCE", "SALDO", "SALDO", "GUTHABEN", "SALDO", "SALDO", "SALDO"
    
    Upsert_Dico tblDic, "FRM_DEVISE", "Devise", "Currency", "Divisa", "Moeda", "Währung", "Valuta", "Valuta", "Valuta"
    Upsert_Dico tblDic, "BTN_BACK", "RETOUR AU MENU", "BACK TO MENU", "VOLVER AL MENÚ", "VOLTAR AO MENU", "ZURÜCK ZUM MENÜ", "TORNA AL MENU", "TERUG NAAR MENU", "TILLBAKA TILL MENY"
    Upsert_Dico tblDic, "NO_DATA", "Aucun compte actif à cette date", "No active accounts on this date", "Sin cuentas en esta fecha", "Sem contas nesta data", "Keine Konten", "Nessun conto", "Geen rekeningen", "Inga konton"
End Sub

Private Sub Upsert_Dico(tbl As ListObject, k As String, fr As String, en As String, es As String, pt As String, de As String, it As String, nl As String, sv As String)
    Dim i As Long: For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 1).Value = k Then
            tbl.DataBodyRange(i, 2).Value = fr: tbl.DataBodyRange(i, 3).Value = en: tbl.DataBodyRange(i, 4).Value = es
            tbl.DataBodyRange(i, 5).Value = pt: tbl.DataBodyRange(i, 6).Value = de: tbl.DataBodyRange(i, 7).Value = it
            tbl.DataBodyRange(i, 8).Value = nl: tbl.DataBodyRange(i, 9).Value = sv
            Exit Sub
        End If
    Next i
    Dim nr As ListRow: Set nr = tbl.ListRows.Add
    nr.Range(1, 1).Value = k: nr.Range(1, 2).Value = fr: nr.Range(1, 3).Value = en: nr.Range(1, 4).Value = es
    nr.Range(1, 5).Value = pt: nr.Range(1, 6).Value = de: nr.Range(1, 7).Value = it: nr.Range(1, 8).Value = nl: nr.Range(1, 9).Value = sv
End Sub

Private Function TR(Cle As String) As String
    TR = MOD_02_AppHome_Global.TR(Cle)
End Function

' -------------------------------------------------------------------------
' 3. MOTEUR DE RENDU UI (Time Slider, Devises, Violet Zebra)
' -------------------------------------------------------------------------
Public Sub GENERER_NET_WORTH_DASHBOARD()
    Garantir_Lexique_NetWorth
    
    Dim DeviseFiltre As String: DeviseFiltre = Obtenir_Parametre("NW_FILTRE_DEV", "MUR")
    Dim MoisFiltre As String: MoisFiltre = Obtenir_Parametre("NW_FILTRE_MOIS", Format(Date, "yyyy-mm"))

    Dim wsNW As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next: ThisWorkbook.Sheets("NET_WORTH").Delete: On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsNW = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("APP_HOME"))
    wsNW.Name = "NET_WORTH"
    wsNW.Activate
    
    ' --- 1. POLICE GLOBALE ET ZOOM 100% ---
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
    wsNW.Cells.Font.Name = "ADLaM Display"
    wsNW.Cells.Font.Size = 10
    wsNW.Cells.Interior.Color = RGB(248, 248, 250)
    
    ' --- 2. BANDEAU SUPÉRIEUR ---
    wsNW.Range("A1:Z5").Interior.Color = RGB(65, 105, 225) ' Bleu Royal
    
    ' --- 3. BOUTON RETOUR TACTILE (Jaune Royal) ---
    Dim btnBack As Shape
    Set btnBack = wsNW.Shapes.AddShape(msoShapeRoundedRectangle, 20, 15, 150, 30)
    btnBack.Name = "BTN_RETOUR_TACTILE_NW"
    btnBack.Fill.ForeColor.RGB = RGB(250, 218, 94) ' Jaune Royal
    btnBack.Line.Visible = msoFalse
    btnBack.TextFrame2.WordWrap = msoFalse: btnBack.TextFrame2.MarginLeft = 0: btnBack.TextFrame2.MarginRight = 0
    With btnBack.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 4: .OffsetX = 1: .OffsetY = 2: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    btnBack.TextFrame2.TextRange.Text = "<  " & TR("BTN_BACK")
    btnBack.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnBack.TextFrame2.TextRange.Font.Bold = True: btnBack.TextFrame2.TextRange.Font.Size = 10: btnBack.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    btnBack.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnBack.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnBack.OnAction = "MOD_05_Advanced_Modules.ANIMATION_RETOUR_TACTILE_NW"
    
    ' --- 4. TITRE VECTORIEL ---
    Dim shpTitle As Shape
    Set shpTitle = wsNW.Shapes.AddTextbox(msoTextOrientationHorizontal, 210, 10, 300, 40)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = UCase(TR("NETW_T")) & vbCrLf & "Snapshot Mode"
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 18: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 10: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' --- 5. LE TIME SLIDER (Machine à remonter le temps) ---
    Dim LabelDate As String: LabelDate = UCase(Format(CDate(MoisFiltre & "-01"), "mmmm yyyy"))
    Dessiner_Widget wsNW, "BTN_NW_PREV_MONTH", "<", 290, 15, 25, 32, RGB(220, 220, 220), RGB(0, 0, 0), "MOD_05_Advanced_Modules.MOIS_PRECEDENT_NW"
    Dessiner_Widget wsNW, "LBL_NW_MONTH", LabelDate, 320, 15, 160, 32, RGB(220, 220, 220), RGB(0, 0, 0), ""
    Dessiner_Widget wsNW, "BTN_NW_NEXT_MONTH", ">", 485, 15, 25, 32, RGB(220, 220, 220), RGB(0, 0, 0), "MOD_05_Advanced_Modules.MOIS_SUIVANT_NW"
    
    ' --- 6. LE CONVERTISSEUR DE DEVISES ---
    Dim devLeft As Integer: devLeft = 540
    Dessiner_Widget wsNW, "BTN_NW_DEV_MUR", "MUR", devLeft, 15, 50, 32, IIf(DeviseFiltre = "MUR", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "MUR", RGB(40, 40, 40), vbWhite), "MOD_05_Advanced_Modules.CHANGER_DEVISE_NW"
    Dessiner_Widget wsNW, "BTN_NW_DEV_EUR", "EUR", devLeft + 55, 15, 50, 32, IIf(DeviseFiltre = "EUR", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "EUR", RGB(40, 40, 40), vbWhite), "MOD_05_Advanced_Modules.CHANGER_DEVISE_NW"
    Dessiner_Widget wsNW, "BTN_NW_DEV_USD", "USD", devLeft + 110, 15, 50, 32, IIf(DeviseFiltre = "USD", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "USD", RGB(40, 40, 40), vbWhite), "MOD_05_Advanced_Modules.CHANGER_DEVISE_NW"
    Dessiner_Widget wsNW, "BTN_NW_DEV_XOF", "XOF", devLeft + 165, 15, 50, 32, IIf(DeviseFiltre = "XOF", RGB(250, 218, 94), RGB(40, 70, 180)), IIf(DeviseFiltre = "XOF", RGB(40, 40, 40), vbWhite), "MOD_05_Advanced_Modules.CHANGER_DEVISE_NW"
    
    ' --- MOTEUR DE TAUX DE CHANGE ---
    Dim dictTaux As Object: Set dictTaux = CreateObject("Scripting.Dictionary")
    dictTaux("MUR") = 1: dictTaux("EUR") = 49.5: dictTaux("USD") = 46.2: dictTaux("GBP") = 58.1: dictTaux("ZAR") = 2.4: dictTaux("XOF") = 0.083
    
    ' --- 7. PHASE D'EXTRACTION (ETL CUMULATIF ACCOUNT-CENTRIC CORRIGÉ) ---
    Dim tblCompte As ListObject, tblFact As ListObject
    On Error Resume Next
    Set tblCompte = ThisWorkbook.Sheets("DIM_Compte").ListObjects("T_DIM_Compte")
    Set tblFact = ThisWorkbook.Sheets("FACT_Transaction").ListObjects("T_FACT_Transaction")
    On Error GoTo 0
    
    Dim TotAssets As Double: TotAssets = 0
    Dim TotLiab As Double: TotLiab = 0
    Dim arrConsolide() As Variant
    Dim Ligne As Long: Ligne = 0
    
    If Not tblCompte Is Nothing Then
        If tblCompte.ListRows.Count > 0 Then
            Dim arrCpt As Variant: arrCpt = tblCompte.DataBodyRange.Value
            Dim arrTx As Variant
            Dim aDesTransactions As Boolean: aDesTransactions = False
            
            If Not tblFact Is Nothing Then
                If tblFact.ListRows.Count > 0 Then
                    If tblFact.ListColumns.Count >= 7 Then
                        arrTx = tblFact.DataBodyRange.Value
                        aDesTransactions = True
                    End If
                End If
            End If
            
            ReDim arrConsolide(1 To UBound(arrCpt, 1), 1 To 4)
            Dim i As Long, j As Long
            Dim ID_Cpt As String, TypeCpt As String, StatutActif As String
            Dim SoldeNatif As Double, SoldeConvertiKPI As Double
            Dim FluxType As String, CatTypeDict As Object
            Dim CptDevise As String, TauxC_Native As Double, TauxFiltre As Double
            Dim dTx As Date
            
            Set CatTypeDict = Charger_Dico_Interne("T_DIM_Categorie", 1, 3)
            ' Taux d'affichage global du Dashboard
            TauxFiltre = IIf(dictTaux.exists(DeviseFiltre), dictTaux(DeviseFiltre), 1)
            
            For i = 1 To UBound(arrCpt, 1)
                ' =========================================================================
                ' CORRECTION 1 : RÉSOLUTION DU "STATUT FANTÔME" (COMPTES OMIS)
                ' Si la colonne Est_Actif (5) est vide (à cause de l'auto-apprentissage),
                ' on force la valeur à "OUI" pour qu'il soit reconnu par le Bilan.
                ' =========================================================================
                StatutActif = ""
                If UBound(arrCpt, 2) >= 5 Then StatutActif = UCase(Trim(CStr(arrCpt(i, 5))))
                If StatutActif = "" Then StatutActif = "OUI"
                
                If StatutActif = "OUI" Then
                    ID_Cpt = Trim(CStr(arrCpt(i, 1)))
                    TypeCpt = UCase(Trim(CStr(arrCpt(i, 3))))
                    
                    ' =========================================================================
                    ' CORRECTION 2 : PRIORITÉ À LA DEVISE NATIVE DU COMPTE (DIMENSION)
                    ' La devise par défaut du compte est la seule qui fait foi.
                    ' =========================================================================
                    CptDevise = UCase(Trim(CStr(arrCpt(i, 4))))
                    If CptDevise = "" Then CptDevise = "MUR"
                    TauxC_Native = IIf(dictTaux.exists(CptDevise), dictTaux(CptDevise), 1)
                    
                    SoldeNatif = 0
                    
                    If aDesTransactions Then
                        For j = 1 To UBound(arrTx, 1)
                            If Trim(CStr(arrTx(j, 3))) = ID_Cpt And Trim(CStr(arrTx(j, 1))) <> "" Then
                                If IsDate(arrTx(j, 2)) Then
                                    dTx = CDate(arrTx(j, 2))
                                    If Format(dTx, "yyyy-mm") <= MoisFiltre Then
                                        
                                        ' =========================================================================
                                        ' CORRECTION 3 : SUPPRESSION DU CONFLIT DE DEVISE (LA TRANSACTION EST IGNORÉE)
                                        ' On assume que le Montant saisi EST dans la Devise Native du Compte.
                                        ' On additionne directement le montant sans conversion hasardeuse.
                                        ' =========================================================================
                                        FluxType = IIf(CatTypeDict.exists(Trim(CStr(arrTx(j, 4)))), CatTypeDict(Trim(CStr(arrTx(j, 4)))), "AUTRE")
                                        
                                        If UCase(FluxType) = "REVENU" Or UCase(FluxType) = "TRANSFERT" Then
                                            SoldeNatif = SoldeNatif + CDbl(arrTx(j, 6))
                                        ElseIf UCase(FluxType) = "DEPENSE" Then
                                            SoldeNatif = SoldeNatif - CDbl(arrTx(j, 6))
                                        End If
                                        
                                    End If
                                End If
                            End If
                        Next j
                    End If
                    
                    If SoldeNatif <> 0 Then
                        ' =========================================================================
                        ' CONVERSION GLOBALE DES KPIs (À la toute fin)
                        ' Le Solde Natif est converti dans la devise du Dashboard pour les Totaux.
                        ' =========================================================================
                        SoldeConvertiKPI = (SoldeNatif * TauxC_Native) / TauxFiltre
                        
                        If TypeCpt = "DETTE" Then
                            TotLiab = TotLiab + Abs(SoldeConvertiKPI)
                            SoldeNatif = -Abs(SoldeNatif) ' Le tableau affiche le solde négatif
                        Else
                            TotAssets = TotAssets + SoldeConvertiKPI
                        End If
                        
                        Ligne = Ligne + 1
                        arrConsolide(Ligne, 1) = arrCpt(i, 2)
                        arrConsolide(Ligne, 2) = TypeCpt
                        arrConsolide(Ligne, 3) = CptDevise ' La table affiche la devise NATIVE (ex: XOF)
                        arrConsolide(Ligne, 4) = SoldeNatif ' La table affiche le montant NATIF (ex: 20000)
                    End If
                End If
            Next i
        End If
    End If
    
    ' EMPTY STATE (Si aucun compte n'a de solde à cette date)
    If Ligne = 0 Then
        ReDim arrConsolide(1 To 1, 1 To 4)
        arrConsolide(1, 1) = TR("NO_DATA"): arrConsolide(1, 2) = "-": arrConsolide(1, 3) = "-": arrConsolide(1, 4) = 0
        Ligne = 1
    End If
    
    ' --- 8. CALIBRAGE DE LA GRILLE (Zoom 100%) ---
    wsNW.Columns("A:B").ColumnWidth = 2
    wsNW.Columns("C").ColumnWidth = 40   ' Compte
    wsNW.Columns("D").ColumnWidth = 30   ' Classe d'actif
    wsNW.Columns("E").ColumnWidth = 15   ' Devise
    wsNW.Columns("F").ColumnWidth = 30   ' Solde Actuel
    
    ' --- 9. DESSIN MATHÉMATIQUE DES KPIS ---
    wsNW.Rows("7").RowHeight = 35
    wsNW.Rows("8").RowHeight = 50
    
    Dim ZoneTable As Range: Set ZoneTable = wsNW.Range("C7:F7")
    Dim TotalW As Double: TotalW = ZoneTable.Width
    Dim Gap As Double: Gap = 15
    Dim CardW As Double: CardW = (TotalW - (2 * Gap)) / 3
    Dim Y_Pos As Double: Y_Pos = wsNW.Range("C7").Top
    Dim H_Card As Double: H_Card = 85
    
    ' COULEUR CONDITIONNELLE DU NET WORTH
    Dim NetWorth As Double: NetWorth = TotAssets - TotLiab
    Dim NW_Color As Long
    If NetWorth > 0 Then
        NW_Color = RGB(46, 204, 113) ' Vert
    ElseIf NetWorth < 0 Then
        NW_Color = RGB(231, 76, 60) ' Rouge
    Else
        NW_Color = RGB(128, 128, 128) ' Gris
    End If
    
    Dessiner_Shape_Card wsNW, "CARD_ASSETS", TR("KPI_ASSETS") & " (" & DeviseFiltre & ")", TotAssets, RGB(65, 105, 225), vbWhite, ZoneTable.Left, Y_Pos, CardW, H_Card
    Dessiner_Shape_Card wsNW, "CARD_LIAB", TR("KPI_LIAB") & " (" & DeviseFiltre & ")", TotLiab, RGB(120, 81, 169), vbWhite, ZoneTable.Left + CardW + Gap, Y_Pos, CardW, H_Card
    Dessiner_Shape_Card wsNW, "CARD_NW", TR("KPI_NW") & " (" & DeviseFiltre & ")", NetWorth, NW_Color, vbWhite, ZoneTable.Left + (CardW * 2) + (Gap * 2), Y_Pos, CardW, H_Card
    
    wsNW.Rows("9:11").RowHeight = 15
    
    ' --- 10. DESSIN DE LA TABLE "VIOLET ZEBRA" ---
    Dim Headers As Variant
    Headers = Array(UCase(TR("NW_COMPTE")), UCase(TR("NW_CLASS")), UCase(TR("FRM_DEVISE")), UCase(TR("NW_BALANCE")))
    
    wsNW.Range("C12:F12").Value = Headers
    wsNW.Range("C13").Resize(Ligne, 4).Value = arrConsolide
    
    Dim tblView As ListObject
    If arrConsolide(1, 1) <> TR("NO_DATA") Then
        Set tblView = wsNW.ListObjects.Add(xlSrcRange, wsNW.Range("C12").Resize(Ligne + 1, 4), , xlYes)
    Else
        Set tblView = wsNW.ListObjects.Add(xlSrcRange, wsNW.Range("C12:F13"), , xlYes)
        With wsNW.Range("C13:F13")
            .HorizontalAlignment = xlCenterAcrossSelection: .VerticalAlignment = xlCenter
            .Font.Italic = True: .Font.Color = RGB(220, 220, 225)
        End With
        wsNW.Range("C13").Value = TR("NO_DATA")
    End If
    
    tblView.Name = "VIEW_NetWorth"
    tblView.TableStyle = ""
    tblView.ShowAutoFilterDropDown = False
    
    With tblView.HeaderRowRange
        .Interior.Color = RGB(90, 50, 130)
        .Font.Color = vbWhite: .Font.Bold = True
        .RowHeight = 35: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlNone
    End With
    
    If arrConsolide(1, 1) <> TR("NO_DATA") Then
        With tblView.DataBodyRange
            .RowHeight = 28: .VerticalAlignment = xlCenter: .Borders.LineStyle = xlNone
            .Font.Color = vbWhite ' Texte Blanc
            
            Dim r As Long
            For r = 1 To .Rows.Count
                If r Mod 2 = 0 Then .Rows(r).Interior.Color = RGB(145, 110, 190) Else .Rows(r).Interior.Color = RGB(120, 81, 169)
                ' Dettes en Jaune Royal
                If IsNumeric(.Cells(r, 4).Value) Then
                    If CDbl(.Cells(r, 4).Value) < 0 Then
                        .Cells(r, 4).Font.Color = RGB(250, 218, 94): .Cells(r, 4).Font.Bold = True
                    End If
                End If
            Next r
        End With
        
        tblView.ListColumns(1).DataBodyRange.HorizontalAlignment = xlLeft
        tblView.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter
        tblView.ListColumns(3).DataBodyRange.HorizontalAlignment = xlCenter
        tblView.ListColumns(4).DataBodyRange.HorizontalAlignment = xlRight
        
        On Error Resume Next
        tblView.ListColumns(4).DataBodyRange.NumberFormat = "#,##0.00"
        On Error GoTo 0
    End If
    
    wsNW.Range("A1").Select
End Sub

' -------------------------------------------------------------------------
' 4. MOTEUR INTERACTIF (Time Slider, Devises, Retour Tactile)
' -------------------------------------------------------------------------
Public Sub MOIS_PRECEDENT_NW()
    Anim_Btn_Bleu Application.Caller
    Modifier_Parametre "NW_FILTRE_MOIS", Format(DateAdd("m", -1, CDate(Obtenir_Parametre("NW_FILTRE_MOIS", Format(Date, "yyyy-mm")) & "-01")), "yyyy-mm")
    Rafraichir_NW_Dashboard
End Sub

Public Sub MOIS_SUIVANT_NW()
    Anim_Btn_Bleu Application.Caller
    Modifier_Parametre "NW_FILTRE_MOIS", Format(DateAdd("m", 1, CDate(Obtenir_Parametre("NW_FILTRE_MOIS", Format(Date, "yyyy-mm")) & "-01")), "yyyy-mm")
    Rafraichir_NW_Dashboard
End Sub

Public Sub CHANGER_DEVISE_NW()
    Dim btnName As String: On Error Resume Next: btnName = Application.Caller: On Error GoTo 0
    If btnName <> "" Then
        Modifier_Parametre "NW_FILTRE_DEV", Replace(btnName, "BTN_NW_DEV_", "")
        Rafraichir_NW_Dashboard
    End If
End Sub

Private Sub Rafraichir_NW_Dashboard()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    GENERER_NET_WORTH_DASHBOARD
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

Public Sub ANIMATION_RETOUR_TACTILE_NW()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim btn As Shape: On Error Resume Next: Set btn = ws.Shapes(Application.Caller): On Error GoTo 0
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
    With shp.TextFrame2.TextRange.Font: .Name = "ADLaM Display": .Size = 10: .Bold = True: .Fill.ForeColor.RGB = cTxt: End With
    If Macro <> "" Then shp.OnAction = Macro
End Sub

Private Function Charger_Dico_Interne(NomTable As String, ColKey As Integer, ColVal As Integer) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim tbl As ListObject: On Error Resume Next: Set tbl = ThisWorkbook.Sheets(Split(NomTable, "_", 2)(1)).ListObjects(NomTable): On Error GoTo 0
    If Not tbl Is Nothing Then
        Dim i As Long, Cle As String, Val As String
        For i = 1 To tbl.ListRows.Count
            Cle = Trim(CStr(tbl.DataBodyRange(i, ColKey).Value)): Val = Trim(CStr(tbl.DataBodyRange(i, ColVal).Value))
            If Cle <> "" Then dict(Cle) = Val
        Next i
    End If
    Set Charger_Dico_Interne = dict
End Function

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



