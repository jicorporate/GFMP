Attribute VB_Name = "MOD_02_WMS_Hub"
Option Explicit

' =========================================================================
' MODULE: MOD_02_WMS_Hub
' OBJECTIF: Hub Central WMS, SPA Premium, Moteur i18n Multilingue O(1)
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_3_HUB()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "WMS_ADMIN_2026": Next ws
    
    Preparer_Dictionnaire_WMS
    Preparer_WMS_Hub
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
    MsgBox "HUB INTERNATIONAL WMS DÉPLOYÉ." & vbCrLf & "Le moteur multilingue i18n et les boutons de langue sont actifs.", vbInformation, "WMS v1.0 - i18n"
End Sub

' -------------------------------------------------------------------------
' MOTEUR DE DICTIONNAIRE MULTILINGUE
' -------------------------------------------------------------------------
Private Sub Preparer_Dictionnaire_WMS()
    Dim wsSys As Worksheet, tblDic As ListObject, tblConf As ListObject
    Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    Set tblConf = wsSys.ListObjects("T_SYS_Config")
    
    Dim langExist As Boolean: langExist = False
    Dim i As Long: For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = "LANGUE_UI" Then langExist = True: Exit For
    Next i
    If Not langExist Then
        Dim nrConf As ListRow: Set nrConf = tblConf.ListRows.Add
        nrConf.Range(1, 1).Value = "LANGUE_UI": nrConf.Range(1, 2).Value = GET_SYSTEM_LANGUAGE(): nrConf.Range(1, 3).Value = "Langue UI Globale"
    End If

    On Error Resume Next: Set tblDic = wsSys.ListObjects("T_SYS_Dictionary"): On Error GoTo 0
    If tblDic Is Nothing Then
        wsSys.Columns("E:M").Clear
        wsSys.Range("E1:M1").Value = Array("KEY", "FR", "EN", "ES", "PT", "DE", "IT", "NL", "SV")
        Set tblDic = wsSys.ListObjects.Add(xlSrcRange, wsSys.Range("E1:M2"), , xlYes)
        tblDic.Name = "T_SYS_Dictionary"
        tblDic.TableStyle = "TableStyleMedium15": tblDic.ListRows(1).Delete
    End If
    
    Upsert_Trad tblDic, "WMS_TITLE", "WEALTH MANAGEMENT SYSTEM", "WEALTH MANAGEMENT SYSTEM", "SISTEMA DE GESTIÓN PATRIMONIAL", "SISTEMA DE GESTÃO PATRIMONIAL", "VERMÖGENSVERWALTUNG", "GESTIONE PATRIMONIALE", "VERMOGENSBEHEER", "FÖRMÖGENHETSHANTERING"
    Upsert_Trad tblDic, "WMS_SUB", "Portfolio & Market Analytics", "Portfolio & Market Analytics", "Análisis de Cartera y Mercado", "Análise de Portfólio e Mercado", "Portfolio- & Marktanalysen", "Analisi di Portafoglio e Mercato", "Portfolio & Marktanalyse", "Portfölj- och Marknadsanalys"
    Upsert_Trad tblDic, "WMS_WELC", "Sélectionnez un module pour gérer vos investissements.", "Select a module to manage your investments.", "Seleccione un módulo para gestionar sus inversiones.", "Selecione um módulo para gerenciar seus investimentos.", "Wählen Sie ein Modul zur Verwaltung.", "Seleziona un modulo per gestire i tuoi investimenti.", "Selecteer een module.", "Välj en modul för att hantera dina investeringar."
    Upsert_Trad tblDic, "CARD_T_T", "EXÉCUTER UN ORDRE", "EXECUTE TRADE", "EJECUTAR ORDEN", "EXECUTAR ORDEM", "ORDER AUSFÜHREN", "ESEGUI ORDINE", "ORDER UITVOEREN", "UTFÖR ORDER"
    Upsert_Trad tblDic, "CARD_T_D", "Achat, Vente, Dividendes", "Buy, Sell, Dividends", "Compra, Venta, Dividendos", "Compra, Venda, Dividendos", "Kauf, Verkauf, Dividenden", "Compra, Vendita, Dividendi", "Kopen, Verkopen, Dividenden", "Köp, Sälj, Utdelning"
    Upsert_Trad tblDic, "CARD_P_T", "PERFORMANCE PORTFOLIO", "PORTFOLIO PERFORMANCE", "RENDIMIENTO DE CARTERA", "DESEMPENHO DO PORTFÓLIO", "PORTFOLIO-PERFORMANCE", "PERFORMANCE PORTAFOGLIO", "PORTFOLIO PRESTATIES", "PORTFÖLJENS UTVECKLING"
    Upsert_Trad tblDic, "CARD_P_D", "Valorisation & Plus-Values", "Valuation & Capital Gains", "Valoración y Plusvalías", "Avaliação e Mais-Valias", "Bewertung & Kapitalgewinne", "Valutazione e Plusvalenze", "Waardering & Vermogenswinst", "Värdering & Kapitalvinst"
    Upsert_Trad tblDic, "CARD_M_T", "ANALYSE DE MARCHÉ", "MARKET ANALYTICS", "ANÁLISIS DE MERCADO", "ANÁLISE DE MERCADO", "MARKTANALYSE", "ANALISI DI MERCATO", "MARKTANALYSE", "MARKNADSANALYS"
    Upsert_Trad tblDic, "CARD_M_D", "Suivi des cotations (API)", "Market quotes tracking (API)", "Seguimiento de cotizaciones (API)", "Acompanhamento de cotações (API)", "Verfolgung von Marktnotierungen", "Monitoraggio delle quotazioni (API)", "Marktkoersen volgen", "Marknadsnoteringar (API)"
End Sub

Private Sub Upsert_Trad(tbl As ListObject, k As String, fr As String, en As String, es As String, pt As String, de As String, it As String, nl As String, sv As String)
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

Public Function TR(Clé As String) As String
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    Dim tblConf As ListObject: Set tblConf = wsSys.ListObjects("T_SYS_Config")
    Dim tblDic As ListObject: Set tblDic = wsSys.ListObjects("T_SYS_Dictionary")
    
    Dim Langue As String: Langue = "FR"
    Dim i As Long: For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = "LANGUE_UI" Then Langue = tblConf.DataBodyRange(i, 2).Value: Exit For
    Next i
    
    Dim ColIdx As Integer
    Select Case Langue
        Case "FR": ColIdx = 2: Case "EN": ColIdx = 3: Case "ES": ColIdx = 4: Case "PT": ColIdx = 5
        Case "DE": ColIdx = 6: Case "IT": ColIdx = 7: Case "NL": ColIdx = 8: Case "SV": ColIdx = 9
        Case Else: ColIdx = 3 ' Fallback EN
    End Select
    
    For i = 1 To tblDic.ListRows.Count
        If tblDic.DataBodyRange(i, 1).Value = Clé Then TR = tblDic.DataBodyRange(i, ColIdx).Value: Exit Function
    Next i
    TR = Clé
End Function

Private Function GET_SYSTEM_LANGUAGE() As String
    Dim lcid As Long: On Error Resume Next: lcid = Application.LanguageSettings.LanguageID(2): On Error GoTo 0
    Select Case lcid
        Case 1036, 2060, 3084, 4108, 5132, 6156: GET_SYSTEM_LANGUAGE = "FR"
        Case 1034, 2058, 3082, 4106, 5130: GET_SYSTEM_LANGUAGE = "ES"
        Case 1046, 2070: GET_SYSTEM_LANGUAGE = "PT"
        Case 1031, 2055, 3079, 4103, 5127: GET_SYSTEM_LANGUAGE = "DE"
        Case 1040, 2064: GET_SYSTEM_LANGUAGE = "IT"
        Case 1043, 2067: GET_SYSTEM_LANGUAGE = "NL"
        Case 1053, 2077: GET_SYSTEM_LANGUAGE = "SV"
        Case Else: GET_SYSTEM_LANGUAGE = "EN"
    End Select
End Function

' -------------------------------------------------------------------------
' DESSIN DU HUB MULTILINGUE
' -------------------------------------------------------------------------
Private Sub Preparer_WMS_Hub()
    Dim wsHome As Worksheet: On Error Resume Next: Set wsHome = ThisWorkbook.Sheets("WMS_HOME"): On Error GoTo 0
    If wsHome Is Nothing Then
        Set wsHome = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1)): wsHome.Name = "WMS_HOME"
    Else
        wsHome.Cells.Clear: Dim shp As Shape: For Each shp In wsHome.Shapes: shp.Delete: Next shp: wsHome.Hyperlinks.Delete
    End If
    
    ActiveWindow.DisplayGridlines = False: ActiveWindow.DisplayHeadings = False: ActiveWindow.Zoom = 100
    wsHome.Cells.Font.Name = "ADLaM Display": wsHome.Cells.Font.Size = 10: wsHome.Cells.Interior.Color = RGB(248, 248, 250)
    wsHome.Range("A1:Z5").Interior.Color = RGB(65, 105, 225) ' Bleu Royal
    
    Dim shpTitle As Shape
    Set shpTitle = wsHome.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 15, 600, 50)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = TR("WMS_TITLE") & vbCrLf & TR("WMS_SUB") & " | " & Format(Date, "dd mmmm yyyy")
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 22: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 11: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' LES BOUTONS DE LANGUES SONT DE RETOUR !
    Dim arrLang As Variant: arrLang = Array("FR", "EN", "ES", "PT", "DE", "IT", "NL", "SV")
    Dim i As Integer, xPos As Integer: xPos = 770
    For i = LBound(arrLang) To UBound(arrLang)
        Dessiner_Bouton_Langue wsHome, CStr(arrLang(i)), xPos, 30, 35, 35, "A" & (11 + i)
        xPos = xPos + 40
    Next i
    
    wsHome.Range("C8").Value = TR("WMS_WELC")
    wsHome.Range("C8").Font.Color = RGB(150, 150, 150): wsHome.Range("C8").Font.Italic = True
    
    Dim T_Top As Integer: T_Top = 160: Dim T_Left As Integer: T_Left = 100: Dim T_W As Integer: T_W = 380: Dim T_H As Integer: T_H = 110: Dim Gap As Integer: Gap = 30
    Dessiner_Tuile_WMS wsHome, TR("CARD_T_T") & vbCrLf & TR("CARD_T_D"), T_Left, T_Top, T_W, T_H, RGB(250, 218, 94), RGB(40, 40, 40), "A21"
    Dessiner_Tuile_WMS wsHome, TR("CARD_P_T") & vbCrLf & TR("CARD_P_D"), T_Left + T_W + Gap, T_Top, T_W, T_H, RGB(120, 81, 169), vbWhite, "A22"
    Dessiner_Tuile_WMS wsHome, TR("CARD_M_T") & vbCrLf & TR("CARD_M_D"), T_Left, T_Top + T_H + Gap, T_W, T_H, RGB(46, 204, 113), vbWhite, "A23"
    wsHome.Activate: wsHome.Range("A1").Select
End Sub

Private Sub Dessiner_Bouton_Langue(ws As Worksheet, Texte As String, Gauche As Integer, Haut As Integer, Largeur As Integer, Hauteur As Integer, CelluleCible As String)
    Dim btn As Shape: Set btn = ws.Shapes.AddShape(msoShapeOval, Gauche, Haut, Largeur, Hauteur)
    btn.Fill.ForeColor.RGB = RGB(40, 70, 180): btn.Line.ForeColor.RGB = vbWhite: btn.Line.Weight = 1.5
    btn.TextFrame2.WordWrap = msoFalse: btn.TextFrame2.MarginLeft = 0: btn.TextFrame2.MarginRight = 0
    btn.TextFrame2.TextRange.Text = Texte: btn.TextFrame2.TextRange.Font.Name = "ADLaM Display": btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite: btn.TextFrame2.TextRange.Font.Bold = True: btn.TextFrame2.TextRange.Font.Size = 10
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle: btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    ws.Hyperlinks.Add Anchor:=btn, Address:="", SubAddress:="'" & ws.Name & "'!" & CelluleCible
End Sub

Private Sub Dessiner_Tuile_WMS(ws As Worksheet, Texte As String, Gauche As Integer, Haut As Integer, Largeur As Integer, Hauteur As Integer, CoulFond As Long, CoulTexte As Long, CelluleCible As String)
    Dim btn As Shape: Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, Gauche, Haut, Largeur, Hauteur)
    btn.Fill.ForeColor.RGB = CoulFond: btn.Line.Visible = msoFalse
    With btn.Shadow: .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 8: .OffsetY = 4: .Transparency = 0.5: End With
    btn.TextFrame2.TextRange.Text = Texte: btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CoulTexte: btn.TextFrame2.VerticalAnchor = msoAnchorMiddle: btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btn.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": btn.TextFrame2.TextRange.Lines(1).Font.Bold = True: btn.TextFrame2.TextRange.Lines(1).Font.Size = 16
    btn.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": btn.TextFrame2.TextRange.Lines(2).Font.Bold = False: btn.TextFrame2.TextRange.Lines(2).Font.Size = 11
    ws.Hyperlinks.Add Anchor:=btn, Address:="", SubAddress:="'" & ws.Name & "'!" & CelluleCible
End Sub

Public Sub EXECUTER_CHANGER_LANGUE(LangueCible As String)
    Application.EnableEvents = True: Application.ScreenUpdating = False
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config"): wsSys.Unprotect "WMS_ADMIN_2026"
    Dim tblConf As ListObject: Set tblConf = wsSys.ListObjects("T_SYS_Config")
    Dim i As Long: For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = "LANGUE_UI" Then tblConf.DataBodyRange(i, 2).Value = LangueCible: Exit For
    Next i
    wsSys.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "WMS_ADMIN_2026": Next ws
    Preparer_WMS_Hub
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

Public Sub EXECUTER_ROUTER_TRADE(): On Error Resume Next: USF_Trade.Show: On Error GoTo 0: End Sub
Public Sub EXECUTER_ROUTER_PORTFOLIO(): MOD_05_Portfolio_Dashboard.DEPLOYER_WMS_ETAPE_6_DASHBOARD: End Sub
Public Sub EXECUTER_ROUTER_MARKET(): MOD_06_Market_Dashboard.DEPLOYER_WMS_ETAPE_7_MARKET: End Sub
