Attribute VB_Name = "MOD_05_Portfolio_Dashboard"
Option Explicit

' =========================================================================
' MODULE: MOD_05_Portfolio_Dashboard
' OBJECTIF: Interrogation DAX en RAM & Rendu UI (Violet Zebra, Solid Cards)
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_6_DASHBOARD()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "WMS_ADMIN_2026": Next ws
    
    Generer_Dashboard_Performance
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
    
    MsgBox "LE DASHBOARD DE PERFORMANCE EST OPÉRATIONNEL !" & vbCrLf & vbCrLf & _
           "1. Le moteur DAX extrait les données en RAM (O(1))." & vbCrLf & _
           "2. L'interface Violet Zebra et les Solid Cards sont générées.", vbInformation, "WMS v1.0 - Étape 6"
End Sub

Public Sub Generer_Dashboard_Performance()
    ' --- DEBUT PATCH 2 (Génération Idempotente Portfolio Dash) ---
    Dim wsDash As Worksheet
    On Error Resume Next: Set wsDash = ThisWorkbook.Sheets("PORTFOLIO_DASH"): On Error GoTo 0
    
    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("WMS_HOME"))
        wsDash.Name = "PORTFOLIO_DASH"
    Else
        wsDash.Visible = xlSheetVisible
        wsDash.Cells.Clear
        Dim shp As Shape: For Each shp In wsDash.Shapes: shp.Delete: Next shp
        Dim tbl As ListObject: For Each tbl In wsDash.ListObjects: tbl.Delete: Next tbl
    End If
    wsDash.Activate
    ' --- FIN PATCH 2 ---
    
    ' --- 1. FORÇAGE DE LA POLICE GLOBALE ET DU ZOOM 100% ---
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    wsDash.Cells.Font.Name = "ADLaM Display"
    wsDash.Cells.Font.Size = 10
    wsDash.Cells.Interior.Color = RGB(248, 248, 250)
    
    ' --- 2. BANDEAU SUPÉRIEUR ---
    wsDash.Range("A1:Z5").Interior.Color = RGB(65, 105, 225) ' Bleu Royal
    
    ' --- 3. BOUTON RETOUR TACTILE ---
    Dim btnBack As Shape
    Set btnBack = wsDash.Shapes.AddShape(msoShapeRoundedRectangle, 20, 15, 160, 32)
    btnBack.Name = "BTN_RETOUR_HUB"
    btnBack.Fill.ForeColor.RGB = RGB(250, 218, 94) ' Jaune Royal
    btnBack.Line.Visible = msoFalse
    btnBack.TextFrame2.TextRange.Text = "<  RETOUR AU HUB"
    btnBack.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnBack.TextFrame2.TextRange.Font.Bold = True: btnBack.TextFrame2.TextRange.Font.Size = 10: btnBack.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    btnBack.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnBack.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With btnBack.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 4: .OffsetX = 0: .OffsetY = 2: .Transparency = 0.5
    End With
    btnBack.OnAction = "MOD_05_Portfolio_Dashboard.ANIMATION_RETOUR"
    
    ' --- 4. TITRE VECTORIEL ---
    Dim shpTitle As Shape
    Set shpTitle = wsDash.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 10, 400, 40)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = "PERFORMANCE PORTFOLIO" & vbCrLf & "Valorisation en temps réel | " & Format(Date, "dd mmm yyyy")
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 18: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 10: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' --- 5. EXTRACTION DAX (LA PUISSANCE DU DATA MODEL EN RAM) ---
    Dim conn As Object, rs As Object
    Dim TotInvested As Double, TotValue As Double, TotPnL As Double
    Dim Ligne As Long: Ligne = 0
    Dim arrData() As Variant
    'Dim baseDev As String: baseDev = MOD_00_WMS_Architecture.Obtenir_Parametre("SYS_DEVISE_BASE", "USD")
    ' --- DEBUT PATCH (Appel de la fonction globale depuis MOD_03) ---
    Dim baseDev As String: baseDev = MOD_03_Market_ETL.Obtenir_Parametre("SYS_DEVISE_BASE", "USD")
    ' --- FIN PATCH ---
    
    ' Reconnexion silencieuse au Moteur Power Pivot
    On Error Resume Next
    Set conn = ThisWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Requête DAX absolue (Évalue uniquement les actifs avec des parts > 0)
    Dim daxQuery As String
    daxQuery = "EVALUATE FILTER(SUMMARIZECOLUMNS('T_DIM_Asset'[Ticker_Symbole], 'T_DIM_Asset'[Nom_Actif], ""Shares"", [Total_Shares], ""Invested"", [Invested_Capital], ""Price"", [Current_Price], ""Value"", [Market_Value], ""PnL"", [Unrealized_PnL]), [Shares] <> 0)"
    
    rs.Open daxQuery, conn
    
    If Not rs.EOF Then
        Dim rawData As Variant
        rawData = rs.GetRows ' Récupère tout en mémoire instantanément
        Dim numCols As Long: numCols = UBound(rawData, 1)
        Dim numRows As Long: numRows = UBound(rawData, 2)
        ReDim arrData(1 To numRows + 1, 1 To 7)
        
        Dim r As Long
        For r = 0 To numRows
            Ligne = Ligne + 1
            arrData(Ligne, 1) = rawData(0, r) ' Ticker
            arrData(Ligne, 2) = rawData(1, r) ' Nom
            arrData(Ligne, 3) = CDbl(rawData(2, r)) ' Shares
            arrData(Ligne, 4) = CDbl(rawData(3, r)) ' Invested
            arrData(Ligne, 5) = CDbl(rawData(4, r)) ' Current Price
            arrData(Ligne, 6) = CDbl(rawData(5, r)) ' Market Value
            arrData(Ligne, 7) = CDbl(rawData(6, r)) ' Unrealized PnL
            
            TotInvested = TotInvested + arrData(Ligne, 4)
            TotValue = TotValue + arrData(Ligne, 6)
            TotPnL = TotPnL + arrData(Ligne, 7)
        Next r
    End If
    rs.Close
    On Error GoTo 0
    
    ' --- GESTION EMPTY STATE ---
    If Ligne = 0 Then
        ReDim arrData(1 To 1, 1 To 7)
        arrData(1, 1) = "Aucune position active": arrData(1, 2) = "-": arrData(1, 3) = 0: arrData(1, 4) = 0: arrData(1, 5) = 0: arrData(1, 6) = 0: arrData(1, 7) = 0
        Ligne = 1
    End If
    
    ' --- 6. CALIBRAGE DE LA GRILLE (100% Zoom) ---
    wsDash.Columns("A:B").ColumnWidth = 2
    wsDash.Columns("C").ColumnWidth = 15  ' Ticker
    wsDash.Columns("D").ColumnWidth = 30  ' Nom
    wsDash.Columns("E").ColumnWidth = 15  ' Shares
    wsDash.Columns("F").ColumnWidth = 20  ' Invested
    wsDash.Columns("G").ColumnWidth = 15  ' Price
    wsDash.Columns("H").ColumnWidth = 20  ' Market Value
    wsDash.Columns("I").ColumnWidth = 20  ' PnL
    
    ' --- 7. DESSIN DES KPIS (SOLID CARDS 3D) ---
    wsDash.Rows("7").RowHeight = 35: wsDash.Rows("8").RowHeight = 50
    Dim ZoneTable As Range: Set ZoneTable = wsDash.Range("C7:I7")
    Dim CardW As Double: CardW = (ZoneTable.Width - 30) / 3
    
    Dim PnLColor As Long
    ' --- DEBUT PATCH (Correction Syntaxe Bloc If) ---
    If TotPnL > 0 Then
        PnLColor = RGB(46, 204, 113)
    ElseIf TotPnL < 0 Then
        PnLColor = RGB(231, 76, 60)
    Else
        PnLColor = RGB(128, 128, 128)
    End If
    ' --- FIN PATCH ---
    
    Dessiner_Card wsDash, "TOTAL INVESTI (" & baseDev & ")", TotInvested, RGB(65, 105, 225), vbWhite, ZoneTable.Left, wsDash.Range("C7").Top, CardW, 85
    Dessiner_Card wsDash, "VALEUR DE MARCHÉ (" & baseDev & ")", TotValue, RGB(120, 81, 169), vbWhite, ZoneTable.Left + CardW + 15, wsDash.Range("C7").Top, CardW, 85
    Dessiner_Card wsDash, "PLUS-VALUE LATENTE (" & baseDev & ")", TotPnL, PnLColor, vbWhite, ZoneTable.Left + (CardW * 2) + 30, wsDash.Range("C7").Top, CardW, 85
    
    wsDash.Rows("9:11").RowHeight = 15
    
    ' --- 8. TABLEAU "VIOLET ZEBRA" ---
    wsDash.Range("C12:I12").Value = Array("TICKER", "ACTIF", "QTÉ", "INVESTI (" & baseDev & ")", "PRIX ACTUEL", "VALORISATION", "PLUS-VALUE")
    wsDash.Range("C13").Resize(Ligne, 7).Value = arrData
    
    Dim tblView As ListObject
    Set tblView = wsDash.ListObjects.Add(xlSrcRange, wsDash.Range("C12").Resize(Ligne + 1, 7), , xlYes)
    tblView.Name = "VIEW_Portfolio"
    tblView.TableStyle = ""
    tblView.ShowAutoFilterDropDown = False
    
    ' En-tête Violet Sombre
    With tblView.HeaderRowRange
        .Interior.Color = RGB(90, 50, 130): .Font.Color = vbWhite: .Font.Bold = True
        .RowHeight = 35: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlNone
    End With
    
    If arrData(1, 1) <> "Aucune position active" Then
        With tblView.DataBodyRange
            .RowHeight = 28: .VerticalAlignment = xlCenter: .Borders.LineStyle = xlNone: .Font.Color = vbWhite
            Dim rIdx As Long: For rIdx = 1 To .Rows.Count
                If rIdx Mod 2 = 0 Then .Rows(rIdx).Interior.Color = RGB(145, 110, 190) Else .Rows(rIdx).Interior.Color = RGB(120, 81, 169)
                ' Format monétaire
                .Cells(rIdx, 3).NumberFormat = "#,##0.0000"
                .Cells(rIdx, 4).NumberFormat = "#,##0.00"
                .Cells(rIdx, 5).NumberFormat = "#,##0.00"
                .Cells(rIdx, 6).NumberFormat = "#,##0.00"
                .Cells(rIdx, 7).NumberFormat = "#,##0.00"
                ' Couleurs PnL
                If CDbl(.Cells(rIdx, 7).Value) > 0 Then
                    .Cells(rIdx, 7).Font.Color = RGB(46, 204, 113): .Cells(rIdx, 7).Font.Bold = True
                ElseIf CDbl(.Cells(rIdx, 7).Value) < 0 Then
                    .Cells(rIdx, 7).Font.Color = RGB(250, 218, 94): .Cells(rIdx, 7).Font.Bold = True ' Jaune Alerte sur fond violet
                End If
            Next rIdx
        End With
    End If
    
    wsDash.Range("A1").Select
End Sub

Private Sub Dessiner_Card(ws As Worksheet, Titre As String, Valeur As Double, CoulFond As Long, CoulTexte As Long, L As Double, t As Double, W As Double, H As Double)
    Dim shp As Shape: Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, t, W, H)
    shp.Fill.ForeColor.RGB = CoulFond: shp.Line.Visible = msoFalse
    With shp.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 6: .OffsetY = 3: .Transparency = 0.6
    End With
    shp.TextFrame2.TextRange.Text = Titre & vbCrLf & Format(Valeur, "#,##0.00")
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle: shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shp.TextFrame2.TextRange.Lines(1).Font.Size = 11: shp.TextFrame2.TextRange.Lines(1).Font.Bold = True: shp.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = CoulTexte
    shp.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shp.TextFrame2.TextRange.Lines(2).Font.Size = 28: shp.TextFrame2.TextRange.Lines(2).Font.Bold = True: shp.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = CoulTexte
End Sub

Public Sub ANIMATION_RETOUR()
    Dim btn As Shape: On Error Resume Next: Set btn = ActiveSheet.Shapes(Application.Caller): On Error GoTo 0
    If Not btn Is Nothing Then
        btn.Fill.ForeColor.RGB = RGB(220, 190, 60): btn.Shadow.Visible = msoFalse: btn.Top = btn.Top + 2: btn.Left = btn.Left + 2
        Dim t As Single: t = Timer: Do While Timer < t + 0.15: DoEvents: Loop
        btn.Fill.ForeColor.RGB = RGB(250, 218, 94): btn.Shadow.Visible = msoTrue: btn.Top = btn.Top - 2: btn.Left = btn.Left - 2
    End If
    On Error Resume Next: ThisWorkbook.Sheets("WMS_HOME").Activate: ThisWorkbook.Sheets("WMS_HOME").Range("A1").Select: On Error GoTo 0
End Sub

