Attribute VB_Name = "MOD_06_Market_Dashboard"
Option Explicit

' =========================================================================
' MODULE: MOD_06_Market_Dashboard
' OBJECTIF: Explorateur de Marché Boursier (Watchlist & Market Data)
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_7_MARKET()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "WMS_ADMIN_2026": Next ws
    
    Generer_Dashboard_Market
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
    
    MsgBox "LE DASHBOARD D'ANALYSE DE MARCHÉ EST OPÉRATIONNEL !" & vbCrLf & vbCrLf & _
           "Il surveille désormais tous vos actifs via les données de clôture importées par Power Query.", vbInformation, "WMS v1.0 - Étape 7"
End Sub

Public Sub Generer_Dashboard_Market()
    ' --- DEBUT PATCH 1 (Génération Idempotente Market Dash) ---
    Dim wsDash As Worksheet
    On Error Resume Next: Set wsDash = ThisWorkbook.Sheets("MARKET_DASH"): On Error GoTo 0
    
    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("WMS_HOME"))
        wsDash.Name = "MARKET_DASH"
    Else
        wsDash.Visible = xlSheetVisible
        wsDash.Cells.Clear
        Dim shp As Shape: For Each shp In wsDash.Shapes: shp.Delete: Next shp
        Dim tbl As ListObject: For Each tbl In wsDash.ListObjects: tbl.Delete: Next tbl
    End If
    wsDash.Activate
    ' --- FIN PATCH 1 ---
    
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
    btnBack.Name = "BTN_RETOUR_HUB_MKT"
    btnBack.Fill.ForeColor.RGB = RGB(250, 218, 94) ' Jaune Royal
    btnBack.Line.Visible = msoFalse
    btnBack.TextFrame2.TextRange.Text = "<  RETOUR AU HUB"
    btnBack.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnBack.TextFrame2.TextRange.Font.Bold = True: btnBack.TextFrame2.TextRange.Font.Size = 10: btnBack.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    btnBack.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnBack.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With btnBack.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 4: .OffsetX = 0: .OffsetY = 2: .Transparency = 0.5
    End With
    btnBack.OnAction = "MOD_06_Market_Dashboard.ANIMATION_RETOUR_MKT"
    
    ' --- 4. TITRE VECTORIEL ---
    Dim shpTitle As Shape
    Set shpTitle = wsDash.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 10, 400, 40)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = "ANALYSE DE MARCHÉ" & vbCrLf & "Explorateur des Cotations | " & Format(Date, "dd mmm yyyy")
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 18: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 10: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' --- 5. EXTRACTION DAX (LECTURE DU BIG DATA) ---
    Dim conn As Object, rs As Object
    Dim Ligne As Long: Ligne = 0
    Dim arrData() As Variant
    
    On Error Resume Next
    Set conn = ThisWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Requête DAX : On extrait tous les actifs et leur dernier prix de clôture connu
    Dim daxQuery As String
    daxQuery = "EVALUATE SUMMARIZECOLUMNS('T_DIM_Asset'[Ticker_Symbole], 'T_DIM_Asset'[Nom_Actif], 'T_DIM_Asset'[Classe_Actif], 'T_DIM_Asset'[Devise_Cotation], ""Price"", [Current_Price])"
    
    rs.Open daxQuery, conn
    
    If Not rs.EOF Then
        Dim rawData As Variant
        rawData = rs.GetRows
        Dim numCols As Long: numCols = UBound(rawData, 1)
        Dim numRows As Long: numRows = UBound(rawData, 2)
        ReDim arrData(1 To numRows + 1, 1 To 5)
        
        Dim r As Long
        For r = 0 To numRows
            Ligne = Ligne + 1
            arrData(Ligne, 1) = rawData(0, r) ' Ticker
            arrData(Ligne, 2) = rawData(1, r) ' Nom
            arrData(Ligne, 3) = rawData(2, r) ' Classe
            arrData(Ligne, 4) = rawData(3, r) ' Devise Origine
            If IsNull(rawData(4, r)) Then
                arrData(Ligne, 5) = "Non disponible"
            Else
                arrData(Ligne, 5) = CDbl(rawData(4, r)) ' Dernier Prix
            End If
        Next r
    End If
    rs.Close
    On Error GoTo 0
    
    ' --- GESTION EMPTY STATE ---
    If Ligne = 0 Then
        ReDim arrData(1 To 1, 1 To 5)
        arrData(1, 1) = "Aucun actif": arrData(1, 2) = "-": arrData(1, 3) = "-": arrData(1, 4) = "-": arrData(1, 5) = 0
        Ligne = 1
    End If
    
    ' --- 6. CALIBRAGE DE LA GRILLE ---
    wsDash.Columns("A:B").ColumnWidth = 2
    wsDash.Columns("C").ColumnWidth = 15  ' Ticker
    wsDash.Columns("D").ColumnWidth = 40  ' Nom
    wsDash.Columns("E").ColumnWidth = 20  ' Classe
    wsDash.Columns("F").ColumnWidth = 15  ' Devise
    wsDash.Columns("G").ColumnWidth = 20  ' Dernier Prix
    
    wsDash.Rows("7").RowHeight = 20 ' Marge
    
    ' --- 7. TABLEAU "VIOLET ZEBRA" ---
    wsDash.Range("C8:G8").Value = Array("TICKER", "ACTIF", "CLASSE D'ACTIF", "DEVISE (NATIVE)", "DERNIER PRIX DE CLÔTURE")
    wsDash.Range("C9").Resize(Ligne, 5).Value = arrData
    
    Dim tblView As ListObject
    Set tblView = wsDash.ListObjects.Add(xlSrcRange, wsDash.Range("C8").Resize(Ligne + 1, 5), , xlYes)
    tblView.Name = "VIEW_Market"
    tblView.TableStyle = ""
    tblView.ShowAutoFilterDropDown = False
    
    ' En-tête Violet Sombre
    With tblView.HeaderRowRange
        .Interior.Color = RGB(90, 50, 130): .Font.Color = vbWhite: .Font.Bold = True
        .RowHeight = 35: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlNone
    End With
    
    If arrData(1, 1) <> "Aucun actif" Then
        With tblView.DataBodyRange
            .RowHeight = 28: .VerticalAlignment = xlCenter: .Borders.LineStyle = xlNone: .Font.Color = vbWhite
            Dim rIdx As Long: For rIdx = 1 To .Rows.Count
                If rIdx Mod 2 = 0 Then .Rows(rIdx).Interior.Color = RGB(145, 110, 190) Else .Rows(rIdx).Interior.Color = RGB(120, 81, 169)
                ' Format
                If IsNumeric(.Cells(rIdx, 5).Value) Then .Cells(rIdx, 5).NumberFormat = "#,##0.00"
            Next rIdx
        End With
    End If
    
    wsDash.Range("A1").Select
End Sub

Public Sub ANIMATION_RETOUR_MKT()
    Dim btn As Shape: On Error Resume Next: Set btn = ActiveSheet.Shapes(Application.Caller): On Error GoTo 0
    If Not btn Is Nothing Then
        btn.Fill.ForeColor.RGB = RGB(220, 190, 60): btn.Shadow.Visible = msoFalse: btn.Top = btn.Top + 2: btn.Left = btn.Left + 2
        Dim t As Single: t = Timer: Do While Timer < t + 0.15: DoEvents: Loop
        btn.Fill.ForeColor.RGB = RGB(250, 218, 94): btn.Shadow.Visible = msoTrue: btn.Top = btn.Top - 2: btn.Left = btn.Left - 2
    End If
    On Error Resume Next: ThisWorkbook.Sheets("WMS_HOME").Activate: ThisWorkbook.Sheets("WMS_HOME").Range("A1").Select: On Error GoTo 0
End Sub

