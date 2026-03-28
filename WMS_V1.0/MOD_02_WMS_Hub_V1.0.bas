Attribute VB_Name = "MOD_02_WMS_Hub"
Option Explicit

' =========================================================================
' MODULE: MOD_02_WMS_Hub
' OBJECTIF: Hub Central WMS, SPA Premium, ADLaM Display 10, UCR
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_3_HUB()
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect "WMS_ADMIN_2026"
    Next ws
    
    Preparer_WMS_Hub
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "LE HUB CENTRAL WMS EST DÉPLOYÉ." & vbCrLf & vbCrLf & _
           "1. L'esthétique 'Solid Cards' 3D est appliquée." & vbCrLf & _
           "2. Le moteur de routage U.C.R est prêt.", vbInformation, "WMS v1.0 - Étape 3"
End Sub

Private Sub Preparer_WMS_Hub()
    Dim wsHome As Worksheet
    On Error Resume Next: Set wsHome = ThisWorkbook.Sheets("WMS_HOME"): On Error GoTo 0
    
    If wsHome Is Nothing Then
        Set wsHome = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsHome.Name = "WMS_HOME"
    Else
        wsHome.Cells.Clear
        Dim shp As Shape: For Each shp In wsHome.Shapes: shp.Delete: Next shp
        wsHome.Hyperlinks.Delete
    End If
    
    ' --- FORÇAGE DU ZOOM ET POLICE ---
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    wsHome.Cells.Font.Name = "ADLaM Display"
    wsHome.Cells.Font.Size = 10
    wsHome.Cells.Interior.Color = RGB(248, 248, 250)
    
    ' --- BANDEAU SUPÉRIEUR ---
    wsHome.Range("A1:Z5").Interior.Color = RGB(65, 105, 225) ' Bleu Royal
    
    ' --- TITRE VECTORIEL ---
    Dim shpTitle As Shape
    Set shpTitle = wsHome.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 15, 600, 50)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = "WEALTH MANAGEMENT SYSTEM" & vbCrLf & "Portfolio & Market Analytics | " & Format(Date, "dd mmmm yyyy")
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 22: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 11: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' --- MESSAGE D'ACCUEIL ---
    wsHome.Range("C8").Value = "Sélectionnez un module pour gérer vos investissements."
    wsHome.Range("C8").Font.Color = RGB(150, 150, 150): wsHome.Range("C8").Font.Italic = True
    
    ' --- LES TUILES DE NAVIGATION (SOLID CARDS) ---
    Dim T_Top As Integer: T_Top = 160
    Dim T_Left As Integer: T_Left = 100
    Dim T_W As Integer: T_W = 380
    Dim T_H As Integer: T_H = 110
    Dim Gap As Integer: Gap = 30
    
    ' 1. Saisie Trade (Jaune Royal)
    Dessiner_Tuile_WMS wsHome, "EXÉCUTER UN ORDRE" & vbCrLf & "Achat, Vente, Dividendes", T_Left, T_Top, T_W, T_H, RGB(250, 218, 94), RGB(40, 40, 40), "A21"
    
    ' 2. Portfolio Dashboard (Violet Royal)
    Dessiner_Tuile_WMS wsHome, "PERFORMANCE PORTFOLIO" & vbCrLf & "Valorisation & Plus-Values", T_Left + T_W + Gap, T_Top, T_W, T_H, RGB(120, 81, 169), vbWhite, "A22"
    
    ' 3. Market Analytics (Vert Émeraude)
    Dessiner_Tuile_WMS wsHome, "ANALYSE DE MARCHÉ" & vbCrLf & "Suivi des cotations (API)", T_Left, T_Top + T_H + Gap, T_W, T_H, RGB(46, 204, 113), vbWhite, "A23"
    
    wsHome.Activate
    wsHome.Range("A1").Select
End Sub

Private Sub Dessiner_Tuile_WMS(ws As Worksheet, Texte As String, Gauche As Integer, Haut As Integer, Largeur As Integer, Hauteur As Integer, CoulFond As Long, CoulTexte As Long, CelluleCible As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, Gauche, Haut, Largeur, Hauteur)
    btn.Fill.ForeColor.RGB = CoulFond
    btn.Line.Visible = msoFalse
    
    With btn.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow
        .Blur = 8: .OffsetX = 0: .OffsetY = 4: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    btn.TextFrame2.TextRange.Text = Texte
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CoulTexte
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    
    With btn.TextFrame2.TextRange.Lines(1).Font
        .Name = "ADLaM Display": .Bold = True: .Size = 16
    End With
    With btn.TextFrame2.TextRange.Lines(2).Font
        .Name = "ADLaM Display": .Bold = False: .Size = 11
    End With
    
    ws.Hyperlinks.Add Anchor:=btn, Address:="", SubAddress:="'" & ws.Name & "'!" & CelluleCible
End Sub

' --- ACTIONS EXÉCUTABLES PAR LE ROUTEUR ---
Public Sub EXECUTER_ROUTER_TRADE()
    On Error GoTo ErrForm
    USF_Trade.Show
    Exit Sub
ErrForm:
    MsgBox "Le Formulaire est introuvable. Exécutez l'Étape 2.", vbCritical
End Sub

Public Sub EXECUTER_ROUTER_PORTFOLIO()
    MsgBox "Le Dashboard de Performance sera généré à l'Étape 4 (Power Query & Data Model).", vbInformation, "Bientôt disponible"
End Sub

Public Sub EXECUTER_ROUTER_MARKET()
    MsgBox "L'analyse de marché sera alimentée par l'API Boursière à l'Étape 5.", vbInformation, "Bientôt disponible"
End Sub

