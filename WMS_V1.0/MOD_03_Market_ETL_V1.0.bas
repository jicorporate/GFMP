Attribute VB_Name = "MOD_03_Market_ETL"
Option Explicit

' =========================================================================
' MODULE: MOD_03_Market_ETL
' OBJECTIF: Générateur de M-Code (Power Query) pour le Pricing Boursier Massif
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_4_ETL()
    Application.ScreenUpdating = False
    
    ' 1. Lecture des Tickers de vos Actifs (DIM_Asset)
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets("DIM_Asset")
    Dim tblA As ListObject: Set tblA = wsA.ListObjects("T_DIM_Asset")
    Dim i As Long
    Dim listTickers As String: listTickers = ""
    Dim listFX As String: listFX = ""
    'Dim baseCurrency As String: baseCurrency = MOD_00_WMS_Architecture.Obtenir_Parametre("SYS_DEVISE_BASE", "USD")
    ' --- DEBUT PATCH (Appel de la fonction locale) ---
    Dim baseCurrency As String: baseCurrency = Obtenir_Parametre("SYS_DEVISE_BASE", "USD")
    ' --- FIN PATCH ---
    Dim assetCur As String
    Dim dictFX As Object: Set dictFX = CreateObject("Scripting.Dictionary")
    
    If tblA.ListRows.Count = 0 Then
        MsgBox "Aucun actif dans le référentiel. Ajoutez des actifs avant de lancer le moteur de prix.", vbExclamation
        Exit Sub
    End If
    
    For i = 1 To tblA.ListRows.Count
        ' Ajout du Ticker pour Yahoo Finance (ex: AAPL, BTC-USD)
        If listTickers <> "" Then listTickers = listTickers & ", "
        listTickers = listTickers & """" & Trim(CStr(tblA.DataBodyRange(i, 2).Value)) & """"
        
        ' Détection des Devises Étrangères pour créer les paires FX nécessaires (ex: EURUSD=X)
        assetCur = UCase(Trim(CStr(tblA.DataBodyRange(i, 5).Value)))
        If assetCur <> baseCurrency Then
            If Not dictFX.exists(assetCur) Then
                dictFX.Add assetCur, True
                If listFX <> "" Then listFX = listFX & ", "
                listFX = listFX & """" & assetCur & baseCurrency & "=X""" ' Format Yahoo pour le FX
            End If
        End If
    Next i
    
    ' 2. Création ou Mise à Jour de la Requête Power Query (FACT_MarketQuotes)
    Dim M_Code_Quotes As String
    M_Code_Quotes = "let" & vbCrLf & _
        "    Tickers = {" & listTickers & "}," & vbCrLf & _
        "    GetData = (ticker as text) =>" & vbCrLf & _
        "        let" & vbCrLf & _
        "            Source = Csv.Document(Web.Contents(""https://query1.finance.yahoo.com/v7/finance/download/"" & ticker & ""?period1=1577836800&period2=2000000000&interval=1d&events=history""),[Delimiter="","", Columns=7, Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
        "            Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & vbCrLf & _
        "            FilteredNulls = Table.SelectRows(Promoted, each ([Close] <> ""null""))," & vbCrLf & _
        "            AddTicker = Table.AddColumn(FilteredNulls, ""Ticker"", each ticker)," & vbCrLf & _
        "            SelectCols = Table.SelectColumns(AddTicker, {""Date"", ""Ticker"", ""Close""})" & vbCrLf & _
        "        in" & vbCrLf & _
        "            SelectCols," & vbCrLf & _
        "    Combined = Table.Combine(List.Transform(Tickers, each GetData(_)))," & vbCrLf & _
        "    ChangedType = Table.TransformColumnTypes(Combined,{{""Date"", type date}, {""Close"", Currency.Type}, {""Ticker"", type text}})" & vbCrLf & _
        "in" & vbCrLf & _
        "    ChangedType"
        
    Injecter_PowerQuery "FACT_MarketQuotes", "Historique massifs des Prix de Clôture", M_Code_Quotes
    
    ' 3. Création ou Mise à Jour de la Requête Power Query (FACT_FX_Rates)
    Dim M_Code_FX As String
    If listFX = "" Then
        ' Si tous les actifs sont dans la devise de base, on crée une table vide pour l'intégrité du modèle DAX
        M_Code_FX = "let Source = #table({""Date"", ""Devise"", ""Taux""}, {}) in Source"
    Else
        M_Code_FX = "let" & vbCrLf & _
            "    PairesFX = {" & listFX & "}," & vbCrLf & _
            "    GetData = (paire as text) =>" & vbCrLf & _
            "        let" & vbCrLf & _
            "            Source = Csv.Document(Web.Contents(""https://query1.finance.yahoo.com/v7/finance/download/"" & paire & ""?period1=1577836800&period2=2000000000&interval=1d&events=history""),[Delimiter="","", Columns=7, Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCrLf & _
            "            Promoted = Table.PromoteHeaders(Source,[PromoteAllScalars=true])," & vbCrLf & _
            "            FilteredNulls = Table.SelectRows(Promoted, each ([Close] <> ""null""))," & vbCrLf & _
            "            AddDevise = Table.AddColumn(FilteredNulls, ""Devise"", each Text.Start(paire, 3))," & vbCrLf & _
            "            SelectCols = Table.SelectColumns(AddDevise, {""Date"", ""Devise"", ""Close""})" & vbCrLf & _
            "        in" & vbCrLf & _
            "            SelectCols," & vbCrLf & _
            "    Combined = Table.Combine(List.Transform(PairesFX, each GetData(_)))," & vbCrLf & _
            "    ChangedType = Table.TransformColumnTypes(Combined,{{""Date"", type date}, {""Devise"", type text}, {""Close"", Currency.Type}})" & vbCrLf & _
            "in" & vbCrLf & _
            "    ChangedType"
    End If
    
    Injecter_PowerQuery "FACT_FX_Rates", "Historique Croisé des Devises (Pour valorisation DAX)", M_Code_FX
    
    Application.ScreenUpdating = True
    MsgBox "L'ETL BOURSIER (BIG DATA) EST DÉPLOYÉ AVEC SUCCÈS !" & vbCrLf & vbCrLf & _
           "Power Query va désormais aspirer quotidiennement les historiques depuis 2020 via Yahoo Finance." & vbCrLf & _
           "La précision financière stricte (Currency.Type) est garantie. Zéro Float.", vbInformation, "WMS v1.0 - Étape 4"
End Sub

' --- Moteur d'Injection Silencieux Power Query (Connection Only pour le Data Model) ---
Private Sub Injecter_PowerQuery(NomRequete As String, Description As String, Code_M As String)
    Dim qry As WorkbookQuery
    Dim existe As Boolean: existe = False
    
    ' Vérifie si la requête existe et la met à jour
    For Each qry In ThisWorkbook.Queries
        If qry.Name = NomRequete Then
            qry.Formula = Code_M
            existe = True
            Exit For
        End If
    Next qry
    
    ' Création de la requête et de sa connexion OLAP (Power Pivot) si inexistante
    If Not existe Then
        Set qry = ThisWorkbook.Queries.Add(Name:=NomRequete, Formula:=Code_M, Description:=Description)
        ' CRÉATION SILENCIEUSE DANS LE DATA MODEL (CONNECTION ONLY)
        Dim connStr As String
        connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & NomRequete & ";Extended Properties="""""
        ThisWorkbook.Connections.Add2 Name:="Connection " & NomRequete, Description:="Connexion Power Pivot pour " & NomRequete, _
            ConnectionString:=connStr, CommandText:=NomRequete, lCmdtype:=6, CreateModelConnection:=True, ImportRelationships:=False
    End If
End Sub

' (Patch Utilitaire manquant de l'étape 1 pour que le système lise les Paramètres)
Public Function Obtenir_Parametre(NomParam As String, ValeurDefaut As String) As String
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    Dim tblConf As ListObject, i As Long
    On Error Resume Next: Set tblConf = wsSys.ListObjects("T_SYS_Config"): On Error GoTo 0
    If tblConf Is Nothing Then Obtenir_Parametre = ValeurDefaut: Exit Function
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = NomParam Then
            Obtenir_Parametre = tblConf.DataBodyRange(i, 2).Value
            Exit Function
        End If
    Next i
    wsSys.Unprotect "WMS_ADMIN_2026"
    Dim nr As ListRow: Set nr = tblConf.ListRows.Add
    nr.Range(1, 1).Value = NomParam: nr.Range(1, 2).Value = ValeurDefaut: nr.Range(1, 3).Value = "Auto-Created"
    wsSys.Protect "WMS_ADMIN_2026", UserInterfaceOnly:=True
    Obtenir_Parametre = ValeurDefaut
End Function

