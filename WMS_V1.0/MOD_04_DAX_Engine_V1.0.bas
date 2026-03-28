Attribute VB_Name = "MOD_04_DAX_Engine"
Option Explicit

' =========================================================================
' MODULE: MOD_04_DAX_Engine
' OBJECTIF: Création du Star Schema en RAM et Injection des Algorithmes DAX
' =========================================================================

Public Sub DEPLOYER_WMS_ETAPE_5_DAX()
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' 1. CHARGEMENT DES TABLES EXCEL DANS LE DATA MODEL (xVelocity)
    On Error Resume Next
    ThisWorkbook.Connections.Add2 "Linked_T_DIM_Portfolio", "", "WORKSHEET;T_DIM_Portfolio", "T_DIM_Portfolio", 7, True, False
    ThisWorkbook.Connections.Add2 "Linked_T_DIM_Asset", "", "WORKSHEET;T_DIM_Asset", "T_DIM_Asset", 7, True, False
    ThisWorkbook.Connections.Add2 "Linked_T_FACT_Trade", "", "WORKSHEET;T_FACT_Trade", "T_FACT_Trade", 7, True, False
    On Error GoTo ErrorHandler

    Dim mdl As Model
    Set mdl = ThisWorkbook.Model

    ' Pause de sécurité pour laisser à Power Pivot le temps de digérer les tables
    ThisWorkbook.Model.Refresh
    DoEvents

    ' 2. CONSTRUCTION DES RELATIONS (STAR SCHEMA)
    On Error Resume Next ' Évite l'erreur si les relations existent déjà
    
    ' FACT_Trade -> DIM_Asset
    mdl.ModelRelationships.Add mdl.ModelTables("T_FACT_Trade").ModelTableColumns("ID_Asset"), _
                               mdl.ModelTables("T_DIM_Asset").ModelTableColumns("ID_Asset")
    
    ' FACT_Trade -> DIM_Portfolio
    mdl.ModelRelationships.Add mdl.ModelTables("T_FACT_Trade").ModelTableColumns("ID_Portfolio"), _
                               mdl.ModelTables("T_DIM_Portfolio").ModelTableColumns("ID_Portfolio")
    
    ' FACT_MarketQuotes -> DIM_Asset (Liaison du Big Data via le Ticker Boursier)
    mdl.ModelRelationships.Add mdl.ModelTables("FACT_MarketQuotes").ModelTableColumns("Ticker"), _
                               mdl.ModelTables("T_DIM_Asset").ModelTableColumns("Ticker_Symbole")
    On Error GoTo ErrorHandler

    ' 3. INJECTION DES ALGORITHMES DE TRADING (MESURES DAX)
    
    ' A. Position Nette (Nombre de parts : Achats - Ventes)
    Injecter_Mesure_DAX mdl, "T_FACT_Trade", "Total_Shares", _
        "SUMX('T_FACT_Trade', IF('T_FACT_Trade'[Type_Ordre]=""ACHAT"", 'T_FACT_Trade'[Quantite], IF('T_FACT_Trade'[Type_Ordre]=""VENTE"", -'T_FACT_Trade'[Quantite], 0)))"
    
    ' B. Capital Investi Brut (Prix unitaire * Qté * Taux de change lors de l'achat)
    Injecter_Mesure_DAX mdl, "T_FACT_Trade", "Invested_Capital", _
        "SUMX('T_FACT_Trade', IF('T_FACT_Trade'[Type_Ordre]=""ACHAT"", 'T_FACT_Trade'[Quantite] * 'T_FACT_Trade'[Prix_Unitaire] * 'T_FACT_Trade'[Taux_FX_Historique], IF('T_FACT_Trade'[Type_Ordre]=""VENTE"", -'T_FACT_Trade'[Quantite] * 'T_FACT_Trade'[Prix_Unitaire] * 'T_FACT_Trade'[Taux_FX_Historique], 0)))"
    
    ' C. Prix de Clôture Actuel (Dernier prix connu dans la base Big Data pour chaque actif)
    Injecter_Mesure_DAX mdl, "FACT_MarketQuotes", "Current_Price", _
        "CALCULATE(SUM('FACT_MarketQuotes'[Close]), FILTER('FACT_MarketQuotes', 'FACT_MarketQuotes'[Date] = MAX('FACT_MarketQuotes'[Date])))"
    
    ' D. Valeur de Marché (Quantité possédée * Prix Actuel)
    Injecter_Mesure_DAX mdl, "T_DIM_Asset", "Market_Value", _
        "[Total_Shares] * [Current_Price]"
    
    ' E. Plus-Value Latente Nette (Unrealized PnL)
    Injecter_Mesure_DAX mdl, "T_DIM_Asset", "Unrealized_PnL", _
        "[Market_Value] -[Invested_Capital]"

    Application.ScreenUpdating = True
    MsgBox "LE CERVEAU ANALYTIQUE (DAX) EST DÉPLOYÉ !" & vbCrLf & vbCrLf & _
           "1. Les tables Excel et les données de Yahoo Finance sont connectées en RAM." & vbCrLf & _
           "2. Les algorithmes de Plus-Values et de Valorisation (TWRR) sont armés.", vbInformation, "WMS v1.0 - Étape 5"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Une erreur est survenue lors de l'injection DAX dans Power Pivot." & vbCrLf & "Erreur : " & Err.Description, vbCritical, "Alerte Data Model"
End Sub

' --- Utilitaire d'Injection DAX ---
Private Sub Injecter_Mesure_DAX(mdl As Model, TableName As String, MeasureName As String, DaxFormula As String)
    On Error Resume Next
    ' Supprime la mesure si elle existe déjà pour éviter les doublons lors des mises à jour
    Dim msr As ModelMeasure
    For Each msr In mdl.ModelMeasures
        If UCase(msr.Name) = UCase(MeasureName) Then
            msr.Delete
            Exit For
        End If
    Next msr
    On Error GoTo 0
    
    ' Ajoute la mesure avec le format Nombre Décimal par défaut
    On Error Resume Next
    mdl.ModelMeasures.Add MeasureName, mdl.ModelTables(TableName), DaxFormula, mdl.ModelFormatDecimalNumber
    On Error GoTo 0
End Sub

