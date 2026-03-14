Attribute VB_Name = "MOD_02_AppHome_Global"
Option Explicit

' =========================================================================
' MODULE: MOD_02_AppHome_Global
' OBJECTIF: Hub SPA Premium, ADLaM Display 10, Zoom 100%, Solid Cards, Zéro Régression
' =========================================================================

Public Sub DEPLOIEMENT_ETAPE_3_GLOBAL()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    
    ' 1. Préparation du Dictionnaire (Lexique du Hub)
    Preparer_Dictionnaire_Global
    
    ' 2. Construction de l'Interface Interactive Premium
    Preparer_Hub_Central
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
    
    MsgBox "LE HUB CENTRAL 'MASTER CLASS' EST DÉPLOYÉ." & vbCrLf & vbCrLf & _
           "1. L'esthétique 'Solid Cards' (Ombres Portées) est appliquée aux menus." & vbCrLf & _
           "2. Police ADLaM Display (Taille 10) et Zoom 100% verrouillés." & vbCrLf & _
           "3. Zéro régression sur le moteur U.C.R : l'interactivité est intacte.", vbInformation, "SFP v1.0 - Hub Premium"
End Sub

' -------------------------------------------------------------------------
' 1. MOTEUR DE DICTIONNAIRE (8 LANGUES LATINES - Anti '???')
' -------------------------------------------------------------------------
Private Sub Preparer_Dictionnaire_Global()
    Dim wsSys As Worksheet, tblDic As ListObject, tblConf As ListObject
    On Error Resume Next: Set wsSys = ThisWorkbook.Sheets("SYS_Config"): On Error GoTo 0
    Set tblConf = wsSys.ListObjects("T_SYS_Config")
    
    Dim langExist As Boolean: langExist = False
    Dim i As Long: For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = "LANGUE_UI" Then langExist = True: Exit For
    Next i
    If Not langExist Then
        Dim nrConf As ListRow: Set nrConf = tblConf.ListRows.Add
        nrConf.Range(1, 1).Value = "LANGUE_UI": nrConf.Range(1, 2).Value = "FR": nrConf.Range(1, 3).Value = "Langue UI Globale"
    End If

    On Error Resume Next: Set tblDic = wsSys.ListObjects("T_SYS_Dictionary"): On Error GoTo 0
    If tblDic Is Nothing Then
        wsSys.Columns("E:M").Clear
        wsSys.Range("E1:M1").Value = Array("KEY", "FR", "EN", "ES", "PT", "DE", "IT", "NL", "SV")
        Set tblDic = wsSys.ListObjects.Add(xlSrcRange, wsSys.Range("E1:M2"), , xlYes)
        tblDic.Name = "T_SYS_Dictionary"
        tblDic.TableStyle = "TableStyleMedium15"
        tblDic.ListRows(1).Delete
    End If
    
    Upsert_Trad tblDic, "APP_TITLE", "SYSTÈME FINANCIER PERSONNEL", "PERSONAL FINANCE SYSTEM", "SISTEMA FINANCIERO PERSONAL", "SISTEMA FINANCEIRO PESSOAL", "PERSÖNLICHES FINANZSYSTEM", "SISTEMA FINANZIARIO", "FINANCIEEL SYSTEEM", "FINANSSYSTEM"
    Upsert_Trad tblDic, "HUB_LOC", "Hub Central", "Central Hub", "Centro Principal", "Hub Central", "Zentraler Hub", "Hub Centrale", "Centrale Hub", "Central Hub"
    Upsert_Trad tblDic, "SAISIE_T", "SAISIE TRANSACTION", "ENTER TRANSACTION", "INGRESAR TRANSACCIÓN", "INSERIR TRANSAÇÃO", "TRANSAKTION ERFASSEN", "INSERISCI TRANSAZIONE", "TRANSACTIE INVOEREN", "ANGE TRANSAKTION"
    Upsert_Trad tblDic, "SAISIE_D", "Ajouter un revenu, dépense ou virement.", "Add an income, expense, or transfer.", "Añadir ingreso, gasto o transferencia.", "Adicionar renda, despesa ou transferência.", "Einkommen, Ausgabe oder Transfer.", "Aggiungi entrata, uscita o bonifico.", "Voeg inkomsten, uitgaven of overboeking toe.", "Lägg till inkomst, utgift eller överföring."
    Upsert_Trad tblDic, "DASH_T", "DASHBOARD CASHFLOW", "CASHFLOW DASHBOARD", "PANEL DE FLUJO DE CAJA", "PAINEL DE FLUXO DE CAIXA", "CASHFLOW-DASHBOARD", "DASHBOARD FLUSSI", "CASHFLOW DASHBOARD", "CASHFLOW DASHBOARD"
    Upsert_Trad tblDic, "DASH_D", "Analyser les flux consolidés.", "Analyze consolidated flows.", "Analizar flujos consolidados.", "Analisar fluxos consolidados.", "Konsolidierte Flüsse analysieren.", "Analizza i flussi consolidati.", "Analyseer de stromen.", "Analysera flöden."
    Upsert_Trad tblDic, "BUDG_T", "PILOTAGE BUDGÉTAIRE", "BUDGET TRACKING", "CONTROL PRESUPUESTARIO", "CONTROLE ORÇAMENTÁRIO", "BUDGETKONTROLLE", "CONTROLLO BUDGET", "BUDGETBEHEER", "BUDGETKONTROLL"
    Upsert_Trad tblDic, "BUDG_D", "Suivi des enveloppes (ZBB).", "Track budget envelopes.", "Seguimiento de presupuestos.", "Acompanhamento de orçamentos.", "Verfolgung der Budgets.", "Traccia i budget.", "Volg uw budgetten.", "Spåra dina budgetar."
    Upsert_Trad tblDic, "NETW_T", "BILAN PATRIMONIAL", "NET WORTH STATEMENT", "BALANCE PATRIMONIAL", "BALANÇO PATRIMONIAL", "VERMÖGENSBILANZ", "BILANCIO PATRIMONIALE", "VERMOGENSOVERZICHT", "FÖRMÖGENHETSRAPPORT"
    Upsert_Trad tblDic, "NETW_D", "Calcul de la Valeur Nette.", "Net Worth calculation.", "Cálculo del patrimonio neto.", "Cálculo do patrimônio líquido.", "Berechnung des Nettovermögens.", "Calcolo del patrimonio netto.", "Berekening nettowaarde.", "Beräkning av nettovärde."
    Upsert_Trad tblDic, "WELCOME", "Sélectionnez un module d'application ci-dessous pour démarrer.", "Select an application module below to get started.", "Seleccione un módulo de aplicación a continuación.", "Selecione um módulo de aplicativo abaixo.", "Wählen Sie unten ein Anwendungsmodul aus.", "Seleziona un modulo dell'applicazione di seguito.", "Selecteer hieronder een applicatiemodule.", "Välj en applikationsmodul nedan."
    Upsert_Trad tblDic, "TT_LANG", "Changer la langue : ", "Change language : ", "Cambiar idioma : ", "Mudar idioma : ", "Sprache ändern : ", "Cambia lingua : ", "Taal wijzigen : ", "Ändra språk : "
    ' --- DEBUT PATCH 1 (Lexique Bouton Sync FX) ---
    Upsert_Trad tblDic, "BTN_SYNC", "ACTUALISER TAUX FX", "UPDATE FX RATES", "ACTUALIZAR TIPOS FX", "ATUALIZAR TAXAS FX", "FX-KURSE AKTUALISIEREN", "AGGIORNA TASSI FX", "FX-TARIEVEN BIJWERKEN", "UPPDATERA FX-KURSER"
    ' --- FIN PATCH 1 ---
    ' --- DEBUT PATCH 4 (Lexique des Données d'Usine) ---
    Upsert_Trad tblDic, "Salaire / Revenus Pro", "Salaire / Revenus Pro", "Salary / Pro Income", "Salario / Ingresos Pro", "Salário / Renda Pro", "Gehalt / Einkommen", "Stipendio / Reddito", "Salaris / Inkomen", "Lön / Inkomst"
    Upsert_Trad tblDic, "Intérêts / Dividendes", "Intérêts / Dividendes", "Interests / Dividends", "Intereses / Dividendos", "Juros / Dividendos", "Zinsen / Dividenden", "Interessi / Dividendi", "Rente / Dividenden", "Ränta / Utdelning"
    Upsert_Trad tblDic, "Logement (Loyer/Prêt/Charges)", "Logement (Loyer/Prêt)", "Housing (Rent/Loan)", "Vivienda (Alquiler)", "Moradia (Aluguel)", "Wohnen (Miete/Kredit)", "Abitazione (Affitto)", "Huisvesting", "Bostad (Hyra/Lån)"
    Upsert_Trad tblDic, "Alimentation & Supermarché", "Alimentation", "Groceries & Food", "Alimentación", "Alimentação", "Lebensmittel", "Alimentari", "Boodschappen", "Mat & Livsmedel"
    Upsert_Trad tblDic, "Transports (Essence/Assurance)", "Transports", "Transport (Gas/Ins.)", "Transporte", "Transporte", "Transport", "Trasporti", "Vervoer", "Transport"
    Upsert_Trad tblDic, "Santé & Mutuelle", "Santé & Mutuelle", "Health & Insurance", "Salud y Seguros", "Saúde e Seguros", "Gesundheit", "Salute e Assicurazione", "Gezondheid", "Hälsa & Försäkring"
    Upsert_Trad tblDic, "Loisirs, Sorties & Vacances", "Loisirs & Vacances", "Leisure & Holidays", "Ocio y Vacaciones", "Lazer e Férias", "Freizeit & Urlaub", "Tempo Libero", "Vrije Tijd", "Fritid & Semester"
    Upsert_Trad tblDic, "Virement Interne (Épargne)", "Virement Interne", "Internal Transfer", "Transferencia Interna", "Transferência Interna", "Interner Transfer", "Bonifico Interno", "Interne Overboeking", "Intern Överföring"
    Upsert_Trad tblDic, "Compte Courant Principal", "Compte Courant", "Checking Account", "Cuenta Corriente", "Conta Corrente", "Girokonto", "Conto Corrente", "Betaalrekening", "Lönekonto"
    Upsert_Trad tblDic, "Livret d'Épargne", "Livret d'Épargne", "Savings Account", "Cuenta de Ahorros", "Conta Poupança", "Sparkonto", "Conto Risparmio", "Spaarrekening", "Sparkonto"
    Upsert_Trad tblDic, "Carte de Crédit (Différé)", "Carte de Crédit", "Credit Card", "Tarjeta de Crédito", "Cartão de Crédito", "Kreditkarte", "Carta di Credito", "Creditcard", "Kreditkort"
    ' --- FIN PATCH 4 ---
    ' --- DEBUT PATCH 1 (Lexique UI Paramètres) ---
    Upsert_Trad tblDic, "BTN_SETTING", "PARAMÈTRES (DEVISE)", "SETTINGS (CURRENCY)", "AJUSTES (DIVISA)", "CONFIG. (MOEDA)", "EINSTELLUNGEN", "IMPOSTAZIONI", "INSTELLINGEN", "INSTÄLLNINGAR"
    Upsert_Trad tblDic, "MSG_ASK_BASE", "Saisissez votre Devise Principale (ex: XOF, EUR, USD, CAD) :", "Enter your Base Currency (e.g. USD, EUR) :", "Ingrese su divisa base :", "Insira sua moeda base :", "Basiswährung eingeben :", "Inserisci valuta base :", "Voer basisvaluta in :", "Ange basvaluta :"
    ' --- FIN PATCH 1 ---
    ' --- DEBUT PATCH 1 (Lexique Proxy des Master Data & Types) ---
    Upsert_Trad tblDic, "Salaire / Revenus Pro", "Salaire / Revenus Pro", "Salary / Pro Income", "Salario / Ingresos", "Salário / Renda", "Gehalt / Einkommen", "Stipendio / Reddito", "Salaris / Inkomen", "Lön / Inkomst"
    Upsert_Trad tblDic, "Intérêts / Dividendes", "Intérêts / Dividendes", "Interests / Dividends", "Intereses / Dividendos", "Juros / Dividendos", "Zinsen / Dividenden", "Interessi / Dividendi", "Rente / Dividenden", "Ränta / Utdelning"
    Upsert_Trad tblDic, "Logement (Loyer/Prêt/Charges)", "Logement", "Housing (Rent/Loan)", "Vivienda", "Moradia", "Wohnen", "Abitazione", "Huisvesting", "Bostad"
    Upsert_Trad tblDic, "Alimentation & Supermarché", "Alimentation", "Groceries", "Alimentación", "Alimentação", "Lebensmittel", "Alimentari", "Boodschappen", "Mat"
    Upsert_Trad tblDic, "Transports (Essence/Assurance)", "Transports", "Transport", "Transporte", "Transporte", "Transport", "Trasporti", "Vervoer", "Transport"
    Upsert_Trad tblDic, "Santé & Mutuelle", "Santé & Mutuelle", "Health & Insurance", "Salud", "Saúde", "Gesundheit", "Salute", "Gezondheid", "Hälsa"
    Upsert_Trad tblDic, "Loisirs, Sorties & Vacances", "Loisirs & Vacances", "Leisure & Holidays", "Ocio", "Lazer", "Freizeit", "Tempo Libero", "Vrije Tijd", "Fritid"
    Upsert_Trad tblDic, "Virement Interne (Épargne)", "Virement Interne", "Internal Transfer", "Transferencia Interna", "Transferência Interna", "Interner Transfer", "Bonifico Interno", "Interne Overboeking", "Intern Överföring"
    Upsert_Trad tblDic, "Compte Courant Principal", "Compte Courant", "Checking Account", "Cuenta Corriente", "Conta Corrente", "Girokonto", "Conto Corrente", "Betaalrekening", "Lönekonto"
    Upsert_Trad tblDic, "Livret d'Épargne", "Livret d'Épargne", "Savings Account", "Cuenta de Ahorros", "Conta Poupança", "Sparkonto", "Conto Risparmio", "Spaarrekening", "Sparkonto"
    Upsert_Trad tblDic, "Carte de Crédit (Différé)", "Carte de Crédit", "Credit Card", "Tarjeta de Crédito", "Cartão de Crédito", "Kreditkarte", "Carta di Credito", "Creditcard", "Kreditkort"
    
    ' Traduction des balises Système Backend
    Upsert_Trad tblDic, "LIQUIDITE", "LIQUIDITÉ", "LIQUIDITY", "LIQUIDEZ", "LIQUIDEZ", "LIQUIDITÄT", "LIQUIDITÀ", "LIQUIDITEIT", "LIKVIDITET"
    Upsert_Trad tblDic, "INVESTISSEMENT", "INVESTISSEMENT", "INVESTMENT", "INVERSIÓN", "INVESTIMENTO", "INVESTITION", "INVESTIMENTO", "INVESTERING", "INVESTERING"
    Upsert_Trad tblDic, "DETTE", "DETTE", "DEBT", "DEUDA", "DÍVIDA", "SCHULD", "DEBITO", "SCHULD", "SKULD"
    Upsert_Trad tblDic, "DEPENSE", "DÉPENSE", "EXPENSE", "GASTO", "DESPESA", "AUSGABE", "USCITA", "UITGAVE", "UTGIFT"
    Upsert_Trad tblDic, "REVENU", "REVENU", "INCOME", "INGRESO", "RENDA", "EINKOMMEN", "ENTRATA", "INKOMSTEN", "INKOMST"
    Upsert_Trad tblDic, "TRANSFERT", "TRANSFERT", "TRANSFER", "TRANSFERENCIA", "TRANSFERÊNCIA", "TRANSFER", "TRASFERIMENTO", "OVERDRACHT", "ÖVERFÖRING"
    Upsert_Trad tblDic, "AUTRE", "AUTRE", "OTHER", "OTRO", "OUTRO", "ANDERE", "ALTRO", "ANDERS", "ANNAT"
    ' --- FIN PATCH 1 ---
    ' --- DEBUT PATCH 1 (Lexique Tiers et Autre) ---
    Upsert_Trad tblDic, "Autre (Préciser...)", "Autre (Préciser...)", "Other (Specify...)", "Otro (Especificar...)", "Outro (Especificar...)", "Andere (Angeben...)", "Altro (Specificare...)", "Anders (Specificeren...)", "Annat (Ange...)"
    Upsert_Trad tblDic, "Employeur principal", "Employeur", "Main Employer", "Empleador", "Empregador", "Arbeitgeber", "Datore di Lavoro", "Werkgever", "Arbetsgivare"
    Upsert_Trad tblDic, "Banque / Courtier", "Banque / Courtier", "Bank / Broker", "Banco / Bróker", "Banco / Corretor", "Bank / Makler", "Banca / Broker", "Bank / Makelaar", "Bank / Mäklare"
    Upsert_Trad tblDic, "État / Impôts", "État / Impôts", "State / Taxes", "Estado / Impuestos", "Estado / Impostos", "Staat / Steuern", "Stato / Tasse", "Staat / Belastingen", "Stat / Skatt"
    Upsert_Trad tblDic, "Supermarché / Alimentaire", "Supermarché", "Supermarket", "Supermercado", "Supermercado", "Supermarkt", "Supermercato", "Supermarkt", "Stormarknad"
    Upsert_Trad tblDic, "Propriétaire / Syndic", "Propriétaire", "Landlord / HOA", "Propietario", "Senhorio", "Vermieter", "Proprietario", "Verhuurder", "Hyresvärd"
    Upsert_Trad tblDic, "Station Service", "Station Service", "Gas Station", "Gasolinera", "Posto de Gasolina", "Tankstelle", "Stazione di Servizio", "Benzinestation", "Bensinstation"
    ' --- FIN PATCH 1 ---
    ' --- DEBUT PATCH 1 (Lexique de l'API) ---
    Upsert_Trad tblDic, "MSG_FX_OK", "TAUX DE CHANGE SYNCHRONISÉS." & vbCrLf & "Mise à jour réussie.", "EXCHANGE RATES SYNCHRONIZED." & vbCrLf & "Update successful.", "TIPOS DE CAMBIO SINCRONIZADOS." & vbCrLf & "Actualización exitosa.", "TAXAS DE CÂMBIO SINCRONIZADAS." & vbCrLf & "Atualização bem-sucedida.", "WECHSELKURSE SYNCHRONISIERT." & vbCrLf & "Erfolgreich aktualisiert.", "TASSI DI CAMBIO SINCRONIZZATI." & vbCrLf & "Aggiornamento riuscito.", "WISSELKOERSEN GESYNCHRONISEERD." & vbCrLf & "Succesvol bijgewerkt.", "VÄXELKURSER SYNKRONISERADE." & vbCrLf & "Uppdatering lyckades."
    Upsert_Trad tblDic, "MSG_FX_ERR", "Erreur de connexion au serveur de devises.", "Connection error to currency server.", "Error de conexión al servidor de divisas.", "Erro de conexão ao servidor de moedas.", "Verbindungsfehler zum Währungsserver.", "Errore di connessione al server delle valute.", "Verbindingsfout met valutaserver.", "Anslutningsfel till valutaserver."
    Upsert_Trad tblDic, "MSG_FX_TITLE", "Synchronisation FX", "FX Synchronization", "Sincronización FX", "Sincronização FX", "FX-Synchronisation", "Sincronizzazione FX", "FX-synchronisatie", "FX-synkronisering"
    ' --- FIN PATCH 1 ---
    ' --- DEBUT PATCH 1 (Lexique Erreur API Devise) ---
    Upsert_Trad tblDic, "MSG_ERR_DEV_API", "Devise invalide ou non reconnue par le marché boursier.", "Invalid currency or unrecognized by the market.", "Divisa inválida o no reconocida.", "Moeda inválida ou não reconhecida.", "Ungültige oder nicht erkannte Währung.", "Valuta non valida o non riconosciuta.", "Ongeldige valuta of niet herkend.", "Ogiltig valuta eller okänd."
    ' --- FIN PATCH 1 ---
    ' --- DEBUT PATCH (Traduction exhaustive des Master Data) ---
    Upsert_Trad tblDic, "Prêt Immobilier", "Prêt Immobilier", "Mortgage Loan", "Préstamo Hipotecario", "Empréstimo Imobiliário", "Immobilienkredit", "Mutuo Immobiliare", "Hypothecaire Lening", "Bolån"
    Upsert_Trad tblDic, "Impôts & Taxes", "Impôts & Taxes", "Taxes", "Impuestos", "Impostos", "Steuern", "Tasse", "Belastingen", "Skatter"
    Upsert_Trad tblDic, "Compte Courant Principal", "Compte Courant", "Checking Account", "Cuenta Corriente", "Conta Corrente", "Girokonto", "Conto Corrente", "Betaalrekening", "Lönekonto"
    Upsert_Trad tblDic, "Livret d'Épargne", "Livret d'Épargne", "Savings Account", "Cuenta de Ahorros", "Conta Poupança", "Sparkonto", "Conto Risparmio", "Spaarrekening", "Sparkonto"
    Upsert_Trad tblDic, "Portefeuille Espèces", "Portefeuille Espèces", "Cash Wallet", "Cartera de Efectivo", "Carteira de Dinheiro", "Bargeldbörse", "Portafoglio Contanti", "Portemonnee", "Kontanter"
    Upsert_Trad tblDic, "Assurance Vie", "Assurance Vie", "Life Insurance", "Seguro de Vida", "Seguro de Vida", "Lebensversicherung", "Assicurazione Vita", "Levensverzekering", "Livförsäkring"
    Upsert_Trad tblDic, "PEA / Actions", "PEA / Actions", "Stock Portfolio", "Cartera de Acciones", "Carteira de Ações", "Aktienportfolio", "Portafoglio Azioni", "Aandelenportefeuille", "Aktieportfölj"
    Upsert_Trad tblDic, "Portefeuille Crypto", "Portefeuille Crypto", "Crypto Wallet", "Cartera Crypto", "Carteira Crypto", "Krypto-Wallet", "Portafoglio Crypto", "Cryptoportemonnee", "Kryptoplånbok"
    Upsert_Trad tblDic, "Carte de Crédit (Différé)", "Carte de Crédit", "Credit Card", "Tarjeta de Crédito", "Cartão de Crédito", "Kreditkarte", "Carta di Credito", "Creditcard", "Kreditkort"
    ' --- FIN PATCH ---
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
        Case Else: ColIdx = 2
    End Select
    
    For i = 1 To tblDic.ListRows.Count
        If tblDic.DataBodyRange(i, 1).Value = Clé Then TR = tblDic.DataBodyRange(i, ColIdx).Value: Exit Function
    Next i
    TR = Clé
End Function

' -------------------------------------------------------------------------
' 2. CONSTRUCTION DU HUB (DESIGN PREMIUM & U.C.R INTACT)
' -------------------------------------------------------------------------
Private Sub Preparer_Hub_Central()
    Dim wsHome As Worksheet
    On Error Resume Next: Set wsHome = ThisWorkbook.Sheets("APP_HOME"): On Error GoTo 0
    
    If wsHome Is Nothing Then
        Set wsHome = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsHome.Name = "APP_HOME"
    Else
        wsHome.Cells.Clear
        Dim shp As Shape: For Each shp In wsHome.Shapes: shp.Delete: Next shp
        wsHome.Hyperlinks.Delete
    End If
    
    ' --- FORÇAGE DU ZOOM ET DE LA POLICE GLOBALE ---
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    wsHome.Cells.Font.Name = "ADLaM Display"
    wsHome.Cells.Font.Size = 10
    wsHome.Cells.Interior.Color = RGB(248, 248, 250)
    
    ' --- BANDEAU SUPÉRIEUR (Plus majestueux) ---
    wsHome.Range("A1:Z5").Interior.Color = RGB(65, 105, 225) ' Bleu Royal
    
    ' --- TITRE VECTORIEL ---
    Dim shpTitle As Shape
    Set shpTitle = wsHome.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 15, 500, 50)
    shpTitle.Fill.Visible = msoFalse: shpTitle.Line.Visible = msoFalse
    shpTitle.TextFrame2.TextRange.Text = UCase(TR("APP_TITLE")) & vbCrLf & TR("HUB_LOC") & " | " & Format(Date, "dd mmmm yyyy")
    shpTitle.TextFrame2.TextRange.Lines(1).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(1).Font.Size = 22: shpTitle.TextFrame2.TextRange.Lines(1).Font.Bold = True: shpTitle.TextFrame2.TextRange.Lines(1).Font.Fill.ForeColor.RGB = vbWhite
    shpTitle.TextFrame2.TextRange.Lines(2).Font.Name = "ADLaM Display": shpTitle.TextFrame2.TextRange.Lines(2).Font.Size = 11: shpTitle.TextFrame2.TextRange.Lines(2).Font.Fill.ForeColor.RGB = RGB(220, 220, 255)
    
    ' --- LA SIGNATURE "SFP v1.0" TOP-RIGHT ---
    Dim lblVersion As Shape
    Set lblVersion = wsHome.Shapes.AddTextbox(msoTextOrientationHorizontal, 900, 5, 220, 20)
    lblVersion.Fill.Visible = msoFalse: lblVersion.Line.Visible = msoFalse
    lblVersion.TextFrame2.TextRange.Text = "SFP v1.0"
    lblVersion.TextFrame2.TextRange.Font.Name = "ADLaM Display": lblVersion.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite: lblVersion.TextFrame2.TextRange.Font.Bold = True: lblVersion.TextFrame2.TextRange.Font.Size = 9
    lblVersion.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
    
    ' --- LES 8 LANGUES (Cercles parfaits) ---
    ' --- DEBUT PATCH 2 (Bouton Sync FX Tactile) ---
    ' --- DEBUT PATCH 2 (Bouton Paramètres MVP) ---
    Dim btnSet As Shape
    Set btnSet = wsHome.Shapes.AddShape(msoShapeRoundedRectangle, 440, 30, 150, 35)
    btnSet.Fill.ForeColor.RGB = RGB(128, 128, 128) ' Gris Pro Neutre
    btnSet.Line.Visible = msoFalse
    btnSet.TextFrame2.TextRange.Text = TR("BTN_SETTING")
    btnSet.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnSet.TextFrame2.TextRange.Font.Bold = True: btnSet.TextFrame2.TextRange.Font.Size = 9: btnSet.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
    btnSet.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnSet.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With btnSet.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 4: .OffsetX = 0: .OffsetY = 2: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    wsHome.Hyperlinks.Add Anchor:=btnSet, Address:="", SubAddress:="'" & wsHome.Name & "'!A26", ScreenTip:=TR("BTN_SETTING")
    ' --- FIN PATCH 2 ---
    Dim btnSync As Shape
    Set btnSync = wsHome.Shapes.AddShape(msoShapeRoundedRectangle, 600, 30, 150, 35)
    btnSync.Fill.ForeColor.RGB = RGB(128, 128, 128) 'RGB(250, 218, 94) JAUNE 'RGB(46, 204, 113) ' Vert Émeraude
    btnSync.Line.Visible = msoFalse
    btnSync.TextFrame2.TextRange.Text = TR("BTN_SYNC")
    btnSync.TextFrame2.TextRange.Font.Name = "ADLaM Display": btnSync.TextFrame2.TextRange.Font.Bold = True: btnSync.TextFrame2.TextRange.Font.Size = 9: btnSync.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite 'RGB(0, 0, 0)
    btnSync.TextFrame2.VerticalAnchor = msoAnchorMiddle: btnSync.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    With btnSync.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow: .Blur = 4: .OffsetX = 0: .OffsetY = 2: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    ' Routage U.C.R vers la cellule A25
    wsHome.Hyperlinks.Add Anchor:=btnSync, Address:="", SubAddress:="'" & wsHome.Name & "'!A25", ScreenTip:=TR("BTN_SYNC")
    ' --- FIN PATCH 2 ---
    Dim arrLang As Variant: arrLang = Array("FR", "EN", "ES", "PT", "DE", "IT", "NL", "SV")
    Dim i As Integer, xPos As Integer: xPos = 770
    For i = LBound(arrLang) To UBound(arrLang)
        Dessiner_Bouton_Langue wsHome, CStr(arrLang(i)), xPos, 30, 35, 35, TR("TT_LANG") & arrLang(i), "A" & (11 + i)
        xPos = xPos + 40
    Next i
    
    ' --- MESSAGE D'ACCUEIL ---
    wsHome.Range("C8").Value = TR("WELCOME")
    wsHome.Range("C8").Font.Color = RGB(150, 150, 150): wsHome.Range("C8").Font.Italic = True
    
    ' --- LES TUILES DE NAVIGATION (SOLID CARDS PREMIUM) ---
    Dim T_Top As Integer: T_Top = 160
    Dim T_Left As Integer: T_Left = 100
    Dim T_W As Integer: T_W = 380
    Dim T_H As Integer: T_H = 110
    Dim Gap As Integer: Gap = 30
    
    ' 1. Saisie (Jaune Royal)
    Dessiner_Tuile_Premium wsHome, TR("SAISIE_T") & vbCrLf & TR("SAISIE_D"), T_Left, T_Top, T_W, T_H, RGB(250, 218, 94), RGB(40, 40, 40), TR("SAISIE_T"), "A21"
    
    ' 2. Dashboard Cashflow (Violet Royal)
    Dessiner_Tuile_Premium wsHome, TR("DASH_T") & vbCrLf & TR("DASH_D"), T_Left + T_W + Gap, T_Top, T_W, T_H, RGB(120, 81, 169), vbWhite, TR("DASH_T"), "A22"
    
    ' 3. Budget ZBB (Bleu Royal)
    Dessiner_Tuile_Premium wsHome, TR("BUDG_T") & vbCrLf & TR("BUDG_D"), T_Left, T_Top + T_H + Gap, T_W, T_H, RGB(65, 105, 225), vbWhite, TR("BUDG_T"), "A23"
    
    ' 4. Net Worth (Vert Émeraude)
    Dessiner_Tuile_Premium wsHome, TR("NETW_T") & vbCrLf & TR("NETW_D"), T_Left + T_W + Gap, T_Top + T_H + Gap, T_W, T_H, RGB(46, 204, 113), vbWhite, TR("NETW_T"), "A24"

    wsHome.Activate
    wsHome.Range("A1").Select
End Sub

Private Sub Dessiner_Bouton_Langue(ws As Worksheet, Texte As String, Gauche As Integer, Haut As Integer, Largeur As Integer, Hauteur As Integer, ToolTip As String, CelluleCible As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeOval, Gauche, Haut, Largeur, Hauteur)
    btn.Fill.ForeColor.RGB = RGB(40, 70, 180) ' Bleu Sombre
    btn.Line.ForeColor.RGB = vbWhite: btn.Line.Weight = 1.5
    btn.TextFrame2.WordWrap = msoFalse: btn.TextFrame2.MarginLeft = 0: btn.TextFrame2.MarginRight = 0: btn.TextFrame2.MarginTop = 0: btn.TextFrame2.MarginBottom = 0
    btn.TextFrame2.TextRange.Text = Texte
    btn.TextFrame2.TextRange.Font.Name = "ADLaM Display": btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite: btn.TextFrame2.TextRange.Font.Bold = True: btn.TextFrame2.TextRange.Font.Size = 10
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle: btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    ws.Hyperlinks.Add Anchor:=btn, Address:="", SubAddress:="'" & ws.Name & "'!" & CelluleCible, ScreenTip:=ToolTip
End Sub

Private Sub Dessiner_Tuile_Premium(ws As Worksheet, Texte As String, Gauche As Integer, Haut As Integer, Largeur As Integer, Hauteur As Integer, CoulFond As Long, CoulTexte As Long, ToolTip As String, CelluleCible As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, Gauche, Haut, Largeur, Hauteur)
    btn.Fill.ForeColor.RGB = CoulFond
    btn.Line.Visible = msoFalse
    
    ' Ombre Portée 3D
    With btn.Shadow
        .Type = msoShadow21: .Visible = msoTrue: .Style = msoShadowStyleOuterShadow
        .Blur = 8: .OffsetX = 0: .OffsetY = 4: .Transparency = 0.5: .ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    btn.TextFrame2.TextRange.Text = Texte
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CoulTexte
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    
    With btn.TextFrame2.TextRange.Lines(1).Font
        .Name = "ADLaM Display": .Bold = True: .Size = 16 ' Titre de la tuile très lisible
    End With
    With btn.TextFrame2.TextRange.Lines(2).Font
        .Name = "ADLaM Display": .Bold = False: .Size = 11 ' Description plus douce
    End With
    
    ' Le Moteur d'Interactivité U.C.R (Zéro Régression)
    ws.Hyperlinks.Add Anchor:=btn, Address:="", SubAddress:="'" & ws.Name & "'!" & CelluleCible, ScreenTip:=ToolTip
End Sub

' -------------------------------------------------------------------------
' 3. ACTIONS EXÉCUTABLES PAR LE CERVEAU (ThisWorkbook)
' -------------------------------------------------------------------------
'Public Sub EXECUTER_CHANGER_LANGUE(LangueCible As String)
    'Application.ScreenUpdating = False
    'Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config"): wsSys.Unprotect "SFP_ADMIN_2026"
    'Dim tblConf As ListObject: Set tblConf = wsSys.ListObjects("T_SYS_Config")
    'Dim i As Long: For i = 1 To tblConf.ListRows.Count
        'If tblConf.DataBodyRange(i, 1).Value = "LANGUE_UI" Then tblConf.DataBodyRange(i, 2).Value = LangueCible: Exit For
    'Next i
    'wsSys.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    
    'Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    'Preparer_Hub_Central ' Redessin in-place
    'For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    'Application.ScreenUpdating = True
'End Sub

' --- DEBUT PATCH (Réanimation U.C.R & Auto-Cicatrisation i18n) ---
Public Sub EXECUTER_CHANGER_LANGUE(LangueCible As String)
    ' 1. FORCE LE DÉBLOCAGE DU MOTEUR INTERACTIF (Anti-Freeze)
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    wsSys.Unprotect "SFP_ADMIN_2026"
    
    Dim tblConf As ListObject: Set tblConf = wsSys.ListObjects("T_SYS_Config")
    Dim i As Long, trouve As Boolean: trouve = False
    
    ' 2. LOGIQUE UPSERT (Met à jour si existe, Créé sinon)
    If tblConf.ListRows.Count > 0 Then
        For i = 1 To tblConf.ListRows.Count
            If tblConf.DataBodyRange(i, 1).Value = "LANGUE_UI" Then
                tblConf.DataBodyRange(i, 2).Value = LangueCible
                trouve = True
                Exit For
            End If
        Next i
    End If
    
    If Not trouve Then
        Dim nr As ListRow: Set nr = tblConf.ListRows.Add
        nr.Range(1, 1).Value = "LANGUE_UI"
        nr.Range(1, 2).Value = LangueCible
        nr.Range(1, 3).Value = "Langue UI Globale"
    End If
    
    wsSys.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    
    ' 3. REDESSIN DU HUB
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    Preparer_Hub_Central
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    
    Application.ScreenUpdating = True
End Sub
' --- FIN PATCH ---

Public Sub EXECUTER_ROUTER_SAISIE()
    On Error GoTo ErrForm
    USF_Transaction.Show
    Exit Sub
ErrForm:
    MsgBox "Le Formulaire est introuvable." & vbCrLf & "Veuillez relancer l'Étape 4.", vbCritical, "Gatekeeper Hors Ligne"
End Sub

Public Sub EXECUTER_ROUTER_DASHBOARD()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    MOD_04_Dashboard_ETL.GENERER_DASHBOARD
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

Public Sub EXECUTER_ROUTER_BUDGET()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    MOD_06_Budget_ZBB.GENERER_BUDGET_DASHBOARD
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

Public Sub EXECUTER_ROUTER_NETWORTH()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    MOD_05_Advanced_Modules.GENERER_NET_WORTH_DASHBOARD
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    Application.ScreenUpdating = True
End Sub

' --- DEBUT PATCH 2 (Validation API Préventive) ---
Public Sub EXECUTER_CONFIG_SYSTEME()
    Application.EnableEvents = True
    
    Dim wsSys As Worksheet: Set wsSys = ThisWorkbook.Sheets("SYS_Config")
    Dim tblConf As ListObject: Set tblConf = wsSys.ListObjects("T_SYS_Config")
    Dim currentBase As String: currentBase = "MUR"
    Dim i As Long
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = "SYS_DEVISE_BASE" Then currentBase = tblConf.DataBodyRange(i, 2).Value: Exit For
    Next i
    
    Dim rep As String
    rep = InputBox(TR("MSG_ASK_BASE"), TR("BTN_SETTING"), currentBase)
    If rep = "" Or UCase(Trim(rep)) = currentBase Then Exit Sub
    rep = UCase(Trim(rep))
    
    ' 1. SÉCURITÉ ABSOLUE : On interroge l'API pour valider l'existence réelle de la devise
    If Len(rep) <> 3 Then
        MsgBox TR("MSG_ERR_DEV_API"), vbCritical, TR("BTN_SETTING")
        Exit Sub
    End If
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://open.er-api.com/v6/latest/" & rep, False
    http.send
    If http.Status <> 200 Or InStr(1, http.responseText, """result"":""success""") = 0 Then
        MsgBox TR("MSG_ERR_DEV_API") & vbCrLf & "-> " & rep, vbCritical, TR("BTN_SETTING")
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    ' 2. Mise à jour de SYS_Config (Verrouillée)
    wsSys.Unprotect "SFP_ADMIN_2026"
    Dim found As Boolean: found = False
    For i = 1 To tblConf.ListRows.Count
        If tblConf.DataBodyRange(i, 1).Value = "SYS_DEVISE_BASE" Then
            tblConf.DataBodyRange(i, 2).Value = rep
            found = True: Exit For
        End If
    Next i
    If Not found Then
        Dim rC As ListRow: Set rC = tblConf.ListRows.Add
        rC.Range(1, 1).Value = "SYS_DEVISE_BASE": rC.Range(1, 2).Value = rep: rC.Range(1, 3).Value = "Devise Globale"
    End If
    
    ' 3. Swap Dynamique SANS Doublon dans T_SYS_Devises
    Dim tblDev As ListObject
    On Error Resume Next: Set tblDev = wsSys.ListObjects("T_SYS_Devises"): On Error GoTo 0
    If Not tblDev Is Nothing Then
        Dim rDel As Long
        For rDel = tblDev.ListRows.Count To 1 Step -1
            If UCase(Trim(CStr(tblDev.DataBodyRange(rDel, 1).Value))) = rep Then tblDev.ListRows(rDel).Delete
        Next rDel
        Dim devFound As Boolean: devFound = False
        Dim k As Long
        For k = 1 To tblDev.ListRows.Count
            If UCase(Trim(CStr(tblDev.DataBodyRange(k, 1).Value))) = currentBase Then
                tblDev.DataBodyRange(k, 1).Value = rep
                tblDev.DataBodyRange(k, 2).Value = 1
                devFound = True: Exit For
            End If
        Next k
        If Not devFound Then
            Dim nD As ListRow: Set nD = tblDev.ListRows.Add
            nD.Range(1, 1).Value = rep: nD.Range(1, 2).Value = 1
        End If
    End If
    wsSys.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    
    ' 4. Migration de la table DIM_Compte
    Dim wsCpt As Worksheet: Set wsCpt = ThisWorkbook.Sheets("DIM_Compte")
    wsCpt.Unprotect "SFP_ADMIN_2026"
    Dim tblCpt As ListObject
    On Error Resume Next: Set tblCpt = wsCpt.ListObjects("T_DIM_Compte"): On Error GoTo 0
    If Not tblCpt Is Nothing Then
        If tblCpt.ListRows.Count > 0 Then
            For i = 1 To tblCpt.ListRows.Count
                If UCase(Trim(CStr(tblCpt.DataBodyRange(i, 4).Value))) = currentBase Then tblCpt.DataBodyRange(i, 4).Value = rep
            Next i
        End If
    End If
    wsCpt.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True
    
    ' 5. Téléchargement instantané des nouveaux taux via l'API
    MOD_99_SystemAdmin.ACTUALISER_DEVISES_WEB
    
    Application.ScreenUpdating = True
End Sub
' --- FIN PATCH 2 ---
