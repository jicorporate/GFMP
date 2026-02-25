Attribute VB_Name = "MOD_03_Gatekeeper"
Option Explicit

' =========================================================================
' MODULE: MOD_03_Gatekeeper
' OBJECTIF: Formulaire Saisie, Explicit Auto-Learning, Typage Strict, Zéro Régression
' =========================================================================

Public Sub DEPLOIEMENT_ETAPE_4_GATEKEEPER()
    Application.ScreenUpdating = False
    
    Dim VBP As Object: On Error Resume Next: Set VBP = ThisWorkbook.VBProject: On Error GoTo 0
    If VBP Is Nothing Then MsgBox "Activez l'accès au modèle d'objet VBA.", vbCritical: Exit Sub
    
    Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: ws.Unprotect "SFP_ADMIN_2026": Next ws
    
    ' 1. Injection du vocabulaire et des Tooltips
    Garantir_Lexique_Formulaire
    
    ' 2. Destruction du formulaire obsolète
    Dim VBComp As Object
    For Each VBComp In VBP.VBComponents
        If VBComp.Name = "USF_Transaction" Then VBP.VBComponents.Remove VBComp
    Next VBComp
    
    ' 3. Génération du Designer UI (Zéro modification de la taille globale)
    Set VBComp = VBP.VBComponents.Add(3) ' vbext_ct_MSForm
    VBComp.Properties("Name") = "USF_Transaction"
    
    Dim myForm As Object: Set myForm = VBComp.Designer
    VBComp.Properties("Width") = 280
    VBComp.Properties("Height") = 500
    
    Dim t As Integer: t = 10
    Creer_Controle myForm, "txt_Date", "TextBox", t
    
    ' Le module Compte avec Typage Explicite
    Creer_Controle myForm, "cmb_Compte", "ComboBox", t
    Creer_Controle_Double myForm, "txt_New_Compte", "cmb_New_Cpt_Type", t
    
    ' Le module Catégorie avec Typage Explicite
    Creer_Controle myForm, "cmb_Categorie", "ComboBox", t
    Creer_Controle_Double myForm, "txt_New_Categorie", "cmb_New_Cat_Type", t
    
    ' Le module Tiers (Classique)
    Creer_Controle myForm, "cmb_Tiers", "ComboBox", t
    Creer_Controle myForm, "txt_New_Tiers", "TextBox", t, True
    
    Creer_Controle myForm, "txt_Montant", "TextBox", t
    Creer_Controle myForm, "cmb_Devise", "ComboBox", t
    Creer_Controle myForm, "txt_Description", "TextBox", t
    
    ' Boutons Charte Royale
    Dim ctrl As Object
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Save")
    ctrl.Top = t + 10: ctrl.Left = 30: ctrl.Width = 100: ctrl.Height = 25
    ctrl.BackColor = RGB(250, 218, 94): ctrl.Font.Bold = True
    
    Set ctrl = myForm.Controls.Add("Forms.CommandButton.1", "btn_Cancel")
    ctrl.Top = t + 10: ctrl.Left = 140: ctrl.Width = 100: ctrl.Height = 25
    ctrl.BackColor = RGB(120, 81, 169): ctrl.ForeColor = vbWhite: ctrl.Font.Bold = True
    
    ' 4. Injection du Cerveau Comptable
    VBComp.CodeModule.AddFromString Code_VBA_Formulaire()
    
    For Each ws In ThisWorkbook.Worksheets: ws.Protect "SFP_ADMIN_2026", UserInterfaceOnly:=True: Next ws
    AppActivate ThisWorkbook.Name
    Application.ScreenUpdating = True
    
    MsgBox "L'INTELLIGENCE COMPTABLE EST RÉTABLIE." & vbCrLf & vbCrLf & _
           "1. Lors de l'apprentissage (Autre...), vous devez désormais spécifier le type (Revenu/Dépense, Liquidité/Dette)." & vbCrLf & _
           "2. Zéro régression sur le design et les Tooltips." & vbCrLf & _
           "Vos actifs et passifs seront désormais calculés avec une précision mathématique absolue.", vbInformation, "SFP v3.2 - Correction Bilan"
End Sub

' -------------------------------------------------------------------------
' MOTEUR DE DESSIN UI (Avec l'Innovation "Double Contrôle")
' -------------------------------------------------------------------------
Private Sub Creer_Controle(myForm As Object, Nom As String, TypeCtrl As String, ByRef t As Integer, Optional EstCache As Boolean = False)
    Dim lbl As Object, c As Object
    Set lbl = myForm.Controls.Add("Forms.Label.1", "lbl_" & Nom)
    lbl.Top = t: lbl.Left = 20: lbl.Width = 220: lbl.Height = 12
    Set c = myForm.Controls.Add("Forms." & TypeCtrl & ".1", Nom)
    c.Top = t + 12: c.Left = 20: c.Width = 220: c.Height = 18
    If EstCache Then
        lbl.Visible = False: c.Visible = False: t = t + 35
    Else
        t = t + 35
    End If
End Sub

Private Sub Creer_Controle_Double(myForm As Object, NomTxt As String, NomCmb As String, ByRef t As Integer)
    Dim lbl As Object, cTxt As Object, cCmb As Object
    Set lbl = myForm.Controls.Add("Forms.Label.1", "lbl_" & NomTxt)
    lbl.Top = t: lbl.Left = 20: lbl.Width = 220: lbl.Height = 12: lbl.Visible = False
    
    ' Le champ texte est raccourci pour faire place au menu déroulant de Typage
    Set cTxt = myForm.Controls.Add("Forms.TextBox.1", NomTxt)
    cTxt.Top = t + 12: cTxt.Left = 20: cTxt.Width = 125: cTxt.Height = 18: cTxt.Visible = False
    
    Set cCmb = myForm.Controls.Add("Forms.ComboBox.1", NomCmb)
    cCmb.Top = t + 12: cCmb.Left = 150: cCmb.Width = 90: cCmb.Height = 18: cCmb.Visible = False
    
    t = t + 35
End Sub

' -------------------------------------------------------------------------
' MOTEUR AUTO-CICATRISANT DU FORMULAIRE
' -------------------------------------------------------------------------
Public Sub Garantir_Lexique_Formulaire()
    Dim tblDic As ListObject
    On Error Resume Next: Set tblDic = ThisWorkbook.Sheets("SYS_Config").ListObjects("T_SYS_Dictionary"): On Error GoTo 0
    If tblDic Is Nothing Then Exit Sub
    
    Upsert_Dico tblDic, "FRM_TITLE", "Saisie de Transaction", "Transaction Entry", "Ingreso de Transacción", "Registro de Transação", "Transaktion Erfassen", "Inserimento Transazione", "Transactie Invoer", "Transaktionsinmatning"
    Upsert_Dico tblDic, "FRM_DATE", "Date (JJ/MM/AAAA) :", "Date (DD/MM/YYYY) :", "Fecha (DD/MM/AAAA) :", "Data (DD/MM/AAAA) :", "Datum (TT/MM/JJJJ) :", "Data (GG/MM/AAAA) :", "Datum (DD/MM/JJJJ) :", "Datum (DD/MM/ÅÅÅÅ) :"
    Upsert_Dico tblDic, "FRM_COMPTE", "Compte :", "Account :", "Cuenta :", "Conta :", "Konto :", "Conto :", "Rekening :", "Konto :"
    Upsert_Dico tblDic, "FRM_CAT", "Catégorie :", "Category :", "Categoría :", "Categoria :", "Kategorie :", "Categoria :", "Categorie :", "Kategori :"
    Upsert_Dico tblDic, "FRM_TIERS", "Tiers :", "Payee/Payer :", "Tercero :", "Terceiro :", "Partei :", "Terzo :", "Partij :", "Part :"
    Upsert_Dico tblDic, "FRM_MONTANT", "Montant :", "Amount :", "Monto :", "Valor :", "Betrag :", "Importo :", "Bedrag :", "Belopp :"
    Upsert_Dico tblDic, "FRM_DEVISE", "Devise :", "Currency :", "Divisa :", "Moeda :", "Währung :", "Valuta :", "Valuta :", "Valuta :"
    Upsert_Dico tblDic, "FRM_DESC", "Notes :", "Notes :", "Notas :", "Notas :", "Notizen :", "Note :", "Notities :", "Anteckningar :"
    Upsert_Dico tblDic, "FRM_NEW", "Précisez :", "Specify :", "Especificar :", "Especificar :", "Angeben :", "Specifica :", "Specificeer :", "Ange :"
    Upsert_Dico tblDic, "FRM_SAVE", "ENREGISTRER", "SAVE", "GUARDAR", "SALVAR", "SPEICHERN", "SALVA", "OPSLAAN", "SPARA"
    Upsert_Dico tblDic, "FRM_CANCEL", "ANNULER", "CANCEL", "CANCELAR", "CANCELAR", "ABBRECHEN", "ANNULLA", "ANNULEREN", "AVBRYT"
    
    Upsert_Dico tblDic, "MSG_ERR_MISSING", "Sélection incomplète.", "Missing selection.", "Selección incompleta.", "Seleção incompleta.", "Fehlende Auswahl.", "Selezione incompleta.", "Ontbrekende selectie.", "Saknad markering."
    Upsert_Dico tblDic, "MSG_ERR_AMT", "Montant invalide.", "Invalid amount.", "Monto inválido.", "Valor inválido.", "Ungültiger Betrag.", "Importo non valido.", "Ongeldig bedrag.", "Ogiltigt belopp."
    Upsert_Dico tblDic, "MSG_OK", "Enregistré avec succès !", "Saved successfully!", "¡Guardado con éxito!", "Salvo com sucesso!", "Erfolgreich gespeichert!", "Salvato con successo!", "Succesvol opgeslagen!", "Sparad!"
    
    Upsert_Dico tblDic, "TT_F_DATE", "Saisissez la date.", "Enter date.", "Ingrese fecha.", "Insira data.", "Datum eingeben.", "Inserisci data.", "Voer datum in.", "Ange datum."
    Upsert_Dico tblDic, "TT_F_COMPTE", "Choisissez le compte.", "Choose account.", "Elija cuenta.", "Escolha conta.", "Konto wählen.", "Scegli conto.", "Kies rekening.", "Välj konto."
    Upsert_Dico tblDic, "TT_F_CAT", "Sélectionnez la catégorie.", "Select category.", "Seleccione categoría.", "Selecione categoria.", "Kategorie wählen.", "Seleziona categoria.", "Selecteer categorie.", "Välj kategori."
    Upsert_Dico tblDic, "TT_F_TIERS", "Tiers lié.", "Related party.", "Tercero relacionado.", "Terceiro relacionado.", "Zugehörige Partei.", "Parte correlata.", "Gerelateerde partij.", "Relaterad part."
    Upsert_Dico tblDic, "TT_F_MONTANT", "Valeur absolue.", "Absolute value.", "Valor absoluto.", "Valor absoluto.", "Absoluter Wert.", "Valore assoluto.", "Absolute waarde.", "Absolut värde."
    Upsert_Dico tblDic, "TT_F_DEVISE", "Monnaie.", "Currency.", "Moneda.", "Moeda.", "Währung.", "Valuta.", "Valuta.", "Valuta."
    Upsert_Dico tblDic, "TT_F_DESC", "Notes.", "Notes.", "Notas.", "Notas.", "Notizen.", "Note.", "Notities.", "Anteckningar."
    Upsert_Dico tblDic, "TT_F_SAVE", "Enregistrer.", "Save.", "Guardar.", "Salvar.", "Speichern.", "Salva.", "Opslaan.", "Spara."
    Upsert_Dico tblDic, "TT_F_CANCEL", "Annuler.", "Cancel.", "Cancelar.", "Cancelar.", "Abbrechen.", "Annulla.", "Annuleren.", "Avbryt."
    
    ' Les Tooltips pour l'Explicit Auto-Learning
    Upsert_Dico tblDic, "TT_F_TYPE_CPT", "Type d'actif.", "Asset type.", "Tipo de activo.", "Tipo de ativo.", "Anlageklasse.", "Classe di attività.", "Activaklasse.", "Tillgångsklass."
    Upsert_Dico tblDic, "TT_F_TYPE_CAT", "Type de flux.", "Flow type.", "Tipo de flujo.", "Tipo de fluxo.", "Flusstyp.", "Tipo di flusso.", "Stroomtype.", "Flödestyp."
End Sub

Private Sub Upsert_Dico(tbl As ListObject, k As String, fr As String, en As String, es As String, pt As String, de As String, it As String, nl As String, sv As String)
    Dim i As Long: For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 1).Value = k Then Exit Sub
    Next i
    Dim nR As ListRow: Set nR = tbl.ListRows.Add
    nR.Range(1, 1).Value = k: nR.Range(1, 2).Value = fr: nR.Range(1, 3).Value = en: nR.Range(1, 4).Value = es
    nR.Range(1, 5).Value = pt: nR.Range(1, 6).Value = de: nR.Range(1, 7).Value = it: nR.Range(1, 8).Value = nl: nR.Range(1, 9).Value = sv
End Sub

' -------------------------------------------------------------------------
' LE CERVEAU INJECTÉ (CODE BEHIND)
' -------------------------------------------------------------------------
Private Function Code_VBA_Formulaire() As String
    Dim L() As String: ReDim L(1 To 200): Dim i As Integer: i = 1
    
    L(i) = "Option Explicit": i = i + 1
    L(i) = "Private Function TR(Clé As String) As String": i = i + 1
    L(i) = "    TR = MOD_02_AppHome_Global.TR(Clé)": i = i + 1
    L(i) = "End Function": i = i + 1
    
    L(i) = "Private Sub UserForm_Initialize()": i = i + 1
    L(i) = "    MOD_03_Gatekeeper.Garantir_Lexique_Formulaire": i = i + 1
    L(i) = "    Me.Caption = TR(""FRM_TITLE"")": i = i + 1
    L(i) = "    Me.lbl_txt_Date.Caption = TR(""FRM_DATE"")": i = i + 1
    L(i) = "    Me.lbl_cmb_Compte.Caption = TR(""FRM_COMPTE"")": i = i + 1
    L(i) = "    Me.lbl_cmb_Categorie.Caption = TR(""FRM_CAT"")": i = i + 1
    L(i) = "    Me.lbl_cmb_Tiers.Caption = TR(""FRM_TIERS"")": i = i + 1
    L(i) = "    Me.lbl_txt_Montant.Caption = TR(""FRM_MONTANT"")": i = i + 1
    L(i) = "    Me.lbl_cmb_Devise.Caption = TR(""FRM_DEVISE"")": i = i + 1
    L(i) = "    Me.lbl_txt_Description.Caption = TR(""FRM_DESC"")": i = i + 1
    L(i) = "    Me.lbl_txt_New_Compte.Caption = TR(""FRM_NEW"")": i = i + 1
    L(i) = "    Me.lbl_txt_New_Categorie.Caption = TR(""FRM_NEW"")": i = i + 1
    L(i) = "    Me.lbl_txt_New_Tiers.Caption = TR(""FRM_NEW"")": i = i + 1
    L(i) = "    Me.btn_Save.Caption = TR(""FRM_SAVE"")": i = i + 1
    L(i) = "    Me.btn_Cancel.Caption = TR(""FRM_CANCEL"")": i = i + 1
    L(i) = "    Me.txt_Date.ControlTipText = TR(""TT_F_DATE"")": i = i + 1
    L(i) = "    Me.cmb_Compte.ControlTipText = TR(""TT_F_COMPTE"")": i = i + 1
    L(i) = "    Me.cmb_Categorie.ControlTipText = TR(""TT_F_CAT"")": i = i + 1
    L(i) = "    Me.cmb_Tiers.ControlTipText = TR(""TT_F_TIERS"")": i = i + 1
    L(i) = "    Me.txt_Montant.ControlTipText = TR(""TT_F_MONTANT"")": i = i + 1
    L(i) = "    Me.cmb_Devise.ControlTipText = TR(""TT_F_DEVISE"")": i = i + 1
    L(i) = "    Me.txt_Description.ControlTipText = TR(""TT_F_DESC"")": i = i + 1
    L(i) = "    Me.btn_Save.ControlTipText = TR(""TT_F_SAVE"")": i = i + 1
    L(i) = "    Me.btn_Cancel.ControlTipText = TR(""TT_F_CANCEL"")": i = i + 1
    
    L(i) = "    Me.cmb_New_Cpt_Type.ControlTipText = TR(""TT_F_TYPE_CPT"")": i = i + 1
    L(i) = "    Me.cmb_New_Cat_Type.ControlTipText = TR(""TT_F_TYPE_CAT"")": i = i + 1
    
    L(i) = "    Me.txt_Date.Value = Format(Date, ""dd/mm/yyyy"")": i = i + 1
    L(i) = "    Me.cmb_Devise.List = Array(""MUR"", ""EUR"", ""USD"", ""GBP"", ""ZAR"", ""OXF"")": i = i + 1
    L(i) = "    Me.cmb_Devise.ListIndex = 0": i = i + 1
    
    L(i) = "    Me.cmb_New_Cpt_Type.List = Array(""LIQUIDITE"", ""INVESTISSEMENT"", ""DETTE"")": i = i + 1
    L(i) = "    Me.cmb_New_Cpt_Type.ListIndex = 0": i = i + 1
    L(i) = "    Me.cmb_New_Cat_Type.List = Array(""DEPENSE"", ""REVENU"", ""TRANSFERT"")": i = i + 1
    L(i) = "    Me.cmb_New_Cat_Type.ListIndex = 0": i = i + 1
    
    L(i) = "    Charger_Combo Me.cmb_Compte, ""T_DIM_Compte""": i = i + 1
    L(i) = "    Charger_Combo Me.cmb_Categorie, ""T_DIM_Categorie""": i = i + 1
    L(i) = "    Charger_Combo Me.cmb_Tiers, ""T_DIM_Tiers""": i = i + 1
    L(i) = "End Sub": i = i + 1
    
    L(i) = "Private Sub Charger_Combo(cmb As MSForms.ComboBox, NomTable As String)": i = i + 1
    L(i) = "    Dim tbl As ListObject, k As Long": i = i + 1
    L(i) = "    On Error Resume Next: Set tbl = ThisWorkbook.Sheets(Split(NomTable, ""_"", 2)(1)).ListObjects(NomTable): On Error GoTo 0": i = i + 1
    L(i) = "    If tbl Is Nothing Then Exit Sub": i = i + 1
    L(i) = "    cmb.Clear: cmb.ColumnCount = 2: cmb.ColumnWidths = ""0 pt;200 pt""": i = i + 1
    L(i) = "    If tbl.ListRows.Count > 0 Then": i = i + 1
    L(i) = "        For k = 1 To tbl.ListRows.Count": i = i + 1
    L(i) = "            If Trim(tbl.ListRows(k).Range(1, 2).Value) <> """" Then": i = i + 1
    L(i) = "                cmb.AddItem tbl.ListRows(k).Range(1, 1).Value": i = i + 1
    L(i) = "                cmb.List(cmb.ListCount - 1, 1) = tbl.ListRows(k).Range(1, 2).Value": i = i + 1
    L(i) = "            End If": i = i + 1
    L(i) = "        Next k": i = i + 1
    L(i) = "    End If": i = i + 1
    L(i) = "End Sub": i = i + 1
    
    ' LES ÉVÉNEMENTS QUI AFFICHENT LE DOUBLE CONTRÔLE (Nom + Typage)
    L(i) = "Private Sub cmb_Compte_Change(): Gerer_Visibilite_Double Me.cmb_Compte, Me.txt_New_Compte, Me.cmb_New_Cpt_Type, Me.lbl_txt_New_Compte: End Sub": i = i + 1
    L(i) = "Private Sub cmb_Categorie_Change(): Gerer_Visibilite_Double Me.cmb_Categorie, Me.txt_New_Categorie, Me.cmb_New_Cat_Type, Me.lbl_txt_New_Categorie: End Sub": i = i + 1
    L(i) = "Private Sub cmb_Tiers_Change(): Gerer_Visibilite_Simple Me.cmb_Tiers, Me.txt_New_Tiers, Me.lbl_txt_New_Tiers: End Sub": i = i + 1
    
    L(i) = "Private Sub Gerer_Visibilite_Double(cmb As MSForms.ComboBox, txt As MSForms.TextBox, cmbType As MSForms.ComboBox, lbl As MSForms.Label)": i = i + 1
    L(i) = "    Dim estAutre As Boolean: estAutre = (InStr(1, cmb.Text, ""Autre"", vbTextCompare) > 0 Or InStr(1, cmb.Text, ""Other"", vbTextCompare) > 0)": i = i + 1
    L(i) = "    txt.Visible = estAutre: lbl.Visible = estAutre: cmbType.Visible = estAutre": i = i + 1
    L(i) = "    If Not estAutre Then txt.Value = """"": i = i + 1
    L(i) = "End Sub": i = i + 1
    
    L(i) = "Private Sub Gerer_Visibilite_Simple(cmb As MSForms.ComboBox, txt As MSForms.TextBox, lbl As MSForms.Label)": i = i + 1
    L(i) = "    Dim estAutre As Boolean: estAutre = (InStr(1, cmb.Text, ""Autre"", vbTextCompare) > 0 Or InStr(1, cmb.Text, ""Other"", vbTextCompare) > 0)": i = i + 1
    L(i) = "    txt.Visible = estAutre: lbl.Visible = estAutre": i = i + 1
    L(i) = "    If Not estAutre Then txt.Value = """"": i = i + 1
    L(i) = "End Sub": i = i + 1
    
    L(i) = "Private Function Obtenir_ID(cmb As MSForms.ComboBox, txt As MSForms.TextBox, NomTable As String, TypeSelect As String) As Long": i = i + 1
    L(i) = "    If txt.Visible = False Then Obtenir_ID = CLng(cmb.List(cmb.ListIndex, 0)): Exit Function": i = i + 1
    L(i) = "    Dim valClean As String: valClean = MOD_01_CoreEngine.CLEAN_TEXT(txt.Value)": i = i + 1
    L(i) = "    Dim tbl As ListObject, ws As Worksheet, k As Long": i = i + 1
    L(i) = "    Set ws = ThisWorkbook.Sheets(Split(NomTable, ""_"", 2)(1)): Set tbl = ws.ListObjects(NomTable)": i = i + 1
    L(i) = "    For k = 1 To tbl.ListRows.Count": i = i + 1
    L(i) = "        If UCase(Trim(tbl.ListRows(k).Range(1, 2).Value)) = UCase(valClean) Then Obtenir_ID = tbl.ListRows(k).Range(1, 1).Value: Exit Function": i = i + 1
    L(i) = "    Next k": i = i + 1
    L(i) = "    ws.Unprotect ""SFP_ADMIN_2026""": i = i + 1
    L(i) = "    Dim newRow As ListRow: Set newRow = tbl.ListRows.Add": i = i + 1
    L(i) = "    Dim newID As Long: newID = MOD_01_CoreEngine.GENERER_NOUVEL_ID(NomTable)": i = i + 1
    L(i) = "    newRow.Range(1, 1).Value = newID: newRow.Range(1, 2).Value = valClean: newRow.Range(1, 3).Value = TypeSelect": i = i + 1
    L(i) = "    If NomTable = ""T_DIM_Compte"" Then newRow.Range(1, 4).Value = Me.cmb_Devise.Value": i = i + 1
    L(i) = "    ws.Protect ""SFP_ADMIN_2026"", UserInterfaceOnly:=True": i = i + 1
    L(i) = "    Obtenir_ID = newID": i = i + 1
    L(i) = "End Function": i = i + 1
    
    L(i) = "Private Sub btn_Save_Click()": i = i + 1
    L(i) = "    If Me.cmb_Compte.ListIndex = -1 Or Me.cmb_Categorie.ListIndex = -1 Or Me.cmb_Tiers.ListIndex = -1 Then MsgBox TR(""MSG_ERR_MISSING""), vbCritical: Exit Sub": i = i + 1
    L(i) = "    Dim strMontant As String, dblMontant As Double": i = i + 1
    L(i) = "    strMontant = Replace(Me.txt_Montant.Value, "","", ""."")": i = i + 1
    L(i) = "    dblMontant = Val(strMontant)": i = i + 1
    L(i) = "    If dblMontant <= 0 Then MsgBox TR(""MSG_ERR_AMT""), vbCritical: Exit Sub": i = i + 1
    L(i) = "    If Not IsDate(Me.txt_Date.Value) Then MsgBox TR(""MSG_ERR_MISSING""), vbCritical: Exit Sub": i = i + 1
    
    ' LE CŒUR DE LA CORRECTION : LE TYPAGE EXPLICITE SAISIT PAR L'UTILISATEUR
    L(i) = "    Dim idC As Long, idCat As Long, idT As Long": i = i + 1
    L(i) = "    idC = Obtenir_ID(Me.cmb_Compte, Me.txt_New_Compte, ""T_DIM_Compte"", Me.cmb_New_Cpt_Type.Value)": i = i + 1
    L(i) = "    idCat = Obtenir_ID(Me.cmb_Categorie, Me.txt_New_Categorie, ""T_DIM_Categorie"", Me.cmb_New_Cat_Type.Value)": i = i + 1
    L(i) = "    idT = Obtenir_ID(Me.cmb_Tiers, Me.txt_New_Tiers, ""T_DIM_Tiers"", ""AUTRE"")": i = i + 1
    
    L(i) = "    Dim wsFact As Worksheet: Set wsFact = ThisWorkbook.Sheets(""FACT_Transaction"")": i = i + 1
    L(i) = "    wsFact.Unprotect ""SFP_ADMIN_2026""": i = i + 1
    L(i) = "    Dim nR As ListRow: Set nR = wsFact.ListObjects(""T_FACT_Transaction"").ListRows.Add": i = i + 1
    L(i) = "    nR.Range(1, 1).Value = MOD_01_CoreEngine.GENERER_NOUVEL_ID(""T_FACT_Transaction"")": i = i + 1
    L(i) = "    nR.Range(1, 2).Value = DateValue(Me.txt_Date.Value)": i = i + 1
    L(i) = "    nR.Range(1, 3).Value = idC: nR.Range(1, 4).Value = idCat: nR.Range(1, 5).Value = idT": i = i + 1
    L(i) = "    nR.Range(1, 6).Value = dblMontant": i = i + 1
    L(i) = "    nR.Range(1, 7).Value = Me.cmb_Devise.Value": i = i + 1
    L(i) = "    nR.Range(1, 8).Value = MOD_01_CoreEngine.CLEAN_TEXT(Me.txt_Description.Value)": i = i + 1
    L(i) = "    nR.Range(1, 9).Value = Application.UserName: nR.Range(1, 10).Value = Now": i = i + 1
    L(i) = "    wsFact.Protect ""SFP_ADMIN_2026"", UserInterfaceOnly:=True": i = i + 1
    
    L(i) = "    MsgBox TR(""MSG_OK""), vbInformation: Unload Me": i = i + 1
    L(i) = "End Sub": i = i + 1
    
    L(i) = "Private Sub btn_Cancel_Click(): Unload Me: End Sub": i = i + 1
    
    ReDim Preserve L(1 To i - 1)
    Code_VBA_Formulaire = Join(L, vbCrLf)
End Function

