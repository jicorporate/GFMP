Attribute VB_Name = "M99_Sales_KeyGen"
Public Sub GENERER_LICENCE_CLIENT()
    Dim MachineID As String
    MachineID = InputBox("Saisissez l'ID MACHINE fourni par le client :", "SFP KEYGEN 2.0")
    If MachineID = "" Then Exit Sub
    
    ' Appelle l'algorithme secret
    Dim clientKey As String
    clientKey = CalculateValidKey(UCase(Trim(MachineID)))
    
    InputBox "Copiez cette clé d'activation et envoyez-la au client :", "CLÉ GÉNÉRÉE AVEC SUCCÈS", clientKey
End Sub

Private Function CalculateValidKey(MachineID As String) As String
    Dim cleanID As String: cleanID = Replace(MachineID, "-", "")
    Dim i As Integer, total As Long: total = 0
    For i = 1 To Len(cleanID)
        ' FORÇAGE LONG (CLng et 73&) POUR ÉVITER L'OVERFLOW DU KEYGEN
        total = total + (CLng(Asc(Mid(cleanID, i, 1))) * CLng(i) * 73&)
    Next i
    Dim p1 As String: p1 = UCase(Hex(total * 13))
    Dim p2 As String: p2 = UCase(Hex(total * 7))
    CalculateValidKey = Right("0000" & p1, 4) & "-" & Right("0000" & p2, 4)
End Function
