Attribute VB_Name = "M99_Sales_KeyGen"
' =========================================================================
' OUTIL ADMIN SECRET : SFP KEYGEN (GÉNÉRATEUR DE LICENCES)
' À conserver dans un fichier Excel strictement privé.
' =========================================================================
Public Sub GENERER_LICENCE_CLIENT()
    Dim MachineID As String
    MachineID = InputBox("Saisissez l'ID MACHINE fourni par le client :", "SFP KEYGEN 1.0")
    If MachineID = "" Then Exit Sub
    
    Dim clientKey As String
    clientKey = CalculateValidKey(UCase(Trim(MachineID)))
    
    ' Affiche la clé à envoyer au client
    InputBox "Copiez cette clé d'activation et envoyez-la au client :", "CLÉ GÉNÉRÉE AVEC SUCCÈS", clientKey
End Sub

' L'ALGORITHME DE HACHAGE ASYMÉTRIQUE (LE SECRET INDUSTRIEL)
Private Function CalculateValidKey(MachineID As String) As String
    Dim cleanID As String: cleanID = Replace(MachineID, "-", "")
    Dim i As Integer, total As Long: total = 0
    ' Mathématique de salage (Salt)
    For i = 1 To Len(cleanID)
        total = total + (Asc(Mid(cleanID, i, 1)) * i * 73)
    Next i
    Dim p1 As String: p1 = UCase(Hex(total * 13))
    Dim p2 As String: p2 = UCase(Hex(total * 7))
    CalculateValidKey = Right("0000" & p1, 4) & "-" & Right("0000" & p2, 4)
End Function
