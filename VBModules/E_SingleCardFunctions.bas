Attribute VB_Name = "E_SingleCardFunctions"
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Single card functions
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function cardValue(cardNo)
    If cardNo = 0 Then
        cardValue = 0
    Else
        cardValue = cardNo Mod 13
        If cardValue = 0 Then cardValue = 13
        cardValue = cardValue + 1 'to start at 2. Ace's value is 14
    End If
End Function

Function cardSuit(cardNo)
    If cardNo = 0 Then
        cardSuit = 0
    Else
        cardSuit = Int((cardNo - 1) / 13) + 1
    End If
End Function

