Attribute VB_Name = "B_MonteCarlo"

Sub texasMonteCarloNsims()
    
    'Will be called by buttons 1K, 2K, 5K, 10K
    'Buttons are named "MC1000" "MC2000" "MC5000" or "MC10000"
    'This script will get the value from the button name (removing the "MC"), convert to long, place on the number of simulations cell and run the main montecarlo
    
    btnName = Application.Caller
    nSims = CLng(Right(btnName, Len(btnName) - 2))
    ThisWorkbook.Sheets("Table").Range("NumberOfSimulations").Value = nSims
    Call texasMonteCarlo
   
End Sub

Sub texasMonteCarlo(Optional showMsg As Boolean = False)
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 0 - Speed up
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False
        
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 1 - Define worksheet
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim shTable As Worksheet
    Set shTable = ThisWorkbook.Sheets("Table")
    Dim shAux As Worksheet
    Set shAux = ThisWorkbook.Sheets("Aux")
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 2 - Dim initial arrays
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim h2(1 To 2, 1 To 9) As Long  'all hands from 1 to 9. Index 1: my cards; 2 to 9: oponents
    Dim pot(1 To 5) As Long         'pot
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 3 - My cards
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    h2(1, 1) = shAux.Range("handIDs").Cells(1, 1).Value
    h2(2, 1) = shAux.Range("handIDs").Cells(1, 2).Value
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 4 - Pot cards
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    pot(1) = shAux.Range("potIDs").Cells(1, 1).Value
    pot(2) = shAux.Range("potIDs").Cells(1, 2).Value
    pot(3) = shAux.Range("potIDs").Cells(1, 3).Value
    pot(4) = shAux.Range("potIDs").Cells(1, 4).Value
    pot(5) = shAux.Range("potIDs").Cells(1, 5).Value
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 5 - Number of players
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    nPlayers = shTable.Range("NumberOfPlayers").Cells(1, 1).Value
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 6 - Number of simulations
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    n = shTable.Range("NumberOfSimulations").Cells(1, 1).Value
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 7 - Loop monte carlo
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim h2Temp(1 To 2) As Long
    Dim potTemp(1 To 5) As Long
    Dim usedCards(1 To 23) As Long '9 players x 2 cards each = 18 + 5 from pot = 23
    Dim scores(1 To 9) As Double 'scores from player 1 to player 9 for each iteration
    Dim countWin As Long
    Dim countLose As Long
    Dim countTie As Long
    countWin = 0
    countLose = 0
    countTie = 0
        
    For i = 1 To n
        
        'At beginning of monte carlo iteration, restore bucket of usedCards to empty
        usedCount = 0
        For k = 1 To 23
            usedCards(k) = 0
        Next k
        
        'Add my cards to the bucket
        usedCards(usedCount + 1) = h2(1, 1)
        usedCards(usedCount + 2) = h2(2, 1)
        usedCount = usedCount + 2
        
        ' Create random table (+8 hands and the pot) --------------------------------------------------------------------------------------------------------------
        
        'Random pot (need to be first because there may already be cards on the pot)
        For p = 1 To 5
            If pot(p) <> 0 Then
                potTemp(p) = pot(p)
            Else
                potTemp(p) = randomExcl(usedCards)
            End If
            usedCount = usedCount + 1
            usedCards(usedCount) = potTemp(p)
            'Doevents
        Next p
        
        'Shuffle all other 8 player's cards and place them on the bucket of used cards
        For p = 2 To 9
            
            If p <= nPlayers Then
                'Add one card to player, then add it to the bucket, then add the other card
                h2(1, p) = randomExcl(usedCards): usedCount = usedCount + 1: usedCards(usedCount) = h2(1, p)
                h2(2, p) = randomExcl(usedCards): usedCount = usedCount + 1: usedCards(usedCount) = h2(2, p)
            Else
                h2(1, p) = 0: usedCount = usedCount + 1: usedCards(usedCount) = h2(1, p)
                h2(2, p) = 0: usedCount = usedCount + 1: usedCards(usedCount) = h2(2, p)
            End If
            'Doevents
        Next p
        
        ' Debug print table --------------------------------------------------------------------------------------------------------------
        'strBar = "---------------------------------------------------------------------------------------"
        'strTitle = "P1" & Chr(9) & "P2" & Chr(9) & "P3" & Chr(9) & "P4" & Chr(9) & "P5" & Chr(9) & "P6" & Chr(9) & "P7" & Chr(9) & "P8" & Chr(9) & "P9" & Chr(9) & "Pot"
        'strValue = ""
        'For p = 1 To 9
        '    If p = 1 Then
        '        strTitle = "P" & p
        '        strValue = Format(h2(1, p), "00") & "," & Format(h2(2, p), "00")
        '    Else
        '        strTitle = strTitle & Chr(9) & Chr(9) & "P" & p
        '        strValue = strValue & Chr(9) & Format(h2(1, p), "00") & "," & Format(h2(2, p), "00")
        '    End If
        'Next p
        'strTitle = strTitle & Chr(9) & Chr(9) & "Pot"
        'strValue = strValue & Chr(9) & Format(potTemp(1), "00") & "," & Format(potTemp(2), "00") & "," & Format(potTemp(3), "00") & "," & Format(potTemp(4), "00") & "," & Format(potTemp(5), "00")
        'strOut = strBar & Chr(10) & strTitle & Chr(10) & strValue
        'Debug.Print strOut
        
        'Calculate 9 texas scores
        For p = 1 To 9
            If p <= nPlayers Then
                'pass from 3D array to 2D array
                h2Temp(1) = h2(1, p)
                h2Temp(2) = h2(2, p)
                scores(p) = texasScore(h2Temp, potTemp)
            Else
                scores(p) = 0
            End If
            'Doevents
        Next p
        
        'Get score of player 1
        myScore = scores(1)
        
        'Sort scores
        Call scoreSort(scores, 1, 9)
        
        'Bool to check for victory
        If myScore = scores(9) Then
            If myScore = scores(8) Then
                countTie = countTie + 1
            Else
                countWin = countWin + 1
            End If
        Else
            countLoss = countLoss + 1
        End If
        
        'Write on convergence test
        'ThisWorkbook.Sheets("CT9").Cells(i + 1, 1).Value = i
        'ThisWorkbook.Sheets("CT9").Cells(i + 1, 2).Value = countWin / i
                
        If i Mod 200 = 0 Then
            Debug.Print i & " / " & n & " (W: " & Format(countWin / i, "0.0%") & " / L: " & Format(countLoss / i, "0.0%") & " / T: " & Format(countTie / i, "0.0%") & ")"
            Application.StatusBar = i & " / " & n & " (W: " & Format(countWin / i, "0.0%") & " / L: " & Format(countLoss / i, "0.0%") & " / T: " & Format(countTie / i, "0.0%") & ")"
        End If
    Next i
                
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 8 - Calculate win and tie result
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    rateWin = CDbl(countWin) / CDbl(n)
    rateLoss = CDbl(countLoss) / CDbl(n)
    rateTie = CDbl(countTie) / CDbl(n)
    
    shTable.Range("WinLoseTie").Cells(1, 1).Value = rateWin
    shTable.Range("WinLoseTie").Cells(1, 2).Value = rateTie
    shTable.Range("WinLoseTie").Cells(1, 3).Value = rateLoss
    
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 0 - Speed down
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    If showMsg Then
        strMsg = "Win: " & Format(rateWin, "0.0%") & vbNewLine & _
        "Tie: " & Format(rateTie, "0.0%") & vbNewLine & _
        "Loss: " & Format(rateLoss, "0.0%")
        
        MsgBox strMsg
    End If
            
    
End Sub

Sub testMatch()
    a = 5
    Dim b(1 To 5) As Long
    
    b(1) = 1
    b(2) = 2
    b(3) = 3
    b(4) = 4
    b(5) = 5
    
    Debug.Print isInArray(a, b)
End Sub

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Auxiliary functions
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function isInArray(x, arr)
    
    'isInArray = Not (IsError(Application.match(x, arr, 0)))
    
    isInArray = False
    For k = LBound(arr) To UBound(arr)
        If x = arr(k) Then
            isInArray = True
            Exit For
        ElseIf arr(k) = 0 Then
            Exit For
        End If
    Next k
    
End Function

Function randomCard() 'random card from 1 to 52
    Randomize
    randomCard = Int(1 + Rnd * 52)
End Function

Function randomExcl(used() As Long)
    ok = False
    Do While ok = False
        r = randomCard()
        ok = Not (isInArray(r, used))
    Loop
    randomExcl = r
End Function


