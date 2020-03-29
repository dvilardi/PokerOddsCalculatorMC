Attribute VB_Name = "F_StartingHands"
Sub generateStartingHandsData()
    
    ' ---------------------------------------------------------------------
    ' Define worksheets
    ' ---------------------------------------------------------------------
    Dim shTable As Worksheet
    Dim shAux As Worksheet
    Dim shStarting As Worksheet
    
    Set shTable = ThisWorkbook.Sheets("Table")
    Set shAux = ThisWorkbook.Sheets("Aux")
    Set shStarting = ThisWorkbook.Sheets("StartingHands")
    
    ' ---------------------------------------------------------------------
    ' Number of players
    ' ---------------------------------------------------------------------
    nPlayers = 6
    shTable.Range("NumberOfPlayers").Value = nPlayers
    
    ' ---------------------------------------------------------------------
    ' Define table to generate
    ' ---------------------------------------------------------------------
    Dim rngStarting As Range
    rngName = "starting" & nPlayers & "players"
    Set rngStarting = shStarting.Range(rngName)
    
    ' ---------------------------------------------------------------------
    ' Define loop ranges (for work splitting between instances)
    ' ---------------------------------------------------------------------
    iStart = 1
    iEnd = 2
    
    jStart = 1
    jEnd = 13
    
    
    nHands = (iEnd - iStart + 1) * (jEnd - jStart + 1)
    ' ---------------------------------------------------------------------
    ' Number of monte-carlo simulations
    ' ---------------------------------------------------------------------
    nMC = 10000
    
    shTable.Range("NumberOfSimulations").Value = nMC
    
    ' ---------------------------------------------------------------------
    ' Prepare table (clear pot, clear my hands
    ' ---------------------------------------------------------------------
    Call ClearBoth
        
    ' ---------------------------------------------------------------------
    ' Begin loop
    ' ---------------------------------------------------------------------
    ' The table contains results for suited and off-suit combinations
    ' On the loop, if j is higher or equal to i, the combination is suited
    ' Suited: top right of table. Off suit: bottom left of table
    ' If suited: Card1 = 14-i and Card2 = 14-j (clubs-clubs)
    ' Off suit:  Card1 = 14-i and Card2 = 14-j+13 (clubs-hearts)
    ' Other combinations (diamonds, spades) have the same result.
    ' No need to evaluate all
    h = 0
    
    For i = iStart To iEnd
        'Card 1 number
        Card1 = 14 - i
        
        'Place onto aux
        shAux.Range("handIDs").Cells(1, 1).Value = Card1
            
        For j = jStart To jEnd
            
            'Card 2 number
            If j >= i Then
                Card2 = 14 - j
            Else
                Card2 = 14 - j + 13
            End If
            
            'Place onto aux
            shAux.Range("handIDs").Cells(1, 2).Value = Card2
            
            'Call monte carlo script (false to not show msgbox with results)
            Call texasMonteCarlo(False)
            
            'Get final win rate result and paste on table
            winRate = shTable.Range("WinLoseTie").Cells(1, 1).Value
            rngStarting.Cells(i, j).Value = winRate
            
            'Print progress
            h = h + 1
            Debug.Print "--------------------------------------------------------------------------------------------------------------------------------"
            Debug.Print " Starting hand table progress"
            Debug.Print "--------------------------------------------------------------------------------------------------------------------------------"
            Debug.Print "i = " & i & " from " & iStart & " to " & iEnd
            Debug.Print "j = " & j & " from " & jStart & " to " & jEnd
            Debug.Print "h = " & h & " / " & nHands & " (" & Format(h / nHands, "0.0%") & ")"
            Debug.Print "--------------------------------------------------------------------------------------------------------------------------------"
            
            'Doevents
        Next j
    Next i
    
End Sub


Function pk(pn, n, k)
    pk = pn * (n - 1) / (k - 1 + (n - k) * pn)
End Function
