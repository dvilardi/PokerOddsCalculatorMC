Attribute VB_Name = "C_TexasFunctions"
Function texasScore(hand2() As Long, pot() As Long)
    Dim hand5temp(1 To 5) As Long
    Dim hand7(1 To 7) As Long
    
    'Create an array of 7 cards
    hand7(1) = hand2(1):    hand7(2) = hand2(2)
    hand7(3) = pot(1):    hand7(4) = pot(2):    hand7(5) = pot(3):    hand7(6) = pot(4):    hand7(7) = pot(5)
    
    'Initialize score as zero
    maxScore = 0
    
    'Loop all texas holdem possibilities (5 card combination within 7 possible picks)
    For i = 1 To 7
        For j = i + 1 To 7
            For k = j + 1 To 7
                For l = k + 1 To 7
                    For m = l + 1 To 7
                        'Place hand on temporary array
                        hand5temp(1) = hand7(i)
                        hand5temp(2) = hand7(j)
                        hand5temp(3) = hand7(k)
                        hand5temp(4) = hand7(l)
                        hand5temp(5) = hand7(m)
                        
                        'Calculate score of temporary hand
                        scoreTemp = hand5Score(hand5temp)
                                                
                        'Compare with max
                        If scoreTemp > maxScore Then
                            maxScore = scoreTemp
                        End If
                        
                        'Doevents
                    Next m
                Next l
            Next k
        Next j
    Next i
                        
    'Pass max possible score to output
    texasScore = maxScore
    
End Function

