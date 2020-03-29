Attribute VB_Name = "D_Hand5Functions"
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Hand5 (5 cards) functions
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub getHand5ValuesAndSuits(hand5() As Long, hand5Values() As Long, hand5Suits() As Long)
    For i = 1 To 5
        hand5Values(i) = cardValue(hand5(i))
        hand5Suits(i) = cardSuit(hand5(i))
    Next i
End Sub

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Poker scores
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' [x] High card:        0.h5h4h3h2h1
' [x] 1 pair:           1.ppccbbaa      For x-x-a-b-c, a-x-x-b-c, a-b-x-x-c, a-b-c-x-x
' [x] 2 pairs:          2.qqppaa        For p-p-q-q-a, a-p-p-q-q, p-p-a-q-q
' [x] 3 of a kind       3.h5h4h3h2h1
' [x] Straight:         4.h5h4h3h2h1
' [x] Flush:            5.h5h4h3h2h1
' [x] Full house:       6.h5h4h3h2h1
' [x] Four of a kind:   7.h5h4h3h2h1
' [x] Straight flush:   8.h5h4h3h2h1

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Main function
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Function hand5Score(hand5() As Long)
    Dim hS(1 To 5) As Long
    Dim hV(1 To 5) As Long
    
    'Get hV and hS
    Call getHand5ValuesAndSuits(hand5, hV, hS)
    
    'Sort hV and hS (will be disconnected but that's ok). No need to re-sort on each boolean hand5 check
    Call Quicksort(hS, 1, 5)
    Call Quicksort(hV, 1, 5)
    
    'Calculate score based on hand5 config
    If isStraightFlush(hV, hS) Then
        hand5Score = scoreStraightFlush(hV)
    ElseIf is4s(hV) Then
        hand5Score = score4s(hV)
    ElseIf isFullHouse(hV) Then
        hand5Score = scoreFullHouse(hV)
    ElseIf isFlush(hS) Then
        hand5Score = scoreFlush(hV)
    ElseIf isStraight(hV) Then
        hand5Score = scoreStraight(hV)
    ElseIf is3s(hV) Then
        hand5Score = score3s(hV)
    ElseIf is22s(hV) Then
        hand5Score = score22s(hV)
    ElseIf is2s(hV) Then
        hand5Score = score2s(hV)
    Else
        hand5Score = scoreHighCard(hV)
    End If
    
    
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Flush logic (bool and score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function isFlush(hand5Suits() As Long)
    
    'Sort by suits (destructive. actually changes hand5Suits (byRef, not byValue. VBA doesn't pass arrays by value)
    'Call Quicksort(hand5Suits, 1, 5)
    
    'Check if first and last suits are equal and pass results (boolean) to final function value
    isFlush = hand5Suits(1) = hand5Suits(5)
      
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function scoreFlush(hand5Values() As Long)
    'Score = 5.eeddccbbaa (a-b-c-d-e hand5)
    
    scoreFlush = 5 + (hand5Values(5) / 100) + (hand5Values(4) / (100 ^ 2)) + (hand5Values(3) / (100 ^ 3)) + (hand5Values(2) / (100 ^ 4)) + (hand5Values(1) / (100 ^ 5))

End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Straight logic (bool and score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function isStraight(hand5Values() As Long)
    
    'Sort by hand5Values (destructive. actually changes hand5Values array (byRef, not byValue. VBA doesn't pass arrays by value)
    'Call Quicksort(hand5Values, 1, 5)
    
    'Assume it's not straight (false)
    isStraight = False
    
    'Now check if result is sequential
    If hand5Values(5) = hand5Values(4) + 1 Then
        If hand5Values(4) = hand5Values(3) + 1 Then
            If hand5Values(3) = hand5Values(2) + 1 Then
                If hand5Values(2) = hand5Values(1) + 1 Then
                    isStraight = True
                End If
            End If
        End If
    End If
    
    'Check for A-2-3-4-5 possibility
    If hand5Values(1) = 2 And hand5Values(2) = 3 And hand5Values(3) = 4 And hand5Values(4) = 5 And hand5Values(5) = 14 Then
        isStraight = True
    End If
      
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function scoreStraight(hand5Values() As Long)
    'Score = 4.eeddccbbaa (a-b-c-d-e hand5 except Ace to 5)
    'Score = 4.ddccbbaaee (Ace to 5 straight)
    
    scoreStraight = 0
    
    If hand5Values(5) = hand5Values(4) + 1 Then
        If hand5Values(4) = hand5Values(3) + 1 Then
            If hand5Values(3) = hand5Values(2) + 1 Then
                If hand5Values(2) = hand5Values(1) + 1 Then
                    scoreStraight = 4 + (hand5Values(5) / 100) + (hand5Values(4) / (100 ^ 2)) + (hand5Values(3) / (100 ^ 3)) + (hand5Values(2) / (100 ^ 4)) + (hand5Values(1) / (100 ^ 5))
                End If
            End If
        End If
    End If
    
    If hand5Values(1) = 2 And hand5Values(2) = 3 And hand5Values(3) = 4 And hand5Values(4) = 5 And hand5Values(5) = 14 Then
        scoreStraight = 4 + (hand5Values(4) / 100) + (hand5Values(3) / (100 ^ 2)) + (hand5Values(2) / (100 ^ 3)) + (hand5Values(1) / (100 ^ 4)) + (hand5Values(5) / (100 ^ 5))
    End If
    
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Straight Flush logic (bool and score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function isStraightFlush(hand5Values() As Long, hand5Suits() As Long)
    isStraightFlush = isStraight(hand5Values) And isFlush(hand5Suits)
End Function

Function scoreStraightFlush(hand5Values() As Long)
    scoreStraightFlush = scoreStraight(hand5Values) + 8 - 4 'same score logic as straight, only changing the integer part from 4 to 8 (making it the hardest hand5)
End Function
' 4 of a kind ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function is4s(hand5Values() As Long)
    
    'Sort hand5 values first
    'Call Quicksort(hand5Values, 1, 5)
    
    'Assume it's false
    is4s = False
    
    'Check for a-a-a-a-b and a-b-b-b-b
    If hand5Values(1) = hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) = hand5Values(4) Then
        is4s = True
    ElseIf hand5Values(2) = hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) = hand5Values(5) Then
        is4s = True
    End If
    
End Function

Function score4s(hand5Values() As Long)

    'Score: 7.qqaa (q-q-q-q-a, a-q-q-q-q)
    
    If hand5Values(1) = hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) = hand5Values(4) Then
        score4s = 7 + (hand5Values(1) / 100) + (hand5Values(5) / (100 ^ 2))
    ElseIf hand5Values(2) = hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) = hand5Values(5) Then
        score4s = 7 + (hand5Values(2) / 100) + (hand5Values(1) / (100 ^ 2))
    End If
    
End Function
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Full house logic (bool and score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function isFullHouse(hand5Values() As Long)
    
    'Sort hand5 values first
    'Call Quicksort(hand5Values, 1, 5)
    
    'Assume it's false
    isFullHouse = False
    
    'Check for a-a-b-b-b and a-a-a-b-b
    If hand5Values(1) = hand5Values(2) And hand5Values(3) = hand5Values(4) And hand5Values(4) = hand5Values(5) Then
        isFullHouse = True
    ElseIf hand5Values(1) = hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(4) = hand5Values(5) Then
        isFullHouse = True
    End If
    
End Function

Function scoreFullHouse(hand5Values() As Long)

    'Score: 6.ttpp (p-p-t-t-t, t-t-t-p-p)
    
    'p-p-t-t-t
    If hand5Values(1) = hand5Values(2) And hand5Values(3) = hand5Values(4) And hand5Values(4) = hand5Values(5) Then
        scoreFullHouse = 6 + (hand5Values(3) / 100) + (hand5Values(1) / (100 ^ 2))
    
    'p-p-p-t-t
    ElseIf hand5Values(1) = hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(4) = hand5Values(5) Then
        scoreFullHouse = 6 + (hand5Values(1) / 100) + (hand5Values(4) / (100 ^ 2))
    Else
        scoreFullHouse = 0
    End If

End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 3 of a kind logic (bool and score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function is3s(hand5Values() As Long)
    
    'Sort hand5 values first
    'Call Quicksort(hand5Values, 1, 5)
    
    'Assume it's false
    is3s = False
    
    'Check for x-x-x-a-b and a-x-x-x-b and a-b-x-x-x
    If hand5Values(1) = hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(3) <> hand5Values(5) And hand5Values(4) <> hand5Values(5) Then
        is3s = True
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) <> hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        is3s = True
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) = hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        is3s = True
    End If
    
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function score3s(hand5Values() As Long)
    
    'Score  3.ttbbaa (t-t-t-a-b, a-t-t-t-b, a-b-t-t-t)
    
    't-t-t-a-b
    If hand5Values(1) = hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(3) <> hand5Values(5) And hand5Values(4) <> hand5Values(5) Then
        score3s = 3 + (hand5Values(1) / 100) + (hand5Values(5) / (100 ^ 2)) + (hand5Values(4) / (100 ^ 3))
        
    'a-t-t-t-b
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) <> hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        score3s = 3 + (hand5Values(2) / 100) + (hand5Values(5) / (100 ^ 2)) + (hand5Values(1) / (100 ^ 3))
        
    'a-b-t-t-t
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) = hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        score3s = 3 + (hand5Values(3) / 100) + (hand5Values(2) / (100 ^ 2)) + (hand5Values(1) / (100 ^ 3))
    
    Else
        score3s = 0
        
    End If
    
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 2 pairs logic (bool and score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function is22s(hand5Values() As Long)
    
    'Sort hand5 values first
    'Call Quicksort(hand5Values, 1, 5)
    
    'Assume it's false
    is22s = False
    
    'Check for a-a-b-b-x and x-a-a-b-b and a-a-x-b-b
    If hand5Values(1) = hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) <> hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        is22s = True
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(4) = hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        is22s = True
    ElseIf hand5Values(1) = hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(4) = hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        is22s = True
    End If
    
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function score22s(hand5Values() As Long)
    
    'Score = 2.qqppaa (p-p-q-q-a, a-p-p-q-q, p-p-a-q-q)
    
    'p-p-q-q-a
    If hand5Values(1) = hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(3) = hand5Values(4) And hand5Values(4) <> hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        score22s = 2 + (hand5Values(3) / 100) + (hand5Values(1) / (100 ^ 2)) + (hand5Values(5) / (100 ^ 3))
    
    'a-p-p-q-q
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(2) = hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(4) = hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        score22s = 2 + (hand5Values(4) / 100) + (hand5Values(2) / (100 ^ 2)) + (hand5Values(1) / (100 ^ 3))
    
    'p-p-a-q-q
    ElseIf hand5Values(1) = hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(4) = hand5Values(5) And hand5Values(1) <> hand5Values(5) Then
        score22s = 2 + (hand5Values(4) / 100) + (hand5Values(1) / (100 ^ 2)) + (hand5Values(3) / (100 ^ 3))
    
    Else
        score22s = 0
    
    End If
    
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 2 of a kind logic (bool and score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function is2s(hand5Values() As Long)
    
    'Sort hand5 values first
    'Call Quicksort(hand5Values, 1, 5)
    
    'Assume it's false
    is2s = False
    
    'Check for x-x-a-b-c and a-x-x-b-c and a-b-x-x-c and a-b-c-x-x
    If hand5Values(1) = hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(2) <> hand5Values(4) And hand5Values(2) <> hand5Values(5) And hand5Values(3) <> hand5Values(4) And hand5Values(3) <> hand5Values(5) And hand5Values(4) <> hand5Values(5) Then
        is2s = True
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(1) <> hand5Values(4) And hand5Values(1) <> hand5Values(5) And hand5Values(2) = hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(3) <> hand5Values(5) And hand5Values(4) <> hand5Values(5) Then
        is2s = True
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(1) <> hand5Values(3) And hand5Values(1) <> hand5Values(5) And hand5Values(2) <> hand5Values(3) And hand5Values(2) <> hand5Values(5) And hand5Values(3) = hand5Values(4) And hand5Values(4) <> hand5Values(5) Then
        is2s = True
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(1) <> hand5Values(3) And hand5Values(1) <> hand5Values(4) And hand5Values(2) <> hand5Values(3) And hand5Values(2) <> hand5Values(4) And hand5Values(3) <> hand5Values(4) And hand5Values(4) = hand5Values(5) Then
        is2s = True
    End If
    
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function score2s(hand5Values() As Long)
    'Score = 1.xxccbbaa (pair, highest isolated, second highest isolated, lowest)
    
    'x-x-a-b-c
    If hand5Values(1) = hand5Values(2) And hand5Values(2) <> hand5Values(3) And hand5Values(2) <> hand5Values(4) And hand5Values(2) <> hand5Values(5) And hand5Values(3) <> hand5Values(4) And hand5Values(3) <> hand5Values(5) And hand5Values(4) <> hand5Values(5) Then
        score2s = 1 + (hand5Values(1) / 100) + (hand5Values(5) / (100 ^ 2)) + (hand5Values(4) / (100 ^ 3)) + (hand5Values(3) / (100 ^ 4))
    
    'a-x-x-b-c
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(1) <> hand5Values(4) And hand5Values(1) <> hand5Values(5) And hand5Values(2) = hand5Values(3) And hand5Values(3) <> hand5Values(4) And hand5Values(3) <> hand5Values(5) And hand5Values(4) <> hand5Values(5) Then
        score2s = 1 + (hand5Values(2) / 100) + (hand5Values(5) / (100 ^ 2)) + (hand5Values(4) / (100 ^ 3)) + (hand5Values(1) / (100 ^ 4))
    
    'a-b-x-x-c
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(1) <> hand5Values(3) And hand5Values(1) <> hand5Values(5) And hand5Values(2) <> hand5Values(3) And hand5Values(2) <> hand5Values(5) And hand5Values(3) = hand5Values(4) And hand5Values(4) <> hand5Values(5) Then
        score2s = 1 + (hand5Values(3) / 100) + (hand5Values(5) / (100 ^ 2)) + (hand5Values(2) / (100 ^ 3)) + (hand5Values(1) / (100 ^ 4))
    
    'a-b-c-x-x
    ElseIf hand5Values(1) <> hand5Values(2) And hand5Values(1) <> hand5Values(3) And hand5Values(1) <> hand5Values(4) And hand5Values(2) <> hand5Values(3) And hand5Values(2) <> hand5Values(4) And hand5Values(3) <> hand5Values(4) And hand5Values(4) = hand5Values(5) Then
        score2s = 1 + (hand5Values(4) / 100) + (hand5Values(3) / (100 ^ 2)) + (hand5Values(2) / (100 ^ 3)) + (hand5Values(1) / (100 ^ 4))
    
    'Just to be safe
    Else
        score2s = 0
    End If
End Function

' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' High card logic (score)
' ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function scoreHighCard(hand5Values() As Long)
    scoreHighCard = (hand5Values(5) / 100) + (hand5Values(4) / (100 ^ 2)) + (hand5Values(3) / (100 ^ 3)) + (hand5Values(2) / (100 ^ 4)) + (hand5Values(1) / (100 ^ 5))
End Function

