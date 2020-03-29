Attribute VB_Name = "A_DashboardControls"
Sub clearMyCards()
    'Clear cards on dashboard
    'Reset last edited to zero
    ThisWorkbook.Sheets("Table").Range("MyCards").ClearContents
    ThisWorkbook.Sheets("Aux").Range("LastEditedMyCards").Value = 0
    ThisWorkbook.Sheets("Aux").Range("handIDs").Value = 0
End Sub

Sub clearPot()
    'Clear cards on dashboard
    'Reset last edited to zero
    ThisWorkbook.Sheets("Table").Range("Pot").ClearContents
    ThisWorkbook.Sheets("Aux").Range("LastEditedPot").Value = 0
    ThisWorkbook.Sheets("Aux").Range("potIDs").Value = 0
End Sub

Sub ClearBoth()
    Call clearMyCards
    Call clearPot
    ThisWorkbook.Sheets("Aux").Range("CurrentlyEditing").Value = 1
End Sub

Sub changeNoOfPlayers()
    btnName = Application.Caller
    
    nPlayers = CInt(Mid(btnName, 4, 1)) 'btnName structure is btnXplayers (wanna get X as an integer)
    ThisWorkbook.Sheets("Table").Range("NumberOfPlayers").Value = nPlayers
    
End Sub

Sub placeCard()
    
    'Get card name
    cardName = Application.Caller
    
    'Get card number
    cardNo = CInt(Right(cardName, 2))
    
    'Row containing cardNo on card table
    rowAux = Application.WorksheetFunction.match(cardNo, ThisWorkbook.Sheets("Aux").ListObjects("CardDB").DataBodyRange.Columns(3), 0)
    
    'Final card string
    strAux = ThisWorkbook.Sheets("Aux").ListObjects("CardDB").DataBodyRange.Cells(rowAux, 5).Value
    
    'Check whether to edit my cards or pot
    toEdit = ThisWorkbook.Sheets("Aux").Range("CurrentlyEditing").Value
    
    'Check last edited and maxEdit
    If toEdit = 1 Then 'my cards
        lastEdited = ThisWorkbook.Sheets("Aux").Range("LastEditedMyCards").Value
        maxEdit = 2
    Else 'pot
        lastEdited = ThisWorkbook.Sheets("Aux").Range("LastEditedPot").Value
        maxEdit = 5
    End If
    
    'Add one to lastEdited
    If lastEdited >= maxEdit Then
        lastEdited = 1
    Else
        lastEdited = lastEdited + 1
    End If
    
    'Write on table and color
    If toEdit = 1 Then 'my cards
        
        'Write on table
        ThisWorkbook.Sheets("Table").Range("MyCards").Cells(1, lastEdited).Value = strAux
        
        'Write on aux
        ThisWorkbook.Sheets("Aux").Range("LastEditedMyCards").Value = lastEdited
        ThisWorkbook.Sheets("Aux").Range("handIDs").Cells(1, lastEdited).Value = cardNo
        
        'Color
        If cardSuit(cardNo) = 2 Or cardSuit(cardNo) = 4 Then
            ThisWorkbook.Sheets("Table").Range("MyCards").Cells(1, lastEdited).Font.Color = RGB(255, 0, 0) 'red
        Else
            ThisWorkbook.Sheets("Table").Range("MyCards").Cells(1, lastEdited).Font.Color = RGB(0, 0, 0) 'black
        End If
    Else 'pot
        
        'Write on table
        ThisWorkbook.Sheets("Table").Range("Pot").Cells(1, lastEdited).Value = strAux
        
        'Write on aux
        ThisWorkbook.Sheets("Aux").Range("LastEditedPot").Value = lastEdited
        ThisWorkbook.Sheets("Aux").Range("potIDs").Cells(1, lastEdited).Value = cardNo
        
        'Color
        If cardSuit(cardNo) = 2 Or cardSuit(cardNo) = 4 Then
            ThisWorkbook.Sheets("Table").Range("Pot").Cells(1, lastEdited).Font.Color = RGB(255, 0, 0) 'red
        Else
            ThisWorkbook.Sheets("Table").Range("Pot").Cells(1, lastEdited).Font.Color = RGB(0, 0, 0) 'black
        End If
    End If
    
    'Color

End Sub

