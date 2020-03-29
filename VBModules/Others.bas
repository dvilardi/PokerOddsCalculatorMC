Attribute VB_Name = "Others"
Sub Quicksort(vArray() As Long, arrLbound As Long, arrUbound As Long)
'Sorts a one-dimensional VBA array from smallest to largest
'using a very fast quicksort algorithm variant.
Dim pivotVal As Variant
Dim vSwap    As Variant
Dim tmpLow   As Long
Dim tmpHi    As Long

tmpLow = arrLbound
tmpHi = arrUbound
pivotVal = vArray((arrLbound + arrUbound) \ 2)
 
While (tmpLow <= tmpHi) 'divide
   While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
      tmpLow = tmpLow + 1
   Wend
  
   While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
      tmpHi = tmpHi - 1
   Wend
 
   If (tmpLow <= tmpHi) Then
      vSwap = vArray(tmpLow)
      vArray(tmpLow) = vArray(tmpHi)
      vArray(tmpHi) = vSwap
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If
Wend
 
  If (arrLbound < tmpHi) Then Quicksort vArray, arrLbound, tmpHi 'conquer
  If (tmpLow < arrUbound) Then Quicksort vArray, tmpLow, arrUbound 'conquer
End Sub

Sub scoreSort(vArray() As Double, arrLbound As Long, arrUbound As Long)
'Sorts a one-dimensional VBA array from smallest to largest
'using a very fast quicksort algorithm variant.
Dim pivotVal As Variant
Dim vSwap    As Variant
Dim tmpLow   As Long
Dim tmpHi    As Long

tmpLow = arrLbound
tmpHi = arrUbound
pivotVal = vArray((arrLbound + arrUbound) \ 2)
 
While (tmpLow <= tmpHi) 'divide
   While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)
      tmpLow = tmpLow + 1
   Wend
  
   While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)
      tmpHi = tmpHi - 1
   Wend
 
   If (tmpLow <= tmpHi) Then
      vSwap = vArray(tmpLow)
      vArray(tmpLow) = vArray(tmpHi)
      vArray(tmpHi) = vSwap
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If
Wend
 
  If (arrLbound < tmpHi) Then scoreSort vArray, arrLbound, tmpHi 'conquer
  If (tmpLow < arrUbound) Then scoreSort vArray, tmpLow, arrUbound 'conquer
End Sub
'-------------------------------------------------------------------------------------------------'


Sub generateHandRanking()
    
    Dim shAux As Worksheet
    Set shAux = ThisWorkbook.Sheets("Aux")
    rowAux = 1
    colAux = 11
    
    For i = 1 To 52
        For j = 1 To 52
            rowAux = rowAux + 1
            shAux.Cells(rowAux, colAux).Value = i
            shAux.Cells(rowAux, colAux + 1).Value = j
            DoEvents
        Next j
    Next i
    
    MsgBox "done"
End Sub
