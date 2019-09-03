Attribute VB_Name = "Sorting"
Public Sub QuickSortVariants(vArray As Variant, inLow As Long, inHi As Long)
      
   Dim pivot   As Variant
   Dim tmpSwap As Variant
   Dim tmpLow  As Long
   Dim tmpHi   As Long
    
   tmpLow = inLow
   tmpHi = inHi
    
   pivot = vArray((inLow + inHi) \ 2)
  
   While (tmpLow <= tmpHi)
  
      While (vArray(tmpLow) < pivot And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      
      While (pivot < vArray(tmpHi) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend

      If (tmpLow <= tmpHi) Then
         tmpSwap = vArray(tmpLow)
         vArray(tmpLow) = vArray(tmpHi)
         vArray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   
   Wend
  
   If (inLow < tmpHi) Then QuickSortVariants vArray, inLow, tmpHi
   If (tmpLow < inHi) Then QuickSortVariants vArray, tmpLow, inHi
  
End Sub

