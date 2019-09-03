Attribute VB_Name = "Utility"
Public Function ConvStringToInt(ByRef NumericString As String) As Integer

   If Len(NumericString) > 2 Then
      TypeID$ = Left(NumericString, 2)
   End If
   Select Case TypeID$
      Case "0x"
         StringVal% = Val("&H" & Mid(NumericString, 3))
      Case "&H"
         StringVal% = Val(NumericString)
      Case Else
         StringVal% = Val(NumericString)
   End Select
   ConvStringToInt = StringVal%
   
End Function

Public Function NullTermByteToString(ByRef ByteArray As String) As String

   ConvName$ = StrConv(ByteArray, vbUnicode)
   NewName$ = ""
   NameLen& = Len(ConvName$)
   TermLoc& = InStr(1, ConvName$, Chr(0)) - 1
   If (NameLen& > 1) And (TermLoc& < NameLen&) Then _
      NewName$ = Left(ConvName$, TermLoc&)
   NullTermByteToString = NewName$

End Function

Public Function FindInString(StringToSearch As String, CharToFind As String, Locations As Variant) As Long
   'returns number of occurrances of CharToFind in StringToSearch
   'or returns -1 if no occurrance is found
   'returns location of all occurrances in Locations array variant
   Dim LocsFound() As Long
   
   Do
      CurLoc& = RetVal& + 1
      RetVal& = InStr(CurLoc&, StringToSearch, CharToFind)
      If Not RetVal& = 0 Then
         ReDim Preserve LocsFound(NumLocs&)
         LocsFound(NumLocs&) = RetVal&
         NumLocs& = NumLocs& + 1
      End If
   Loop While RetVal& > 0
   Locations = LocsFound()
   FindInString = NumLocs& - 1

End Function

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

