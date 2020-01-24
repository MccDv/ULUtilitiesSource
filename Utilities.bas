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

Function GetBitOffset(PortNum As Long) As Long

   Select Case PortNum
      Case 0, 10
         Offset& = 0
      Case Is > 10
         Offset& = 8
         For CurPort& = 11 To PortNum - 1
            Select Case CurPort&
               Case 12, 13, 16, 17, 20, 21, 24, 25
                  Offset& = Offset& + 4
               Case 11, 14, 15, 18, 19, 22, 23, 26, 27
                  Offset& = Offset& + 8
               Case 28, 29, 32, 33, 36, 37, 40, 41
                  Offset& = Offset& + 4
               Case 30, 31, 34, 35, 38, 39
                  Offset& = Offset& + 8
            End Select
         Next
   End Select
   GetBitOffset = Offset&
   
End Function


