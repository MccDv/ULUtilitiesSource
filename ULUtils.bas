Attribute VB_Name = "ULUtils"
Public Function ParseFWVersion(ByVal DecimalFWVer As Long) As String

   Dim FWVersion As String
   
   FWVersion = "unknown"
   If DecimalFWVer > 0 Then
      FWVersion = Hex(DecimalFWVer)
      Select Case Len(FWVersion)
         Case 1
            FWVersion = "0.0" & FWVersion
         Case 2
            FWVersion = "0." & FWVersion
         Case 3
            FWVersion = Left(FWVersion, 1) & _
               "." & Right(FWVersion, 2)
         Case 4
            FWVersion = Left(FWVersion, 2) & _
               "." & Right(FWVersion, 2)
      End Select
   End If
   ParseFWVersion = FWVersion

End Function

Public Function StripString(ByVal RawString As String) As String

   Dim Stripped As String
   Dim Location As Long
   
   Location = InStr(1, RawString, Chr(0))
   If Location > 1 Then Stripped = _
      Left(RawString, Location - 1)
   Location = InStr(1, RawString, Chr(10))
   If Location > 1 Then Stripped = _
      Left(RawString, Location - 1)
      Location = InStr(1, RawString, Chr(13))
   If Location > 1 Then Stripped = _
      Left(RawString, Location - 1)
   StripString = Trim(Stripped)
   
End Function
