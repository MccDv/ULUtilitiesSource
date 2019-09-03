Attribute VB_Name = "GPIBInterface"
Public Function InitGPIB() As Boolean

   On Error GoTo MissingGPIB
   Dim InitSuccess As Boolean

   InitSuccess = False
   SendIFC (0)   'make sure there's a library installed
   If (ibsta < 0) Then
      'library exists, but no board installed or configured
      'see if MCC signal controllers are installed
      InitGPIB = InitSuccess
      Exit Function
   End If

   InitSuccess = True
   InitGPIB = InitSuccess
   Exit Function

MissingGPIB:
   InitSuccess = False
   InitGPIB = InitSuccess
   Exit Function

End Function

Function GetAddressList(ByRef ListOfAddresses() As Integer, _
   ByVal AddrRange As Integer) As Integer

   Dim i As Integer, AddsFound As Integer
   Dim ErrorMessage As String
   ReDim addrs(AddrRange) As Integer
   ReDim results(AddrRange) As Integer
   
   For i = 0 To AddrRange
      addrs(i) = i + 2
      results(i) = -1
   Next
   addrs(AddrRange) = NOADDR
   
   FindLstn 0, addrs(), results(), AddrRange ' + 1
   If (ibsta And EERR) = EERR Then
      StatString = ParseStatus()
      ErrorString = GetErrorMessage(ErrorMessage)
      MsgBox StatString, vbOKOnly, ErrorString
      GetAddressList = 0
   Else
      For i = 0 To AddrRange
         If Not (results(i) = -1) Then
            ReDim Preserve ListOfAddresses(AddsFound)
            ListOfAddresses(AddsFound) = results(i)
            AddsFound = AddsFound + 1
         End If
      Next
      GetAddressList = AddsFound
   End If

   
End Function

Public Sub GPIBWrite(ByVal Address As Integer, ByVal Command As String)

   Send 0, Address, Command, DABend
   If (ibsta And EERR) = EERR Then
      Dim ErrorMessage As String
      StatString = ParseStatus()
      ErrorString = GetErrorMessage(ErrorMessage)
      MsgBox StatString, vbOKOnly, ErrorString
   End If
   
End Sub

Public Sub GPIBRead(ByVal Address As Integer, ByRef DataRead As String, ByVal BufSize As Integer)

   Dim Location As Long
   
   DataRead = Space(BufSize)
   Receive 0, Address, DataRead, STOPend
   If (ibsta And EERR) = EERR Then
      Dim ErrorMessage As String
      StatString = ParseStatus()
      ErrorString = GetErrorMessage(ErrorMessage)
      MsgBox StatString, vbOKOnly, ErrorString
   End If
   Location = InStr(1, DataRead, Chr(10)) - 1
   If Location > 0 Then
      DataRead = Left(DataRead, Location)
      Location = InStr(1, DataRead, Chr(13)) - 1
      If Location = 0 Then
         DataRead = ""
      ElseIf Location > 0 Then
         DataRead = Left(DataRead, Location)
      End If
   Else
      DataRead = Trim(DataRead)
   End If
   
End Sub

Public Function ParseStatus() As String

   Dim StatBit As Integer, StatWeight As Long
   Dim StatString As String
   
   For StatBit = 0 To 15
      StatWeight = 2 ^ StatBit
      If (ibsta And StatWeight) = StatWeight Then
         StatString = StatString & Choose(StatBit + 1, "DCAS, ", _
            "DTAS, ", "LACS, ", "TACS, ", "AATN, ", "CIC, ", _
            "RREM, ", "LOK, ", "CMPL, ", "EEVENT, ", "SPOLL, ", _
            "RQS, ", "SRQI, ", "EEND, ", "TIMO, ", "EERR, ")
      End If
   Next
   StatString = Left(StatString, Len(StatString) - 2)
   ParseStatus = StatString
   
End Function

Public Function GetErrorMessage(ByRef Message As String) As String

   Select Case iberr
      Case 0
         GetErrorMessage = "EDVR"
         Message = ""
      Case 1
         GetErrorMessage = "ECIC"
         Message = ""
      Case 2
         GetErrorMessage = "ENOL"
         Message = ""
      Case 3
         GetErrorMessage = "EADR"
         Message = ""
      Case 4
         GetErrorMessage = "EARG"
         Message = ""
      Case 5
         GetErrorMessage = "ESAC"
         Message = ""
      Case 6
         GetErrorMessage = "EABO"
         Message = ""
      Case 7
         GetErrorMessage = "ENEB"
         Message = ""
      Case 10
         GetErrorMessage = "EOIP"
         Message = ""
      Case 11
         GetErrorMessage = "ECAP"
         Message = ""
      Case 12
         GetErrorMessage = "EFSO"
         Message = ""
      Case 14
         GetErrorMessage = "EBUS"
         Message = ""
      Case 15
         GetErrorMessage = "ESTB"
         Message = ""
      Case 16
         GetErrorMessage = "ESRQ"
         Message = ""
      Case 14
         GetErrorMessage = "ETAB"
         Message = ""
      Case 15
         GetErrorMessage = "EHDL"
         Message = ""
      Case 16
         GetErrorMessage = ""
         Message = ""
      Case 20
         GetErrorMessage = ""
         Message = ""
      Case 23
         GetErrorMessage = ""
         Message = ""
   End Select
   
End Function
