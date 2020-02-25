Attribute VB_Name = "TemplateDefs"
Global gnBoardEnum() As Integer, gnNumBoards As Integer
'library used (gnLibType values)
Global Const INVALIDLIB = -1
Global Const UNILIB = 0
Global Const NETLIB = 1
Global Const MSGLIB = 2

Global Const MAIN_FORM = 0
Global Const ANALOG_IN = 1
Global Const ANALOG_OUT = 2
Global Const ANALOG_IO = 1 'in and out are actually the same form
Global Const DIGITAL_IN = 3
Global Const DIGITAL_OUT = 4
Global Const DIGITAL_IO = 3 'in and out are actually the same form
Global Const COUNTER = 5
Global Const UTILITIES = 6
Global Const Config = 7
Global Const GPIB_CTL = 8
Global Const LOGFUNC = 9

Dim mfNoForm As Form

Public Function GetNumInstacalDevs() As Long
   
   'Get number of boards installed in Instacal
   InfoType& = GLOBALINFO
   DevNum& = 0
   ConfigItem& = GINUMBOARDS
   'to do - this beeps on startup sometimes (USB-QUAD08 installed)
   ULStat = cbGetConfig(InfoType&, 0, DevNum&, ConfigItem&, ValConfig&)
   If ULStat = CFGFILENOTFOUND Then
      GetNumInstalled = 0
      Exit Function
   End If
   ConfigVal& = ValConfig&
   If SaveFunc(mfNoForm, GetConfig, ULStat, _
      InfoType&, 0, DevNum&, ConfigItem&, ConfigVal&, _
      A6, A7, A8, A9, A10, A11, 0) Then Exit Function
   
   NumBoards& = ConfigVal&
      'this number is actually MAX_NUM_BOARDS (should be GI_NUMBOARDS)
      'the following For/Next compensates for this
   For board& = 0 To NumBoards& - 1
      InfoType& = BOARDINFO
      DevNum& = 0
      ConfigItem& = BIBOARDTYPE
      ULStat = cbGetConfig(InfoType&, board&, DevNum&, ConfigItem&, ValConfig&)
      ConfigVal& = ValConfig&
      If SaveFunc(mfNoForm, GetConfig, ULStat, InfoType&, _
         board&, DevNum&, ConfigItem&, ConfigVal&, A6, A7, _
         A8, A9, A10, A11, 0) Then Exit Function
      If ConfigVal& <> 0 Then
         'determine boards relative position
         '(second board installed could be board 5, etc.)
         ReDim Preserve gnBoardEnum(Installed%)
         gnBoardEnum(Installed%) = board&
         Installed% = Installed% + 1
         
         'check if this board has an associated memory board
         InfoType& = BOARDINFO
         DevNum& = 0
         ConfigItem& = BIDTBOARD
         ULStat = cbGetConfig(InfoType&, board&, DevNum&, ConfigItem&, ValConfig&)
         ConfigVal& = ValConfig&
         If SaveFunc(mfNoForm, GetConfig, ULStat, InfoType&, _
            board&, DevNum&, ConfigItem&, ConfigVal&, A6, A7, _
            A8, A9, A10, A11, 0) Then Exit Function
         If ConfigVal& > -1 Then
            'determine boards relative position
            '(second board installed could be board 5, etc.)
            'and if it has already been detected as its own board number (UL16)
            ReDim Preserve gnBoardEnum(Installed%)
            gnBoardEnum(Installed%) = ConfigVal&
            Installed% = Installed% + 1
         End If
      End If
   Next board&
   gnNumBoards = Installed%
   GetNumInstacalDevs = Installed%

End Function

Public Function GetNameOfBoard(ByVal BoardNum As Integer, _
   Optional LibType As Integer = UNILIB) As String

   'redirects based on optional parameter
   Select Case LibType
      Case UNILIB
         GetNameOfBoard = ""
         BoardName$ = Space$(BOARDNAMELEN)
         ULStat = cbGetBoardName(BoardNum, BoardName$)
         If Len(BoardName$) Then
            'Drop the space characters
            BoardName$ = RTrim$(BoardName$)
            StringSize% = Len(BoardName$)
            'lop off the null
            If StringSize% > 0 Then BoardName$ = Left$(BoardName$, StringSize% - 1)
            GetNameOfBoard = BoardName$
         Else
            GetNameOfBoard = ""
         End If
         If SaveFunc(mfNoForm, GetBoardName, ULStat, _
            BoardNum, BoardName$, A3, A4, A5, A6, A7, _
            A8, A9, A10, A11, 0) Then Exit Function
      Case NETLIB
         BoardName$ = gnBoardEnum(BoardNum)
      Case MSGLIB
         'BoardName$ = GetNameOfMsgBoard(BoardNum)
         GetNameOfBoard = BoardName$
         Exit Function
   End Select

End Function

Public Function SaveFunc(CallingForm As Form, FuncID _
   As Integer, FuncStat, A1, A2, A3, A4, A5, A6, A7, _
   A8, A9, A10, A11, AuxHandle) As Integer

   Dim ErrOccurred As Boolean
   Dim FuncString As String, FuncName As String
   Dim ErrMsg As String * 100
   
   FuncString = GetFunctionName(FuncID)
   FuncName = Left(FuncString, Len(FuncString) - 2)
   'More$ = GetFunctionString(FuncID)
   ArgStr$ = "("
   For a% = 1 To 11
      ArgVal = Choose(a%, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11)
      If Not IsEmpty(ArgVal) Then
         If VarType(ArgVal) = vbString Then
            ArgStr$ = ArgStr$ & ArgVal & ", "
         Else
            ArgStr$ = ArgStr$ & Format(ArgVal, "0") & ", "
         End If
      End If
   Next
   ArgLen& = Len(ArgStr$)
   If ArgLen& > 1 Then ArgStr$ = Left(ArgStr$, ArgLen& - 2)
   ArgStr$ = ArgStr$ & ")"
   ErrOccurred = (FuncStat <> 0)
   ULStat = cbGetErrMsg(FuncStat, ErrMsg)
   Loca& = InStr(1, ErrMsg, Chr(0))
   ErrorMessage$ = Left(ErrMsg, Loca& - 1)
   If Not (CallingForm Is Nothing) Then
      If ErrOccurred Then
         CallingForm.lblStatus.ForeColor = &HFF
         CallingForm.lblStatus.Caption = ErrorMessage$ & "  Function = " & FuncName & ArgStr$
      Else
         CallingForm.lblStatus.ForeColor = &HFF0000
         CallingForm.lblStatus.Caption = "Function = " & FuncName & ArgStr$
      End If
   Else
      If ErrOccurred Then
         MsgBox ErrorMessage$ & "  Function = " & _
            FuncString, vbExclamation, "Error Occurred"
      End If
   End If
   SaveFunc = ErrOccurred
   
End Function

Public Sub PrintMain(MainString As String)

   frmMain.lblStatus.ForeColor = &HFF0000
   frmMain.lblStatus.Caption = MainString
   
End Sub


Function GetConfigString573(ByVal InfoType&, ByVal _
   BoardNum&, ByVal DevNum&, ByVal ConfigItem&, ByRef _
   ConfigVal$, ByRef ConfigLen&) As Long

   GetConfigString573 = cbGetConfigString(InfoType&, _
      BoardNum&, DevNum&, ConfigItem&, ConfigVal$, ConfigLen&)

End Function

Public Function GetDiscoverOption() As Boolean

   'stub that sets removal of undiscovered
   'devices to enabled (optional in UT)
   GetDiscoverOption = True
   
End Function

Public Sub RemoveDiscoveryForm()

   'has no function in this application
   'stub to allow for use of UT module
   
End Sub
