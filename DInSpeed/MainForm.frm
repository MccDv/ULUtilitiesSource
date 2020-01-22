VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Digital Input Speed Test"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkArray 
      Caption         =   "Use DInArray"
      Height          =   255
      Left            =   4860
      TabIndex        =   18
      Top             =   1620
      Width           =   2055
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtPortIndex 
      Height          =   285
      Left            =   420
      TabIndex        =   10
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   "300"
      Top             =   2340
      Width           =   1335
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Port Input"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Bit Input"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtFirstBit 
      Height          =   225
      Left            =   3180
      TabIndex        =   6
      Text            =   "0"
      Top             =   1560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtLastBit 
      Height          =   225
      Left            =   3180
      TabIndex        =   5
      Text            =   "-1"
      Top             =   1860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame fraBoard 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton cmdFlashLED 
         Caption         =   "FlashLED"
         Height          =   315
         Left            =   6360
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
      Begin VB.ComboBox cmbBoard 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "F5 to update, Ctl-F5 for Ethernet, Shift-Ctl-F5 for remote Ethernet"
         Top             =   180
         Width           =   3075
      End
      Begin VB.Label lblBoardNumber 
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3420
         TabIndex        =   3
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.Label lblPortsAvailable 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblResult 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4980
      TabIndex        =   17
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lblPortIndex 
      Caption         =   "Port Index"
      Height          =   195
      Left            =   1140
      TabIndex        =   16
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2580
      TabIndex        =   15
      Top             =   1140
      Width           =   2055
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Iterations"
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label lblFirstBit 
      Caption         =   "First Bit"
      Height          =   195
      Left            =   3720
      TabIndex        =   13
      Top             =   1620
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblLastBit 
      Caption         =   "Last Bit"
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   6915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlBoardNum As Long, mlNumBits As Long
Dim mlFirstBit As Long, mlLastBit As Long
Dim mlNumPorts As Long
Dim mlPortNum As Long, mlNumArrayPorts As Long
Dim mbDoBits As Boolean

Private Sub cmbBoard_Click()

   Dim AddBoard As Boolean
   Dim FuncType As Long
   
   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board Number " & _
         mlBoardNum & " (type 0x" & Prefix$ & pID$ & ")"
      AddBoard = CheckForDigital(mlBoardNum)
      lblPortsAvailable.Caption = "Number of ports available: " & mlNumPorts
   Else
      lblBoardNumber.Caption = "No Boards Installed"
      AddBoard = False
   End If
   Me.lblPortType.Caption = ""
   If AddBoard Then
      Me.cmdStart.Enabled = True
      GetPortType 0
   End If
   
End Sub

Private Sub cmdFlashLED_Click()

   cmdFlashLED.Enabled = False
   ULStat = cbFlashLED(mlBoardNum)
   x% = SaveFunc(Me, FlashLED, ULStat, _
      mlBoardNum, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   cmdFlashLED.Enabled = True
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF5 Then
      If (Shift And vbCtrlMask) Then
         If (Shift And vbShiftMask) = vbShiftMask Then
            frmRemoteNetDlg.Show 1, Me
            If Not (frmRemoteNetDlg.txtHostName.Text = "") Then
               HostName = frmRemoteNetDlg.txtHostName.Text
               HostPort& = Val(frmRemoteNetDlg.txtHostPort.Text)
               Timeout& = Val(frmRemoteNetDlg.txtTimeout.Text)
               DevsFound& = UpdateDevices(True, HostName, HostPort&, Timeout&)
            End If
            Unload frmRemoteNetDlg
         Else
            DevsFound& = UpdateDevices(True)
         End If
      Else
         DevsFound& = UpdateDevices(False)
      End If
   End If
   
End Sub

Private Sub Form_Load()

   DevsFound& = UpdateDevices(False)
   If Not DevsFound& = 0 Then
      GetPortType 0
      cmdStart.Enabled = True
   End If

End Sub

Private Function UpdateDevices(ByVal CheckNet As Boolean, _
   Optional HostString As Variant, Optional HostPort As Long, _
   Optional Timeout As Long) As Long

   Dim devInterface As DaqDeviceInterface
   
   devInterface = USB_IFC + BLUETOOTH_IFC
   If CheckNet Then devInterface = _
      USB_IFC + BLUETOOTH_IFC + ETHERNET_IFC
   DevsInstalled& = GetNumInstacalDevs()
   If IsMissing(HostString) Then
      DevsFound& = DiscoverDevices(devInterface, True)
   Else
      DevsFound& = DiscoverDevices(devInterface, _
         True, HostString, HostPort, Timeout)
   End If

   cmbBoard.Clear
   For i% = 0 To gnNumBoards - 1
      BoardNum% = gnBoardEnum(i%)
      BoardName$ = GetNameOfBoard(BoardNum%)
      cmbBoard.AddItem BoardName$, i%
   Next i%
   If cmbBoard.ListCount > 0 Then cmbBoard.ListIndex = 0
   
End Function

Private Sub Form_Resize()

   If Me.Width < 6795 Then Me.Width = 6795
   lblStatus.Width = Me.Width
   lblStatus.Left = 0
   lblStatus.Top = (Me.Height - 780)
   fraBoard.Left = 0
   fraBoard.Width = Me.Width
   fraBoard.Top = -80
   
End Sub

Private Function CheckForDigital(ByVal BoardNum As Long) As Boolean
   
   Dim ValidBoard As Boolean
   Dim ReportError As Long, HandleError As Long
   Dim DefaultPort As Long, DefaultNumBits As Long
   Dim FirstBit As Long
   ReportError = DONTPRINT
   HandleError = DONTSTOP
   
   ValidBoard = False
   SetDigitalIODefaults ReportError, HandleError
   Dim ChannelType As Long
   ChannelType = PORTIN
   NumDIChans = FindPortsOfType(BoardNum, ChannelType, _
      PROGPORT, DefaultPort, DefaultNumBits, FirstBit)
   mlNumPorts = NumDIChans
   If Not (NumDIChans = 0) Then
      ValidBoard = True
      mlPortNum = DefaultPort
      mnResolution = DefaultNumBits
   End If
   cmdStart.Enabled = ValidBoard
   CheckForDigital = ValidBoard

End Function

Private Sub cmdStart_Click()

   Dim FirstBit As Long, LastBit As Long
   
   Offset& = GetBitOffset()
   FirstBit = Val(txtFirstBit.Text) + Offset&
   LastBit = Val(txtLastBit.Text) + Offset&
   Me.cmdStart.Enabled = False
   Iterations& = Val(txtRateEstimate.Text)
   
   ULStat& = cbDConfigPort(mlBoardNum, mlPortNum, DIGITALIN)
   PortNum& = mlPortNum
   If mlPortNum > 10 Then PortNum& = FIRSTPORTA
   StartTime! = Timer
   If mbDoBits Then
      For i& = 1 To Iterations&
         For CurBit& = FirstBit To LastBit
            ULStat& = cbDBitIn(mlBoardNum, PortNum&, CurBit&, BitVal%)
         Next
      Next
   Else
      DataVal% = 0
      ULStat& = cbDIn(mlBoardNum, mlPortNum, DataVal%)
      For i& = 1 To Iterations&
         ULStat& = cbDIn(mlBoardNum, mlPortNum, DataVal%)
      Next
   End If
   elapsedTime! = Timer - StartTime!
   Me.cmdStart.Enabled = True
   inputRate! = 1 / (elapsedTime! / Iterations&)
   FormatString$ = "0.00 Hz"
   divisor! = 1#
   If inputRate! > 999 Then
      FormatString$ = "0.00 kHz"
      divisor! = 1000#
   End If
   Me.lblResult.Caption = Format(inputRate! / divisor!, FormatString$)
   
End Sub

Function GetBitOffset() As Long

   Select Case mlPortNum
      Case 0, 10
         Offset& = 0
      Case Is > 10
         Offset& = 8
         For CurPort& = 12 To mlPortNum
            Select Case CurPort&
               Case 12, 13, 16, 17, 20, 21, 24, 25
                  Offset& = Offset& + 4
               Case 14, 15, 18, 19, 22, 23, 26, 27
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

Private Sub optPort_Click(Index As Integer)

   mbDoBits = optPort(1).Value
   txtFirstBit.Visible = mbDoBits
   txtLastBit.Visible = mbDoBits
   lblFirstBit.Visible = mbDoBits
   lblLastBit.Visible = mbDoBits
   GetPortType 0
   
End Sub

Private Sub GetPortType(ByVal PortIndex As Long)

   Dim OverBit As Boolean
   
   ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
   PortIndex, DIDEVTYPE, DevType&)
   If Not ULStat& = 0 Then
      Me.lblPortType.Caption = "Invalid Port"
      cmdStart.Enabled = False
   Else
      mlPortNum = DevType&
      Me.lblPortType.Caption = GetPortString(mlPortNum)
      cmdStart.Enabled = True
   End If
   
   ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
      PortIndex, DINUMBITS, NumBits&)
   CurLast& = Val(txtLastBit.Text)
   OverBit = Not (CurLast& < NumBits&)
   If (CurLast& < 0) Or OverBit Then
      mlLastBit = NumBits& - 1
      txtLastBit.Text = Format(mlLastBit, "0")
   End If
   CurFirst& = Val(txtFirstBit.Text)
   If (CurFirst& > mlLastBit) Then
      mlFirstBit = 0
      txtFirstBit.Text = "0"
   End If
   
End Sub

Private Sub ArrayLoop()
   
   Dim OutArrayLow(1) As Long
   Dim OutArrayHigh(1) As Long
   
   Me.cmdStart.Enabled = False
   Iterations& = Val(txtRateEstimate.Text) * 5
   
   ULStat& = cbDConfigPort(mlBoardNum, mlPortNum, DIGITALIN)
   If Me.chkArray.Value = 1 Then
      OutArrayLow(0) = 0
      OutArrayLow(1) = 0
      OutArrayHigh(0) = (2 ^ mlNumBits) - 1
      OutArrayHigh(1) = (2 ^ mlNumBits) - 1
      FirstArrayPort& = mlPortNum + 1
      LastArrayPort& = mlPortNum + (mlNumArrayPorts - 1)
      For PortNum& = FirstArrayPort& To LastArrayPort&
         ULStat& = cbDConfigPort(mlBoardNum, _
            PortNum&, DIGITALOUT)
      Next
      ULStat& = cbDOutArray(mlBoardNum, mlPortNum, _
         LastArrayPort&, OutArrayLow(0))
      For i& = 0 To Iterations&
         ULStat& = cbDOutArray(mlBoardNum, mlPortNum, _
            LastArrayPort&, OutArrayHigh(0))
         ULStat& = cbDOutArray(mlBoardNum, mlPortNum, _
            LastArrayPort&, OutArrayLow(0))
      Next
   End If
   
End Sub

Private Sub chkArray_Click()

   Dim PortEnabled As Boolean
   
   PortEnabled = Not (chkArray.Value = 0)
   txtPortIndex.Text = "0"
   txtPortIndex.Enabled = (chkArray.Value = 0)
   If PortEnabled Then
      ULStat& = cbGetConfig(BOARDINFO, mlBoardNum, _
         0, BIDINUMDEVS, ConfigVal&)
      For DevNum& = 0 To ConfigVal& - 1
         ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
            DevNum&, DIDEVTYPE, DevType&)
         ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
            DevNum&, DINUMBITS, NumBits&)
         If PrevDevType& > 0 Then
            mlNumArrayPorts = mlNumArrayPorts + 1
            If Not (mlNumBits = NumBits&) Then
               mlNumArrayPorts = 1
               Exit For
            End If
         Else
            mlNumArrayPorts = 1
            mlNumBits = NumBits&
            PrevDevType& = DevType&
         End If
      Next
      PortTypeString$ = "Port Type "
      If mlNumArrayPorts > 1 Then
         For PortNum& = mlPortNum To mlPortNum + (mlNumArrayPorts - 1)
            PortTypeString$ = PortTypeString$ & _
               Format(PortNum&, "0") & ", "
         Next
      End If
      PortTypeString$ = Left(PortTypeString$, Len(PortTypeString$) - 2)
   Else
      PortTypeString$ = "Port Type " & _
         Format(mlPortNum, "0")
   End If
   Me.lblPortType.Caption = PortTypeString$

End Sub

Private Sub txtPortIndex_Change()

   PortIndex& = Val(Me.txtPortIndex.Text)
   GetPortType PortIndex&

End Sub
