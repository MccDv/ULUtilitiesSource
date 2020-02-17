VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Digital Input Speed Test"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtResult 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CheckBox chkArray 
      Caption         =   "Use DInArray"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   2040
      Width           =   915
   End
   Begin VB.TextBox txtPortIndex 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "0"
      Top             =   1020
      Width           =   495
   End
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1260
      TabIndex        =   9
      Text            =   "300"
      Top             =   2220
      Width           =   1335
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Port Input"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Bit Input"
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtFirstBit 
      Height          =   225
      Left            =   2940
      TabIndex        =   6
      Text            =   "0"
      Top             =   1440
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtLastBit 
      Height          =   225
      Left            =   4500
      TabIndex        =   5
      Text            =   "-1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame fraBoard 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.CheckBox chkUlErrors 
         Caption         =   "UL Errors"
         Height          =   195
         Left            =   5640
         TabIndex        =   20
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdFlashLED 
         Caption         =   "FlashLED"
         Height          =   315
         Left            =   6840
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
         Width           =   1815
      End
   End
   Begin VB.Label lblPortsAvailable 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   18
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblPortIndex 
      Caption         =   "Port Index"
      Height          =   195
      Left            =   780
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   15
      Top             =   1080
      Width           =   5835
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Iterations"
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   1980
      Width           =   795
   End
   Begin VB.Label lblFirstBit 
      Caption         =   "First Bit"
      Height          =   195
      Left            =   3480
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblLastBit 
      Caption         =   "Last Bit"
      Height          =   195
      Left            =   5040
      TabIndex        =   12
      Top             =   1440
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
      Width           =   7515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlBoardNum As Long, mlNumBits As Long
Dim mlFirstBit As Long, mlLastBit As Long
Dim mlNumPorts As Long, mlPortIndex As Long
Dim mlPortNum As Long, mlNumArrayPorts As Long
Dim mbDoBits As Boolean, mlTotalBits As Long
Dim msPortList As String
Dim mlErrReporting As Long, mlErrHandling As Long

Private Sub cmbBoard_Click()

   Dim AddBoard As Boolean
   
   If mlNumPorts > 0 Then ConfigureOutputs False
   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board: " & _
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
      GetPortType
   End If
   txtResult.Text = ""
   
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
               TimeOut& = Val(frmRemoteNetDlg.txtTimeout.Text)
               DevsFound& = UpdateDevices(True, HostName, HostPort&, TimeOut&)
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

   mlErrReporting = DONTPRINT
   mlErrHandling = DONTSTOP
   DevsFound& = UpdateDevices(False)
   If Not DevsFound& = 0 Then
      mlPortIndex = 0
      GetPortType
      cmdStart.Enabled = True
   End If
   chkUlErrors_Click

End Sub

Private Function UpdateDevices(ByVal CheckNet As Boolean, _
   Optional HostString As Variant, Optional HostPort As Long, _
   Optional TimeOut As Long) As Long

   Dim devInterface As DaqDeviceInterface
   
   devInterface = USB_IFC + BLUETOOTH_IFC
   If CheckNet Then devInterface = _
      USB_IFC + BLUETOOTH_IFC + ETHERNET_IFC
   DevsInstalled& = GetNumInstacalDevs()
   If IsMissing(HostString) Then
      DevsFound& = DiscoverDevices(devInterface, True)
   Else
      DevsFound& = DiscoverDevices(devInterface, _
         True, HostString, HostPort, TimeOut)
   End If

   cmbBoard.Clear
   For i% = 0 To gnNumBoards - 1
      BoardNum% = gnBoardEnum(i%)
      BoardName$ = GetNameOfBoard(BoardNum%)
      cmbBoard.AddItem BoardName$, i%
   Next i%
   If cmbBoard.ListCount > 0 Then cmbBoard.ListIndex = 0
   UpdateDevices = gnNumBoards
   
End Function

Private Sub Form_Resize()

   If Me.Width < 6795 Then Me.Width = 6795
   lblStatus.Width = Me.Width
   lblStatus.Left = 0
   lblStatus.Top = (Me.Height - 780)
   fraBoard.Left = 0
   fraBoard.Width = Me.Width
   fraBoard.Top = -80
   cmdFlashLED.Left = Me.Width - 1400
   chkUlErrors.Left = Me.Width - 2600
   
End Sub

Private Sub chkUlErrors_Click()

   Dim ulError As Long
   
   If chkUlErrors.Value = 1 Then
      mlErrReporting = PRINTALL
   Else
      mlErrReporting = DONTPRINT
   End If
   
   ulError = cbErrHandling(mlErrReporting, mlErrHandling)
   If ulError <> 0 Then
      ErrMessage$ = GetULError(ulError)
      txtResult.Text = ErrMessage$
   End If
   
End Sub

Private Sub cmdStart_Click()

   Dim NumBits As Long
   Dim spanPorts As Boolean
   
   mlFirstBit = Val(txtFirstBit.Text)
   mlLastBit = Val(txtLastBit.Text)
   GetPortType
   spanPorts = (chkArray.Value = 1)
   NumBits = (mlLastBit - mlFirstBit) + 1
   cmdStart.Enabled = False
   txtResult.Text = ""
   Iterations& = Val(txtRateEstimate.Text)
   
   If mbDoBits Then
      BitPort& = FIRSTPORTA
      If mlPortNum < 10 Then BitPort& = AUXPORT
      ULStat& = cbDBitIn(mlBoardNum, BitPort&, mlFirstBit, BitVal%)
      If ULStat& <> 0 Then
         ErrMessage$ = GetULError(ULStat&)
         txtResult.Text = ErrMessage$
         Exit Sub
      End If
      StartTime! = Timer
      For i& = 1 To Iterations&
         For CurBit& = mlFirstBit To mlLastBit
            ULStat& = cbDBitIn(mlBoardNum, BitPort&, CurBit&, BitVal%)
         Next
      Next
      elapsedTime! = (Timer - StartTime!) / NumBits
   Else
      If spanPorts Then
         ReDim InArray(mlNumArrayPorts - 1) As Long
         LastArrayPort& = mlPortNum + (mlNumArrayPorts - 1)
         ULStat& = cbDInArray(mlBoardNum, mlPortNum, _
            LastArrayPort&, InArray(0))
         If ULStat& <> 0 Then
            ErrMessage$ = GetULError(ULStat&)
            txtResult.Text = ErrMessage$
            Exit Sub
         End If
         StartTime! = Timer
         For i& = 0 To Iterations&
            ULStat& = cbDInArray(mlBoardNum, mlPortNum, _
               LastArrayPort&, InArray(0))
         Next
         elapsedTime! = Timer - StartTime!
      Else
         DataVal% = 0
         ULStat& = cbDIn(mlBoardNum, mlPortNum, DataVal%)
         If ULStat& <> 0 Then
            ErrMessage$ = GetULError(ULStat&)
            txtResult.Text = ErrMessage$
            Exit Sub
         End If
         StartTime! = Timer
         For i& = 0 To Iterations&
            ULStat& = cbDIn(mlBoardNum, mlPortNum, DataVal%)
         Next
         elapsedTime! = Timer - StartTime!
      End If
   End If
   Me.cmdStart.Enabled = True
   inputRate! = 1 / (elapsedTime! / Iterations&)
   FormatString$ = "0.00 Hz"
   divisor! = 1#
   If inputRate! > 999 Then
      FormatString$ = "0.00 kHz"
      divisor! = 1000#
   End If
   txtResult.Text = "Update rate: " & Format(inputRate! / divisor!, FormatString$)
   
End Sub

Private Sub GetPortType()

   Dim OverBit As Boolean
   Dim ArrayEnabled As Boolean
   Dim TrimVal As Integer
   Dim sPortList As String
   
   ArrayEnabled = Not (chkArray.Value = 0)
   If ArrayEnabled Then
      mlTotalBits = 0
      ULStat& = cbGetConfig(BOARDINFO, mlBoardNum, _
         0, BIDINUMDEVS, ConfigVal&)
      For DevNum& = mlPortIndex To ConfigVal& - 1
         ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
            DevNum&, DIDEVTYPE, DevType&)
         ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
            DevNum&, DINUMBITS, NumBits&)
         mlTotalBits = mlTotalBits + NumBits&
         If PrevDevType& > 0 Then
            mlNumArrayPorts = mlNumArrayPorts + 1
            If Not (mlNumBits = NumBits&) Then
               mlNumArrayPorts = 1
               Exit For
            End If
            If Not mbDoBits Then
               sPortList = sPortList & ", " & GetPortString(DevType&)
            Else
               Me.txtLastBit.Text = mlTotalBits - 1
            End If
         Else
            mlNumArrayPorts = 1
            mlNumBits = NumBits&
            PrevDevType& = DevType&
            sPortList = GetPortString(DevType&)
         End If
      Next
      msPortList = sPortList
   Else
      ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
         mlPortIndex, DIDEVTYPE, DevType&)
      If Not ULStat& = 0 Then
         msPortList = "Invalid Port"
         cmdStart.Enabled = False
      Else
         mlPortNum = DevType&
         msPortList = GetPortString(mlPortNum)
         cmdStart.Enabled = True
      End If
      
      ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
         mlPortIndex, DINUMBITS, NumBits&)
      CurLast& = Val(txtLastBit.Text)
      OverBit = Not (CurLast& < NumBits&)
      mlLastBit = NumBits& - 1
      txtLastBit.Text = Format(mlLastBit, "0")
      CurFirst& = Val(txtFirstBit.Text)
      If (CurFirst& > mlLastBit) Then
         mlFirstBit = 0
         txtFirstBit.Text = "0"
      End If
   End If
   
   Offset& = GetBitOffset(mlBoardNum, mlPortIndex)
   mlFirstBit = Val(txtFirstBit.Text) + Offset&
   mlLastBit = Val(txtLastBit.Text) + Offset&
   BitList$ = ""
   If mbDoBits Then
      BitList$ = " (Bit " & mlFirstBit & " to " & mlLastBit & ")"
   End If
   lblPortType.Caption = msPortList & BitList$
   
End Sub

Private Sub chkArray_Click()

   GetPortType

End Sub

Private Sub txtPortIndex_Change()

   mlPortIndex = Val(Me.txtPortIndex.Text)
   GetPortType

End Sub

Private Sub optPort_Click(Index As Integer)

   mbDoBits = optPort(1).Value
   If mbDoBits Then
      chkArray.Caption = "Span ports"
   Else
      chkArray.Caption = "Use DOutArray"
   End If
   txtFirstBit.Visible = mbDoBits
   txtLastBit.Visible = mbDoBits
   lblFirstBit.Visible = mbDoBits
   lblLastBit.Visible = mbDoBits
   GetPortType
   
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
   Else
      txtResult.Text = ""
   End If
   CheckForDigital = ValidBoard
   cmdStart.Enabled = ValidBoard

End Function

Private Sub ConfigureOutputs(SetOutputs As Boolean)

   Dim Direction As Long, PortNum As Long
   
   Direction = DIGITALIN
   If SetOutputs Then Direction = DIGITALOUT
   For i% = 0 To mlNumPorts - 1
      ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
         i%, DIDEVTYPE, PortNum)
      ULStat& = cbDConfigPort(mlBoardNum, PortNum, Direction)
   Next
   
End Sub

