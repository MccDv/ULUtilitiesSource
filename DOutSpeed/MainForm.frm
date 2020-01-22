VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Form"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   "300"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2100
      Width           =   1215
   End
   Begin VB.CheckBox chkArray 
      Caption         =   "Use DOutArray"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtPortIndex 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   495
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
   Begin VB.Label lblResult 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4920
      TabIndex        =   13
      Top             =   2220
      Width           =   2055
   End
   Begin VB.Label lblPortsAvailable 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   780
      Width           =   3015
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Rate Estimate"
      Height          =   195
      Left            =   3420
      TabIndex        =   11
      Top             =   2220
      Width           =   1275
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2460
      TabIndex        =   10
      Top             =   1260
      Width           =   2835
   End
   Begin VB.Label lblPortIndex 
      Caption         =   "Port Index"
      Height          =   195
      Left            =   900
      TabIndex        =   6
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2820
      Width           =   6915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlBoardNum As Long, mlNumBits As Long
Dim mlNumPorts As Long
Dim mlPortNum As Long, mlNumArrayPorts As Long

Private Sub cmbBoard_Click()

   Dim AddBoard As Boolean
   Dim previousBoard As Long
   
   previousBoard = mlBoardNum
   If mlNumPorts > 0 Then ConfigureOutputs False
   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, configVal&)
      pID$ = Hex(configVal&)
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
      If mlNumPorts > 0 Then ConfigureOutputs True
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

   ULStat& = cbErrHandling(DONTPRINT, DONTSTOP)
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
   UpdateDevices = gnNumBoards
   
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If mlNumPorts > 0 Then ConfigureOutputs False
   
End Sub

Private Sub Form_Resize()

   If Me.Width < 6795 Then Me.Width = 6795
   lblStatus.Width = Me.Width
   lblStatus.Left = 0
   lblStatus.Top = (Me.Height - 780)
   fraBoard.Left = 0
   fraBoard.Width = Me.Width
   fraBoard.Top = -80
   
End Sub

Private Sub chkArray_Click()

   Dim PortEnabled As Boolean
   Dim TrimVal As Integer
   
   PortEnabled = Not (chkArray.Value = 0)
   txtPortIndex.Text = "0"
   txtPortIndex.Enabled = (chkArray.Value = 0)
   If PortEnabled Then
      ULStat& = cbGetConfig(BOARDINFO, mlBoardNum, _
         0, BIDINUMDEVS, configVal&)
      For DevNum& = 0 To configVal& - 1
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
      TrimVal = 0
      PortTypeString$ = "Port Type "
      If mlNumArrayPorts > 1 Then
         For PortNum& = mlPortNum To mlPortNum + (mlNumArrayPorts - 1)
            PortTypeString$ = PortTypeString$ & _
               Format(PortNum&, "0") & ", "
         Next
         TrimVal = 2
      Else
         PortTypeString$ = PortTypeString$ & _
            Format(mlPortNum, "0")
      End If
      PortTypeString$ = Left(PortTypeString$, Len(PortTypeString$) - TrimVal)
   Else
      PortTypeString$ = "Port Type " & _
         Format(mlPortNum, "0")
   End If
   Me.lblPortType.Caption = PortTypeString$

End Sub

Private Sub cmdStart_Click()

   Dim OutArrayLow(1) As Long
   Dim OutArrayHigh(1) As Long
   
   Me.cmdStart.Enabled = False
   Iterations& = Val(txtRateEstimate.Text)
   
   ULStat& = cbDConfigPort(mlBoardNum, mlPortNum, DIGITALOUT)
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
      StartTime! = Timer
      For i& = 0 To Iterations&
         ULStat& = cbDOutArray(mlBoardNum, mlPortNum, _
            LastArrayPort&, OutArrayHigh(0))
         ULStat& = cbDOutArray(mlBoardNum, mlPortNum, _
            LastArrayPort&, OutArrayLow(0))
      Next
   Else
      DataVal% = 0
      
      ULStat& = cbDOut(mlBoardNum, mlPortNum, DataVal%)
      StartTime! = Timer
      For i& = 0 To Iterations&
         ULStat& = cbDOut(mlBoardNum, mlPortNum, 255)
         ULStat& = cbDOut(mlBoardNum, mlPortNum, 0)
      Next
   End If
   elapsedTime! = Timer - StartTime!
   Me.cmdStart.Enabled = True
   outputRate! = 1 / (elapsedTime! / Iterations&)
   FormatString$ = "0.00 Hz"
   divisor! = 1#
   If outputRate! > 999 Then
      FormatString$ = "0.00 kHz"
      divisor! = 1000#
   End If
   Me.lblResult.Caption = Format(outputRate! / divisor!, FormatString$)
   
End Sub

Private Sub GetPortType(ByVal PortIndex As Long)

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
   
End Sub

Private Sub txtPortIndex_Change()

   PortIndex& = Val(Me.txtPortIndex.Text)
   GetPortType PortIndex&
   
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
   ChannelType = PORTOUT
   NumDIChans = FindPortsOfType(BoardNum, ChannelType, _
      PROGPORT, DefaultPort, DefaultNumBits, FirstBit)
   mlNumPorts = NumDIChans
   If Not (NumDIChans = 0) Then
      ValidBoard = True
      mlPortNum = DefaultPort
      mnResolution = DefaultNumBits
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
