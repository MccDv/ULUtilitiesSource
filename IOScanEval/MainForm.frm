VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Form"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRateDecrementer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3780
      Top             =   1680
   End
   Begin VB.TextBox txtRate 
      Height          =   315
      Left            =   1920
      TabIndex        =   22
      Text            =   "10000000"
      Top             =   3420
      Width           =   1155
   End
   Begin VB.Frame fraTestType 
      Caption         =   "Select Test"
      Height          =   915
      Left            =   5580
      TabIndex        =   19
      Top             =   720
      Width           =   2595
      Begin VB.OptionButton optTestSelected 
         Caption         =   "Source rate evaluation"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optTestSelected 
         Caption         =   "Max rate evaluation"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.TextBox txtPortIndex 
      Height          =   285
      Left            =   180
      TabIndex        =   15
      Text            =   "0"
      Top             =   2100
      Width           =   495
   End
   Begin VB.CheckBox chkArray 
      Caption         =   "Span ports"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Timer tmrRateGenerator 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4740
      Top             =   1680
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   915
   End
   Begin VB.Frame fraScanType 
      Caption         =   "Scan Type"
      Height          =   915
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5355
      Begin VB.OptionButton optScanType 
         Caption         =   "Counter Input"
         Height          =   195
         Index           =   4
         Left            =   3660
         TabIndex        =   11
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optScanType 
         Caption         =   "Digital Output"
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   10
         Top             =   600
         Width           =   1635
      End
      Begin VB.OptionButton optScanType 
         Caption         =   "Digital Input"
         Height          =   195
         Index           =   2
         Left            =   1980
         TabIndex        =   9
         Top             =   300
         Width           =   1635
      End
      Begin VB.OptionButton optScanType 
         Caption         =   "Analog Output"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1635
      End
      Begin VB.OptionButton optScanType 
         Caption         =   "Analog Input"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   300
         Width           =   1635
      End
   End
   Begin VB.Frame fraBoard 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8115
      Begin VB.CheckBox chkUlErrors 
         Caption         =   "UL Errors"
         Height          =   195
         Left            =   5760
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdFlashLED 
         Caption         =   "FlashLED"
         Height          =   315
         Left            =   6960
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
         Width           =   2115
      End
   End
   Begin VB.Label lblSourceRate 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5460
      TabIndex        =   24
      Top             =   3480
      Width           =   2460
   End
   Begin VB.Label lblActualRate 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3360
      TabIndex        =   23
      Top             =   3480
      Width           =   1395
   End
   Begin VB.Label lblPortIndex 
      Caption         =   "Port Index"
      Height          =   195
      Left            =   840
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1860
      TabIndex        =   17
      Top             =   2160
      Width           =   5835
   End
   Begin VB.Label lblPortsAvailable 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblCurRate 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1980
      TabIndex        =   13
      Top             =   3060
      Width           =   3375
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4980
      Width           =   6915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnAIResolution As Integer, mnAOResolution As Integer
Dim mnDIResolution As Integer, mnDOResolution As Integer
Dim mnCtrResolution As Integer
Dim mnDataValLow, mnDataValHigh As Integer
Dim mlBoardNum As Long, mlFuncType As Long, mlRange As Long
Dim mlErrReporting As Long, mlErrHandling As Long
Dim mlDataBuffer() As Integer
Dim mlDataBuffer32() As Long
Dim mlMemHandle As Long, mlADRange As Long, mlDARange As Long
Dim mlNumSamples As Long
Dim mlTestRate As Long, mlRate As Long
Dim mlFailRate As Long, mlPassRate As Long
Dim mlPortNum As Long, mlPortIndex As Long, msPortList As String
Dim mlNumArrayPorts As Long, mlNumPorts As Long
Dim mbManualRate As Boolean

Private Sub cmbBoard_Click()

   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board: " & _
         mlBoardNum & " (type 0x" & Prefix$ & pID$ & ")"
   Else
      lblBoardNumber.Caption = "No Boards Installed"
   End If
   CheckDeviceSubsystems
   
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

   Me.Caption = App.EXEName & " Test Application"
   DevsFound& = UpdateDevices(False)
   CheckDeviceSubsystems

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
      If HostString = "" Then Exit Function
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

Sub CheckDeviceSubsystems()

   Dim DefaultTrig As Long, Resolution As Long
   Dim CBRange As Long, LowChan As Long
   Dim HighChan As Long, ChannelType As Long
   Dim NumChans As Long, NumBits As Long, FirstBit As Long
   Dim TypeSelected As Boolean
   
   For i% = 0 To 4
      optScanType(i%).Enabled = False
      optScanType(i%).Value = False
   Next
   TypeSelected = False
   ChannelType = ANALOGINPUT
   NumChans = FindAnalogChansOfType(mlBoardNum, ChannelType, _
      Resolution, CBRange, LowChan, DefaultTrig)
   If (NumChans > 0) Then
      optScanType(0).Enabled = True
      optScanType(0).Value = True
      TypeSelected = True
      mnAIResolution = Resolution
      mlADRange = CBRange
   End If

   ChannelType = ANALOGOUTPUT
   NumChans = FindAnalogChansOfType(mlBoardNum, ChannelType, _
      Resolution, CBRange, LowChan, DefaultTrig)
   If (NumChans > 0) Then
      optScanType(1).Enabled = True
      If Not TypeSelected Then optScanType(1).Value = True
      TypeSelected = True
      mnAOResolution = Resolution
      mlDARange = CBRange
   End If

   ChannelType = PORTINSCAN
   NumChans = FindPortsOfType(mlBoardNum, ChannelType, _
      PROGPORT, LowChan, NumBits, FirstBit)
   If (NumChans > 0) Then
      optScanType(2).Enabled = True
      If Not TypeSelected Then optScanType(2).Value = True
      TypeSelected = True
      mnDIResolution = NumBits
      mlNumPorts = NumChans
   End If

   ChannelType = PORTOUTSCAN
   NumChans = FindPortsOfType(mlBoardNum, ChannelType, _
      PROGPORT, LowChan, NumBits, FirstBit)
   If (NumChans > 0) Then
      optScanType(3).Enabled = True
      If Not TypeSelected Then optScanType(3).Value = True
      TypeSelected = True
      mnDOResolution = NumBits
      mlNumPorts = NumChans
   End If
   
   ChannelType = CTRSCAN
   NumChans = FindCountersOfType(mlBoardNum, ChannelType, _
      LowChan)
   If (NumChans > 0) Then
      optScanType(4).Enabled = True
      If Not TypeSelected Then optScanType(4).Value = True
      TypeSelected = True
      mnCtrResolution = 32
   End If

End Sub

Private Sub cmdStart_Click()

   lblStatus.ForeColor = &HFF0000
   lblStatus.Caption = ""
   mlRange = mlADRange
   If mlFuncType = AOFUNCTION Then mlRange = mlDARange
   mlNumSamples = 1000
   Resolution = Choose(mlFuncType, mnAIResolution, mnAOResolution, _
      mnDIResolution, mnDOResolution, mnCtrResolution)
   If mnResolution > 16 Then
      ReDim mlDataBuffer32(mlNumSamples)
      mlMemHandle = cbWinBufAlloc32(mlNumSamples)
   Else
      ReDim mlDataBuffer(mlNumSamples)
      mlMemHandle = cbWinBufAlloc(mlNumSamples)
   End If
   If mlMemHandle = 0 Then
      Exit Sub
   End If
   If optTestSelected(0).Value Then
      If Not mbManualRate Then txtRate.Text = 10000000
      mlRate = Val(txtRate.Text)
      lblActualRate.Caption = ""
      mlTestRate = mlRate / 2
      mlPassRate = 0
      txtRate.Enabled = False
      tmrRateGenerator.Interval = 300
      tmrRateGenerator.Enabled = True
   Else
      If (lblSourceRate = "") Then
         mlRate = Val(txtRate.Text)
      Else
         mlRate = Val(lblCurRate.Caption)
      End If
      If Not (lblActualRate.Caption = "") Then
         mlTestRate = Val(lblActualRate.Caption)
      Else
         mlTestRate = mlRate
      End If
      lblActualRate.Caption = ""
      lblSourceRate = "*"
      tmrRateDecrementer.Enabled = True
   End If
   
End Sub

Private Sub optScanType_Click(Index As Integer)

   Dim showDigital As Boolean
   
   showDigital = False
   Select Case Index
      Case 0
         mlFuncType = AIFUNCTION
      Case 1
         mlFuncType = AOFUNCTION
      Case 2
         mlFuncType = DIFUNCTION
         showDigital = True
         GetPortType
         ConfigureOutputs False
      Case 3
         mlFuncType = DOFUNCTION
         showDigital = True
         GetPortType
         ConfigureOutputs True
      Case 4
         mlFuncType = CTRFUNCTION
   End Select
   txtPortIndex.Visible = showDigital
   lblPortIndex.Visible = showDigital
   lblPortsAvailable.Visible = showDigital
   chkArray.Visible = showDigital
   lblPortType.Visible = showDigital
   
End Sub

Function RunScan() As Long

   Dim ULStat As Long
   Dim curRate As Long
   
   curRate = mlRate
   Select Case mlFuncType
      Case AIFUNCTION
         ULStat = cbAInScan(mlBoardNum, 0, 0, mlNumSamples, curRate, mlRange, mlMemHandle, mlOptions)
      Case AOFUNCTION
         ULStat = cbAOutScan(mlBoardNum, 0, 0, mlNumSamples, curRate, mlRange, mlMemHandle, mlOptions)
      Case DIFUNCTION
         ULStat = cbDInScan(mlBoardNum, 1, mlNumSamples, curRate, mlMemHandle, mlOptions)
      Case DOFUNCTION
         ULStat = cbDOutScan(mlBoardNum, 1, mlNumSamples, curRate, mlMemHandle, mlOptions)
      Case CTRFUNCTION
         ULStat = cbCInScan(mlBoardNum, 0, 0, mlNumSamples, curRate, mlMemHandle, mlOptions)
   End Select
   If Not ULStat = 0 Then
      If Not ULStat = BADRATE Then
         tmrRateGenerator.Enabled = False
         ErrMessage$ = GetULError(ULStat)
         lblStatus.Caption = ErrMessage$
         lblStatus.ForeColor = &HFF
         lblCurRate.Caption = "Error"
         Exit Function
      End If
   End If
   lblCurRate.Caption = curRate
   RunScan = ULStat
   
End Function

Private Sub tmrRateDecrementer_Timer()

   Dim ULStat As Long, ActualRate As Long
   Dim sourceRate As Long, Divisor As Long
   Dim actPer As Single, reqPer As Single
   Dim suffix As String
   
   mlRate = mlRate - (mlRate * 0.0005)
   ULStat = RunScan()
   If Not ULStat = 0 Then
      tmrRateDecrementer.Enabled = False
      Exit Sub
   End If
   lblSourceRate.Caption = lblSourceRate.Caption & "*"
   ActualRate = Val(lblCurRate.Caption)
   If ActualRate < mlTestRate Then
      tmrRateDecrementer.Enabled = False
      reqPer = 1 / mlTestRate
      actPer = 1 / ActualRate
      sourceRate = 1 / (actPer - reqPer)
      Select Case sourceRate
         Case Is < 1000
            Divisor = 1
            suffix = " Hz"
         Case Is < 1000000
            Divisor = 1000
            suffix = " kHz"
         Case Else
            Divisor = 1000000
            suffix = " MHz"
      End Select
      lblSourceRate.Caption = Format(sourceRate / Divisor, "0.00") & suffix
   End If
   
End Sub

Private Sub tmrRateGenerator_Timer()

   Dim ULStat As Long, rateIncrement As Long
   Dim TestRate As Long, finalRate As String
   
   If mlRate < 5000 Then tmrRateGenerator.Interval = 1000
   ULStat = RunScan()
   If ULStat = BADRATE Then
      mlFailRate = mlRate
      If mlPassRate > 0 Then
         rateIncrement = (mlRate - mlPassRate) / 2
         mlRate = mlPassRate + rateIncrement
      Else
         mlRate = mlTestRate
         mlTestRate = mlRate / 2
      End If
   Else
      If Abs(mlPassRate - mlRate) < 2 Then
         tmrRateGenerator.Enabled = False
         finalRate = lblCurRate.Caption
         lblCurRate.Caption = "Max request: " & mlRate & "  Actual rate: " & finalRate
         txtRate.Text = mlRate
         txtRate.Enabled = True
         lblActualRate.Caption = finalRate
         mbManualRate = False
         Exit Sub
      End If
      mlPassRate = mlRate
      rateIncrement = (mlFailRate - mlRate) / 2
      mlRate = mlRate + rateIncrement
      mlTestRate = mlRate
   End If
   
End Sub

Private Sub txtPortIndex_Change()

   mlPortIndex = Val(Me.txtPortIndex.Text)
   GetPortType
   
End Sub

Private Sub GetPortType()

   Dim SpanEnabled As Boolean
   Dim sPortList As String
   
   SpanEnabled = Not (chkArray.Value = 0)
   If SpanEnabled Then
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
            sPortList = sPortList & ", " & GetPortStringEx(DevType&)
         Else
            mlNumArrayPorts = 1
            mlNumBits = NumBits&
            PrevDevType& = DevType&
            sPortList = GetPortStringEx(DevType&)
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
         msPortList = GetPortStringEx(mlPortNum)
         cmdStart.Enabled = True
      End If
   
      ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
         mlPortIndex, DINUMBITS, NumBits&)
   End If
   lblPortType.Caption = msPortList
   
End Sub

Private Sub chkArray_Click()

   GetPortType
   
End Sub

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If mlNumPorts > 0 Then ConfigureOutputs False
   
End Sub

Private Sub ConfigureData()

   Dim FS As Long, HS As Long
   
   FS = 2 ^ mnResolution
   HS = FS / 2
   mnDataValHigh = ULongValToInt(HS + (HS) - 1)
   mnDataValLow = ULongValToInt(HS - (HS))

End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      mbManualRate = True
   End If
   
End Sub
