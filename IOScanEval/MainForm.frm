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
   Begin VB.Timer tmrRateGenerator 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   960
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   1980
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
   Begin VB.Label lblCurRate 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1980
      TabIndex        =   13
      Top             =   2100
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
Dim mlBoardNum As Long, mlFuncType As Long, mlRange As Long
Dim mlErrReporting As Long, mlErrHandling As Long
Dim mnAIResolution As Integer, mnAOResolution As Integer
Dim mnDIResolution As Integer, mnDOResolution As Integer
Dim mnCtrResolution As Integer
Dim mlDataBuffer() As Integer
Dim mlDataBuffer32() As Long
Dim mlMemHandle As Long, mlADRange As Long, mlDARange As Long
Dim mlNumSamples As Long
Dim mlTestRate As Long, mlRate As Long
Dim mlFailRate As Long, mlPassRate As Long

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
   End If

   ChannelType = PORTOUTSCAN
   NumChans = FindPortsOfType(mlBoardNum, ChannelType, _
      PROGPORT, LowChan, NumBits, FirstBit)
   If (NumChans > 0) Then
      optScanType(3).Enabled = True
      If Not TypeSelected Then optScanType(3).Value = True
      TypeSelected = True
      mnDOResolution = NumBits
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
   mlRate = 10000000
   mlTestRate = mlRate / 2
   mlPassRate = 0
   tmrRateGenerator.Enabled = True
   
End Sub

Private Sub optScanType_Click(Index As Integer)

   Select Case Index
      Case 0
         mlFuncType = AIFUNCTION
      Case 1
         mlFuncType = AOFUNCTION
      Case 2
         mlFuncType = DIFUNCTION
      Case 3
         mlFuncType = DOFUNCTION
      Case 4
         mlFuncType = CTRFUNCTION
   End Select
   
End Sub

Function RunScan() As Long

   Dim ULStat As Long
   Dim curRate As Long
   
   curRate = mlRate
   ULStat = cbAInScan(mlBoardNum, 0, 0, mlNumSamples, curRate, mlRange, mlMemHandle, mlOptions)
   lblCurRate.Caption = curRate
   RunScan = ULStat
   
End Function

Private Sub tmrRateGenerator_Timer()

   Dim ULStat As Long, rateIncrement As Long
   Dim TestRate As Long, finalRate As String
   
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
         Exit Sub
      End If
      mlPassRate = mlRate
      rateIncrement = (mlFailRate - mlRate) / 2
      mlRate = mlRate + rateIncrement
      mlTestRate = mlRate
   End If
   
End Sub
