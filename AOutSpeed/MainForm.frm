VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Form"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkUseDelay 
      Caption         =   "Use Delay"
      Height          =   255
      Left            =   5340
      TabIndex        =   18
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtDelayMsecs 
      Height          =   315
      Left            =   6780
      TabIndex        =   17
      Text            =   "0.001"
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPercentFS 
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Text            =   "80"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox txtLastChan 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Text            =   "0"
      Top             =   1440
      Width           =   400
   End
   Begin VB.Timer tmrAOut 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4740
      Top             =   1080
   End
   Begin VB.TextBox txtInterval 
      Height          =   315
      Left            =   6780
      TabIndex        =   13
      Text            =   "500"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkTimer 
      Caption         =   "Use Timer"
      Height          =   255
      Left            =   5340
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1740
      TabIndex        =   7
      Text            =   "300"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtFirstChan 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "0"
      Top             =   1080
      Width           =   400
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   180
      TabIndex        =   5
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Frame fraBoard 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox chkUlErrors 
         Caption         =   "UL Errors"
         Height          =   195
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdFlashLED 
         Caption         =   "FlashLED"
         Height          =   315
         Left            =   6720
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
         Width           =   1755
      End
   End
   Begin VB.Label lblPercentFS 
      Caption         =   "% of Full Scale"
      Height          =   195
      Left            =   780
      TabIndex        =   16
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Iterations"
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   1140
      Width           =   2055
   End
   Begin VB.Label lblPortIndex 
      Caption         =   "DAC Channel"
      Height          =   195
      Left            =   780
      TabIndex        =   8
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3060
      Width           =   6915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlBoardNum As Long, mlLowChan As Long
Dim mlFirstChan As Long, mlRange As Long
Dim mlLastChan As Long
Dim mnResolution As Integer, mlIteration As Long
Dim mnDataValLow As Integer, mnDataValHigh As Integer
Dim blEvenIteration As Boolean
Dim msStartTime As Single
Dim mlErrReporting As Long, mlErrHandling As Long

Private Sub cmbBoard_Click()

   Dim ValidBoard As Boolean

   blEvenIteration = False
   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board: " & _
         mlBoardNum & " (type 0x" & Prefix$ & pID$ & ")"
      ValidBoard = CheckForAnalog(mlBoardNum)
   Else
      lblBoardNumber.Caption = "No Boards Installed"
      ValidBoard = False
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

   Me.Caption = App.EXEName & " Test Application"
   mlErrReporting = DONTPRINT
   mlErrHandling = DONTSTOP
   DevsFound& = UpdateDevices(False)
   ConfigureData
   chkUlErrors_Click

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

Private Function CheckForAnalog(ByVal BoardNum As Long) As Boolean
   
   Dim ValidBoard As Boolean
   Dim ReportError As Long, HandleError As Long
   Dim DefaultTrig As Long, DAResolution As Long
   ReportError = DONTPRINT
   HandleError = DONTSTOP
   
   ValidBoard = False
   SetAnalogIODefaults ReportError, HandleError
   Dim ChannelType As Long
   ChannelType = ANALOGOUTPUT
   NumAOChans = FindAnalogChansOfType(BoardNum, ChannelType, _
      DAResolution, mlRange, mlLowChan, DefaultTrig)
   If Not (NumAOChans = 0) Then
      cmdStart.Enabled = True
      'txtHighChan.Enabled = True
      ValidBoard = True
      mnResolution = DAResolution
      MaxChan = LowChan + NumAOChans - 1
   End If
   CheckForAnalog = ValidBoard

End Function

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

   Dim Chan As Long, NumChans As Long
   Dim IntervalValue As Long
   Dim msecDelay As Single
   
   msecDelay = 0#
   If Me.chkUseDelay.Value = 1 Then
      msecDelay = Val(txtDelayMsecs.Text)
   End If
   mlFirstChan = Val(txtFirstChan.Text)
   mlLastChan = Val(txtLastChan.Text)
   mlIteration = 0
   IntervalValue = Val(txtInterval.Text)
   cmdStart.Enabled = False
   Iterations& = Val(txtRateEstimate.Text)
   NumChans = (mlLastChan - mlFirstChan) + 1
   
   ULStat& = cbAOut(mlBoardNum, mlFirstChan, mlRange, mnDataValLow)
   If ULStat& <> 0 Then
      ErrMessage$ = GetULError(ULStat&)
      txtResult.Text = ErrMessage$
      cmdStart.Enabled = True
      Exit Sub
   End If
   
   ULStat& = cbAOut(mlBoardNum, mlLastChan, mlRange, mnDataValHigh)
   If ULStat& <> 0 Then
      ErrMessage$ = GetULError(ULStat&)
      txtResult.Text = ErrMessage$
      cmdStart.Enabled = True
      Exit Sub
   End If
   
   
   If chkTimer.Value = 1 Then
      tmrAOut.Interval = IntervalValue
      tmrAOut.Enabled = True
   Else
      StartTime! = Timer
      For i& = 0 To Iterations&
         For Chan = mlFirstChan To mlLastChan
            ULStat& = cbAOut(mlBoardNum, Chan, mlRange, mnDataValHigh)
         Next
         Pause msecDelay
         For Chan = mlFirstChan To mlLastChan
            ULStat& = cbAOut(mlBoardNum, Chan, mlRange, mnDataValLow)
         Next
         Pause msecDelay
      Next
      elapsedTime! = (Timer - StartTime!) / (2 * NumChans)
      Me.cmdStart.Enabled = True
      outputRate! = 1 / (elapsedTime! / Iterations&)
      FormatString$ = "0.00 Hz"
      Divisor! = 1#
      If outputRate! > 999 Then
         FormatString$ = "0.00 kHz"
         Divisor! = 1000#
      End If
      txtResult.Text = "Update rate: " & Format(outputRate! / Divisor!, FormatString$)
   End If
   
End Sub

Private Sub ConfigureData()

   Dim pctFSR As Long, FS As Long, HS As Long
   Dim pctFS As Single, bpDivisor As Single
   
   pctFS = Val(Me.txtPercentFS) / 100!
   
   FS = 2 ^ mnResolution
   HS = FS / 2
   mnDataValHigh = ULongValToInt(HS + (HS * pctFS) - 1)
   mnDataValLow = ULongValToInt(HS - (HS * pctFS))

End Sub

Private Sub tmrAOut_Timer()

   Dim endIteration As Long, Chan As Long
   
   endIteration = Val(txtRateEstimate.Text)
   DataVal% = mnDataValLow
   If blEvenIteration Then DataVal% = mnDataValHigh
   
   If mlIteration = 0 Then msStartTime = Timer
   For Chan = mlFirstChan To mlLastChan
      ULStat& = cbAOut(mlBoardNum, Chan, mlRange, DataVal%)
   Next
   blEvenIteration = Not blEvenIteration
   txtResult.Text = mlIteration
   mlIteration = mlIteration + 1
   If mlIteration > endIteration Then
      NumChans = (mlLastChan - mlFirstChan) + 1
      elapsedTime! = (Timer - msStartTime) / (2 * NumChans)
      tmrAOut.Enabled = False
      cmdStart.Enabled = True
      outputRate! = 1 / (elapsedTime! / endIteration)
      FormatString$ = "0.00 Hz"
      Divisor! = 1#
      If outputRate! > 999 Then
         FormatString$ = "0.00 kHz"
         Divisor! = 1000#
      End If
      txtResult.Text = "Update rate: " & Format(outputRate! / Divisor!, FormatString$)
   End If

End Sub

Private Sub chkTimer_Click()

   If chkTimer.Value = 1 Then
      txtRateEstimate.Text = "60"
   End If
   txtInterval.Visible = (chkTimer.Value = 1)
   chkUseDelay.Enabled = (chkTimer.Value = 0)
   
End Sub

Private Sub chkUseDelay_Click()
   
   If chkUseDelay.Value = 1 Then
      txtRateEstimate.Text = "300"
   End If
   txtDelayMsecs.Visible = (chkUseDelay.Value = 1)
   chkTimer.Enabled = (chkUseDelay.Value = 0)

End Sub

Private Sub txtPercentFS_Change()

   ConfigureData
   
End Sub
