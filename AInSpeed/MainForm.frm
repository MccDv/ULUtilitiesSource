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
      TabIndex        =   13
      Top             =   1860
      Width           =   4455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   1860
      Width           =   915
   End
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1260
      TabIndex        =   7
      Text            =   "300"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtFirstChan 
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Text            =   "0"
      Top             =   1140
      Width           =   435
   End
   Begin VB.TextBox txtLastChan 
      Height          =   225
      Left            =   1740
      TabIndex        =   5
      Text            =   "0"
      Top             =   1140
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
         TabIndex        =   14
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
      TabIndex        =   12
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Iterations"
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lblFirstChan 
      Caption         =   "First Chan"
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label lblLastChan 
      Caption         =   "Last Chan"
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   1140
      Width           =   855
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
Dim mlBoardNum As Long, mlNumChans As Long
Dim mlFirstChan As Long, mlLastChan As Long
Dim mlErrReporting As Long, mlErrHandling As Long
Dim mlRange As Long

Private Sub cmbBoard_Click()

   Dim AddBoard As Boolean
   
   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board: " & _
         mlBoardNum & " (type 0x" & Prefix$ & pID$ & ")"
      AddBoard = CheckForAnalog(mlBoardNum)
      lblPortsAvailable.Caption = "Number of channels available: " & mlNumChans
   Else
      lblBoardNumber.Caption = "No Boards Installed"
      AddBoard = False
   End If
   If AddBoard Then
      Me.cmdStart.Enabled = True
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

   mlErrReporting = DONTPRINT
   mlErrHandling = DONTSTOP
   DevsFound& = UpdateDevices(False)
   If Not DevsFound& = 0 Then
      cmdStart.Enabled = True
   End If
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
   
   mlFirstChan = Val(txtFirstChan.Text)
   mlLastChan = Val(txtLastChan.Text)
   NumChans = (mlLastChan - mlFirstChan) + 1
   cmdStart.Enabled = False
   txtResult.Text = ""
   Iterations& = Val(txtRateEstimate.Text)
   
   DataVal% = 0
   ULStat& = cbAIn(mlBoardNum, mlFirstChan, mlRange, DataVal%)
   If ULStat& <> 0 Then
      ErrMessage$ = GetULError(ULStat&)
      txtResult.Text = ErrMessage$
      cmdStart.Enabled = True
      Exit Sub
   End If
         
   DataVal% = 0
   ULStat& = cbAIn(mlBoardNum, mlLastChan, mlRange, DataVal%)
   If ULStat& <> 0 Then
      ErrMessage$ = GetULError(ULStat&)
      txtResult.Text = ErrMessage$
      cmdStart.Enabled = True
      Exit Sub
   End If
   
   StartTime! = Timer
   For i& = 0 To Iterations&
      For Chan& = mlFirstChan To mlLastChan
         ULStat& = cbAIn(mlBoardNum, Chan&, mlRange, DataVal%)
      Next
   Next
   elapsedTime! = Timer - StartTime!
   
   Me.cmdStart.Enabled = True
   inputRate! = 1 / (elapsedTime! / (Iterations& * NumChans))
   FormatString$ = "0.00 Hz"
   divisor! = 1#
   If inputRate! > 999 Then
      FormatString$ = "0.00 kHz"
      divisor! = 1000#
   End If
   txtResult.Text = "Update rate: " & Format(inputRate! / divisor!, FormatString$)
   
End Sub

Private Function CheckForAnalog(ByVal BoardNum As Long) As Boolean
   
   Dim ValidBoard As Boolean
   Dim ReportError As Long, HandleError As Long
   Dim DefaultChan As Long, Resolution As Long
   Dim FirstChan As Long, DefaultRange As Long
   Dim ChannelType As Long, DefaultTrig As Long
   
   ReportError = DONTPRINT
   HandleError = DONTSTOP
   ValidBoard = False
   
   ChannelType = ANALOGINPUT
   NumAIChans = FindAnalogChansOfType(BoardNum, ChannelType, _
      Resolution, DefaultRange, DefaultChan, DefaultTrig)
   mlNumChans = NumAIChans
   If Not (NumAIChans = 0) Then
      ValidBoard = True
      mlFirstChan = DefaultChan
      mlRange = DefaultRange
      mnResolution = Resolution
   Else
      txtResult.Text = ""
   End If
   CheckForAnalog = ValidBoard
   cmdStart.Enabled = ValidBoard

End Function

