VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Analog Output Speed Test"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox txtPortIndex 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "0"
      Top             =   1140
      Width           =   495
   End
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1620
      TabIndex        =   5
      Text            =   "300"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame fraBoard 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
   Begin VB.Label lblPortIndex 
      Caption         =   "DAC Channel"
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Rate Estimate"
      Height          =   195
      Left            =   3060
      TabIndex        =   8
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2700
      Width           =   6915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlBoardNum As Long, mlLowChan As Long
Dim mlChan As Long, mlRange As Long
Dim mnResolution As Integer
Dim msDisplayName As String

Private Sub cmbBoard_Click()

   Dim ValidBoard As Boolean
   
   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board Number " & _
         mlBoardNum & " (type 0x" & Prefix$ & pID$ & ")"
      ValidBoard = CheckForAnalog(mlBoardNum)
   Else
      lblBoardNumber.Caption = "No Boards Installed"
      ValidBoard = False
   End If
   cmdStart.Enabled = ValidBoard
   
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

Private Sub cmdStart_Click()

   Me.cmdStart.Enabled = False
   Select Case mnResolution
      Case 10
         DataValLow% = 8
         DataValHigh% = 1000
      Case 12
         DataValLow% = 16
         DataValHigh% = 4080
      Case 16
         DataValLow% = -200
         DataValHigh% = 200
   End Select

   Iterations& = Val(txtRateEstimate.Text) * 5
   AoRange& = 0
   
   ULStat& = cbAOut(mlBoardNum, mlChan, AoRange&, DataValLow%)
   
   For i& = 0 To Iterations&
      ULStat& = cbAOut(mlBoardNum, mlChan, AoRange&, DataValHigh%)
      DoEvents
      ULStat& = cbAOut(mlBoardNum, mlChan, AoRange&, DataValLow%)
      DoEvents
   Next
   Me.cmdStart.Enabled = True
   
End Sub

