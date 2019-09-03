VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Packets by Rate"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbOption 
      Height          =   315
      ItemData        =   "MainForm.frx":0000
      Left            =   5280
      List            =   "MainForm.frx":000E
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtPacketList 
      Height          =   3495
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdRunScan 
      Caption         =   "Run Scan"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
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
Dim mlBoardNum As Long
Dim mlRange As Long

Private Sub cmbBoard_Click()

   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board Number " & _
         mlBoardNum & " (type 0x" & Prefix$ & pID$ & ")"
   Else
      lblBoardNumber.Caption = "No Boards Installed"
   End If
   
End Sub

Private Sub cmdFlashLED_Click()

   cmdFlashLED.Enabled = False
   ULStat = cbFlashLED(mlBoardNum)
   x% = SaveFunc(Me, FlashLED, ULStat, _
      mlBoardNum, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   cmdFlashLED.Enabled = True
   
End Sub

Private Sub cmdRunScan_Click()

   NumPoints& = 100000
   MemHandle& = cbWinBufAlloc(NumPoints&)
   RateVal& = 1
   
   OptAppend& = cmbOption.ItemData(cmbOption.ListIndex)
   Options& = BACKGROUND + OptAppend&
   txtPacketList.Text = cmbBoard.Text & vbCrLf & vbCrLf
   cmdRunScan.Enabled = False
   For i& = 1 To 100
      CurRate& = i& * 2 ^ i&
      RateVal& = CurRate&
      ULStat = cbAOutScan(mlBoardNum, 0, 0, NumPoints&, RateVal&, mlRange, MemHandle&, Options&)
      If SaveFunc(Me, AOutScan, ULStat, mlBoardNum, 0, 0, _
         NumPoints&, RateVal&, 0, MemHandle&, Options&, A9, A10, A11, 0) Then
         cmdRunScan.Enabled = True
         If ULStat = BADRATE Then
            txtPacketList.Text = txtPacketList.Text & LastRate& & vbTab & PacketSize& & vbCrLf
            txtPacketList.SelStart = Len(txtPacketList.Text)
         Else
            txtPacketList.Text = txtPacketList.Text & "Error (see below)"
         End If
         Exit Sub
      End If
      LastRate& = RateVal&
      
      CurCount& = 0
      Do
         PrevCount& = CurCount&
         ULStat = cbGetStatus(mlBoardNum, Status%, CurCount&, CurIndex&, AOFUNCTION)
         DoEvents
      Loop While (PrevCount& < CurCount&) And (Status% = RUNNING)
      ULStat = cbStopBackground(mlBoardNum, AOFUNCTION)
      If CurCount& > PacketSize& Then
         PacketSize& = CurCount&
         txtPacketList.Text = txtPacketList.Text & LastRate& & vbTab & PacketSize& & vbCrLf
         txtPacketList.SelStart = Len(txtPacketList.Text)
      End If
      DoEvents
   Next
   cmdRunScan.Enabled = True
   
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
   Me.cmbOption.ListIndex = 0

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
   ChannelType = ANALOGOUTPUT
   NumAOChans = FindAnalogChansOfType(mlBoardNum, ChannelType, _
       DAResolution&, CBRange&, LowChan&, DefaultTrig&)
   mlRange = CBRange&
   
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
