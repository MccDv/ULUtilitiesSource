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
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2100
      Width           =   1215
   End
   Begin VB.TextBox txtPortIndex 
      Height          =   285
      Left            =   420
      TabIndex        =   10
      Text            =   "0"
      Top             =   900
      Width           =   495
   End
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   "1"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Port Input"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1380
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Bit Input"
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txtFirstBit 
      Height          =   225
      Left            =   3180
      TabIndex        =   6
      Text            =   "0"
      Top             =   1380
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtLastBit 
      Height          =   225
      Left            =   3180
      TabIndex        =   5
      Text            =   "-1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   435
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
      Caption         =   "Port Index"
      Height          =   195
      Left            =   1140
      TabIndex        =   16
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2580
      TabIndex        =   15
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Iterations"
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   2220
      Width           =   1275
   End
   Begin VB.Label lblFirstBit 
      Caption         =   "First Bit"
      Height          =   195
      Left            =   3720
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblLastBit 
      Caption         =   "Last Bit"
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   1740
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
Dim mlPortNum As Long, mlNumArrayPorts As Long
Dim mbDoBits As Boolean

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
   For I% = 0 To gnNumBoards - 1
      BoardNum% = gnBoardEnum(I%)
      BoardName$ = GetNameOfBoard(BoardNum%)
      cmbBoard.AddItem BoardName$, I%
   Next I%
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
