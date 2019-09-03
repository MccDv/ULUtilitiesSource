VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   Caption         =   "7266 Counter"
   ClientHeight    =   6855
   ClientLeft      =   105
   ClientTop       =   1545
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6855
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConfig 
      Caption         =   "Configure"
      Height          =   2235
      Left            =   3600
      TabIndex        =   42
      Top             =   780
      Width           =   4095
      Begin VB.CheckBox chkGateOn 
         Caption         =   "Gating On"
         Height          =   195
         Left            =   180
         TabIndex        =   50
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cmbIndex 
         Height          =   315
         ItemData        =   "7266Utility.frx":0000
         Left            =   180
         List            =   "7266Utility.frx":0010
         TabIndex        =   49
         Text            =   "cmbIndex"
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cmbFlags 
         Height          =   315
         ItemData        =   "7266Utility.frx":0044
         Left            =   1980
         List            =   "7266Utility.frx":0054
         TabIndex        =   48
         Text            =   "cmbFlags"
         Top             =   840
         Width           =   1995
      End
      Begin VB.CheckBox chkBinary 
         Caption         =   "Binary Data"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ComboBox cmbMode 
         Height          =   315
         ItemData        =   "7266Utility.frx":0093
         Left            =   1980
         List            =   "7266Utility.frx":00A3
         TabIndex        =   46
         Text            =   "cmbMode"
         Top             =   360
         Width           =   1995
      End
      Begin VB.ComboBox cmbQuadSetting 
         Height          =   315
         ItemData        =   "7266Utility.frx":00D3
         Left            =   180
         List            =   "7266Utility.frx":00E3
         TabIndex        =   45
         Text            =   "cmbQuadSetting"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkInvert 
         Caption         =   "Invert Index"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdConfigure 
         Caption         =   "Configure"
         Height          =   435
         Left            =   1980
         TabIndex        =   43
         Top             =   1440
         Width           =   1635
      End
   End
   Begin VB.HScrollBar hsbPreScale 
      Height          =   255
      Left            =   2820
      Max             =   255
      TabIndex        =   40
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoadPrscl 
      Caption         =   "Load Prescale"
      Height          =   320
      Left            =   180
      TabIndex        =   39
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoadPrst 
      Caption         =   "Load Preset"
      Height          =   320
      Left            =   180
      TabIndex        =   38
      Top             =   3540
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoadCt 
      Caption         =   "Load Count"
      Height          =   320
      Left            =   180
      TabIndex        =   37
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1500
      Top             =   2580
   End
   Begin VB.CheckBox chkLoadPreset 
      Caption         =   "Load Preset on Start"
      Height          =   195
      Left            =   3180
      TabIndex        =   31
      Top             =   3600
      Width           =   2595
   End
   Begin VB.TextBox txtPresetVal 
      Height          =   285
      Left            =   1800
      TabIndex        =   30
      Text            =   "200"
      Top             =   3540
      Width           =   1275
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   1995
      Left            =   60
      TabIndex        =   20
      Top             =   4440
      Width           =   6195
      Begin VB.TextBox txtInterval 
         Height          =   315
         Left            =   4080
         TabIndex        =   35
         Text            =   "2000"
         Top             =   1080
         Width           =   795
      End
      Begin VB.CheckBox chkStopOnDelta 
         Caption         =   "Stop on Delta"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.TextBox txtDelta 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Text            =   "10"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkEnableStatTmr 
         Caption         =   "Enable Timer"
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   780
         Width           =   1455
      End
      Begin VB.Frame fraStatChecks 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox chkSign 
            Caption         =   "Sign Bit (MSB)"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox chkPreset 
            Caption         =   "Preset Match"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   660
            Width           =   1575
         End
         Begin VB.CheckBox chkOFlow 
            Caption         =   "Overflow"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkUFlow 
            Caption         =   "Underflow"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   60
            Width           =   1575
         End
         Begin VB.CheckBox chkError 
            Caption         =   "Input Noise"
            Height          =   195
            Left            =   1920
            TabIndex        =   25
            Top             =   60
            Width           =   1455
         End
         Begin VB.CheckBox chkCountUp 
            Caption         =   "Counting Up"
            Height          =   195
            Left            =   1920
            TabIndex        =   24
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chkIndex 
            Caption         =   "Index Valid"
            Height          =   195
            Left            =   1920
            TabIndex        =   23
            Top             =   660
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdStatus 
         Caption         =   "Update Status"
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblIntervalUnit 
         Caption         =   "ms"
         Height          =   255
         Left            =   4980
         TabIndex        =   36
         Top             =   1140
         Width           =   315
      End
   End
   Begin VB.TextBox txtLoadVal 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Text            =   "100"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CheckBox chkReload 
      Caption         =   "Load Count on Start"
      Height          =   195
      Left            =   3180
      TabIndex        =   18
      Top             =   3180
      Width           =   2595
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   17
      Top             =   3720
      Width           =   795
   End
   Begin VB.Frame fraBoard 
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   120
      Width           =   7635
      Begin VB.ComboBox cmbBoard 
         Height          =   315
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "F5 to update, Ctl-F5 for Ethernet, Shift-Ctl-F5 for remote Ethernet"
         Top             =   180
         Width           =   3075
      End
      Begin VB.CommandButton cmdFlashLED 
         Caption         =   "FlashLED"
         Height          =   315
         Left            =   6480
         TabIndex        =   13
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lblBoardNumber 
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3420
         TabIndex        =   15
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdLoop 
      Caption         =   "Start"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   3240
      Width           =   795
   End
   Begin VB.CommandButton cmdStopRead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Quit"
      Default         =   -1  'True
      Height          =   396
      Left            =   6720
      TabIndex        =   3
      Top             =   5940
      Width           =   795
   End
   Begin VB.Timer tmrReadCounter 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   900
      Top             =   2580
   End
   Begin VB.Label lblPreScale 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   41
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   6540
      Width           =   7695
   End
   Begin VB.Label lblNumReads 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "No of readings taken:"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   180
      TabIndex        =   11
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label lblShowNumReads 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2340
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblShowDirection 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2340
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblDirection 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Delta exceeded value:"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   180
      TabIndex        =   7
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label lblShowLoadVal 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2340
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblShowMaxVal 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2340
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblMaxCount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Count delta:"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Label lblShowReadVal 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2340
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblReadValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Previous value:"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   180
      TabIndex        =   2
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Label lblLoadValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Count for counter:"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   1995
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Const mnCounterType As Integer = CTR7266

Dim mlBoardNum As Long
Dim mlCounterNum As Long, mlDelta As Long
Dim ULStat As Long, mbStop As Boolean

Private Sub cmdConfigure_Click()

   ConfigDevice
   
End Sub

Private Sub Form_Load()

   Dim NumCntrs As Long, DevsFound As Long
   Dim Quadrature As Long, CountingMode As Long
   Dim DataEncoding As Long, IndexMode As Long
   Dim InvertIndex As Long, FlagPins As Long, GateEnable As Long
   Dim LoadValue As Long, RegName As Long
   
   ' declare revision level of Universal Library
   ULStat& = cbDeclareRevision(CURRENTREVNUM)
   Dim CallForm As Form
   ULStat& = cbErrHandling(DONTPRINT, DONTSTOP)
   x% = SaveFunc(CallForm, ErrHandling, ULStat, _
      DONTPRINT, DONTSTOP, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   
   DevsFound& = UpdateDevices(False)
   Me.cmbIndex.ListIndex = 0
   Me.cmbFlags.ListIndex = 0
   Me.cmbQuadSetting.ListIndex = 1
   Me.cmbMode.ListIndex = 0
   DoEvents
   InitDevice
   
End Sub

Sub InitDevice()

   NumCntrs& = FindCountersOfType(mlBoardNum, mnCounterType, mlCounterNum)
   
   If NumCntrs& = 0 Then
      Me.lblStatus.Caption = "No valid counters for this device"
   End If

End Sub

Private Sub ConfigDevice()

   Quadrature& = Choose(cmbQuadSetting.ListIndex + 1, NO_QUAD, X1_QUAD, X2_QUAD, X4_QUAD)
   CountingMode& = Choose(cmbMode.ListIndex + 1, NORMAL_MODE, RANGE_LIMIT, NO_RECYCLE, MODULO_N)
   DataEncoding& = Choose(chkBinary.Value + 1, BCD_ENCODING, BINARY_ENCODING)
   IndexMode& = Choose(cmbIndex.ListIndex + 1, INDEX_DISABLED, LOAD_CTR, LOAD_OUT_LATCH, RESET_CTR)
   InvertIndex& = Choose(chkInvert.Value + 1, CBDISABLED, CBENABLED)
   FlagPins& = Choose(cmbFlags.ListIndex + 1, CARRY_BORROW, COMPARE_BORROW, CARRYBORROW_UPDOWN, INDEX_ERROR)
   GateEnable& = Choose(chkGateOn.Value + 1, CBDISABLED, CBENABLED)
   
   ULStat& = cbC7266Config(mlBoardNum, mlCounterNum, Quadrature&, _
      CountingMode&, DataEncoding&, IndexMode&, InvertIndex&, FlagPins&, GateEnable&)
   x% = SaveFunc(Me, C7266Config, ULStat, _
      mlBoardNum, mlCounterNum, Quadrature&, CountingMode&, DataEncoding&, _
      IndexMode&, InvertIndex&, FlagPins&, GateEnable&, A10, A11, 0)

End Sub

Private Sub cmbBoard_Click()

   Dim Initialized As Boolean
   Dim CallForm As Form
   
   Initialized = Not (lblBoardNumber.Caption = "")
   If cmbBoard.ListCount > 0 Then
      BoardIndex% = cmbBoard.ListIndex
      mlBoardNum = gnBoardEnum(BoardIndex%)
      ULStat = cbGetConfig(BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&)
      x% = SaveFunc(CallForm, GetConfig, ULStat, _
         BOARDINFO, mlBoardNum, 0, BIBOARDTYPE, ConfigVal&, A6, A7, A8, A9, A10, A11, 0)
      pID$ = Hex(ConfigVal&)
      Filler& = 4 - Len(pID$)
      If Filler& > 0 Then Prefix$ = String(Filler&, Chr(48))
      lblBoardNumber.Caption = "Board Number " & _
         mlBoardNum & " (type 0x" & Prefix$ & pID$ & ")"
      'If Initialized Then InitDevice
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

Private Function UpdateDevices(ByVal CheckNet As Boolean, _
   Optional HostString As Variant, Optional HostPort As Long, _
   Optional Timeout As Long) As Long

   Dim devInterface As DaqDeviceInterface
   Dim DevsInstalled As Long, DevsFound As Long
   Dim i As Integer
   
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

Private Sub cmdLoop_Click()

   Dim RegName As Long
   Dim LoadCounter As Boolean
   Dim LoadPreset As Boolean
   
   chkGateOn.Enabled = False
   chkBinary.Enabled = False
   chkInvert.Enabled = False
   cmbFlags.Enabled = False
   cmbIndex.Enabled = False
   cmbMode.Enabled = False
   cmbQuadSetting.Enabled = False
   cmdLoop.Enabled = False
   cmdStop.Enabled = True
   cmdConfigure.Enabled = False
   
   Dim LoadValue As Long, PresetValue As Long
   
   LoadValue = Val(Me.txtLoadVal.Text)
   PresetValue = Val(Me.txtPresetVal.Text)
   
   InitDevice
   LoadCounter = (chkReload.Value = 1)
   LoadPreset = (chkLoadPreset.Value = 1)
   If LoadCounter Then LoadCountVal
   If LoadPreset Then LoadPresetVal
   mbStop = False
   ReadCounter
   
End Sub

Private Sub chkEnableStatTmr_Click()

   If chkEnableStatTmr.Value = 0 Then
      cmdStatus.FontBold = True
      tmrStatus.Enabled = False
   End If
   
End Sub

Private Sub cmdLoadCt_Click()

   LoadCountVal
   
End Sub

Private Sub cmdLoadPrscl_Click()

   LoadPrescaleVal
   
End Sub

Private Sub cmdLoadPrst_Click()

   LoadPresetVal
   
End Sub

Sub LoadCountVal()

   Dim LoadValue As Long
   
   LoadValue = Val(Me.txtLoadVal.Text)
   RegName = COUNT1
   ULStat& = cbCLoad32(mlBoardNum, RegName, LoadValue)
   x% = SaveFunc(Me, CLoad32, ULStat, _
      mlBoardNum, RegName, LoadValue, A4, A5, A6, A7, A8, A9, A10, A11, 0)

End Sub

Sub LoadPresetVal()

   Dim PresetValue As Long
   
   PresetValue = Val(Me.txtPresetVal.Text)
   RegName = PRESET1
   ULStat& = cbCLoad32(mlBoardNum, RegName, PresetValue)
   x% = SaveFunc(Me, CLoad32, ULStat, _
      mlBoardNum, RegName, PresetValue, A4, A5, A6, A7, A8, A9, A10, A11, 0)

End Sub

Sub LoadPrescaleVal()

   Dim PrescaleValue As Long
   
   PrescaleValue = Val(Me.lblPreScale.Caption)
   RegName = PRESCALER1
   ULStat& = cbCLoad32(mlBoardNum, RegName, PrescaleValue)
   x% = SaveFunc(Me, CLoad32, ULStat, _
      mlBoardNum, RegName, PrescaleValue, A4, A5, A6, A7, A8, A9, A10, A11, 0)

End Sub

Private Sub ReadCounter()
   
   Dim CBCount As Long, StatusBits As Long
   Dim deltaHit As Boolean, DeltaRead As Long
   Dim NumReadings As Long, StopOnDelta As Boolean
      
   NumReadings = 0
   ULStat& = cbCIn32(mlBoardNum, mlCounterNum, CBCount&)
   x% = SaveFunc(Me, CIn32, ULStat, _
      mlBoardNum, mlCounterNum, CBCount&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   StatusBits = CBCount&
   StopOnDelta = (chkStopOnDelta.Value = 1)
   mlDelta = Val(txtDelta.Text)
   deltaHit = False
   Do
      ULStat& = cbCIn32(mlBoardNum, mlCounterNum, CBCount&)
      x% = SaveFunc(Me, CIn32, ULStat, _
         mlBoardNum, mlCounterNum, CBCount&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
      NumReadings = NumReadings + 1
      Me.lblShowNumReads.Caption = Format$(NumReadings, "0")
      If Not (StatusBits = 0) Then
         DeltaRead = CBCount& - StatusBits
      End If
      Me.lblShowLoadVal.Caption = Format$(CBCount&, "0")
      Me.lblShowMaxVal.Caption = Format$(DeltaRead, "0")
      If StopOnDelta Then deltaHit = (DeltaRead > mlDelta)
      If Not deltaHit Then StatusBits = CBCount&
      DoEvents
   Loop While Not (deltaHit Or mbStop)
   
   Me.lblShowReadVal.Caption = Hex$(StatusBits)
   lblShowDirection.Caption = Hex$(CBCount&)
   chkGateOn.Enabled = True
   chkBinary.Enabled = True
   chkInvert.Enabled = True
   cmbFlags.Enabled = True
   cmbIndex.Enabled = True
   cmbMode.Enabled = True
   cmbQuadSetting.Enabled = True
   cmdLoop.Enabled = True
   cmdStop.Enabled = False
   cmdConfigure.Enabled = True

End Sub

Private Sub tmrReadCounter_Timer()

   Dim CBCount As Long, StatusBits As Long
   
   ULStat& = cbCIn32(mlBoardNum, mlCounterNum, CBCount&)
   x% = SaveFunc(Me, CIn32, ULStat, _
      mlBoardNum, mlCounterNum, CBCount&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   lblShowReadVal.Caption = Format$(CBCount&, "0")

End Sub

Private Sub cmdStatus_Click()

   GetCtrStatus
   Interval% = 2000
   If IsNumeric(txtInterval.Text) Then Interval% = Val(txtInterval.Text)
   If Me.chkEnableStatTmr.Value Then
      cmdStatus.FontBold = False
      tmrStatus.Interval = Interval%
      Me.tmrStatus.Enabled = True
   End If
   
End Sub

Sub GetCtrStatus()

   ULStat& = cbCStatus(mlBoardNum, mlCounterNum, StatusBits&)
   x% = SaveFunc(Me, CStatus, ULStat, _
      mlBoardNum, mlCounterNum, StatusBits&, A4, A5, A6, A7, A8, A9, A10, A11, 0)
   chkUFlow.Value = ((StatusBits& And C_UNDERFLOW) = C_UNDERFLOW) * -1
   chkOFlow.Value = ((StatusBits& And C_OVERFLOW) = C_OVERFLOW) * -1
   chkPreset.Value = ((StatusBits& And C_COMPARE) = C_COMPARE) * -1
   chkSign.Value = ((StatusBits& And C_SIGN) = C_SIGN) * -1
   chkError.Value = ((StatusBits& And C_ERROR) = C_ERROR) * -1
   chkCountUp.Value = ((StatusBits& And C_UP_DOWN) = C_UP_DOWN) * -1
   chkIndex.Value = ((StatusBits& And C_INDEX) = C_INDEX) * -1

End Sub

Private Sub hsbPreScale_Change()

   Me.lblPreScale.Caption = Me.hsbPreScale.Value

End Sub

Private Sub hsbPreScale_Scroll()

   Me.lblPreScale.Caption = Me.hsbPreScale.Value

End Sub

Private Sub chkGateOn_Click()

   InitDevice
   
End Sub

Private Sub cmdStop_Click()

   mbStop = True
   
End Sub

Private Sub cmdStopRead_Click()
   
   mbStop = True
   DoEvents
   End

End Sub

Private Sub tmrStatus_Timer()

   GetCtrStatus
   cmdStatus.FontBold = Not cmdStatus.FontBold
   
End Sub
