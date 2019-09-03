VERSION 5.00
Begin VB.Form frmStatusDisplay 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Scan Rate Test"
   ClientHeight    =   7320
   ClientLeft      =   2820
   ClientTop       =   1920
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7320
   ScaleWidth      =   12135
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   4500
      TabIndex        =   49
      Text            =   "5"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox chkLoopTillActive 
      BackColor       =   &H80000005&
      Caption         =   "Loop status until active"
      Height          =   255
      Left            =   180
      TabIndex        =   48
      Top             =   2820
      Value           =   1  'Checked
      Width           =   2715
   End
   Begin VB.CommandButton cmdFlash 
      Appearance      =   0  'Flat
      Caption         =   "Flash LED"
      Height          =   396
      Left            =   5580
      TabIndex        =   46
      Top             =   6840
      Width           =   1260
   End
   Begin VB.CheckBox chkDelay 
      BackColor       =   &H80000005&
      Caption         =   "Delay Restart"
      Height          =   315
      Left            =   960
      TabIndex        =   45
      Top             =   3900
      Width           =   1635
   End
   Begin VB.TextBox txtDelay 
      Enabled         =   0   'False
      Height          =   285
      Left            =   180
      TabIndex        =   44
      Text            =   "100"
      Top             =   3900
      Width           =   675
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   5940
   End
   Begin VB.Frame fraScanType 
      BackColor       =   &H80000005&
      Caption         =   "Scan Type"
      Height          =   915
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Width           =   5355
      Begin VB.OptionButton optScanType 
         BackColor       =   &H80000005&
         Caption         =   "Counter Input"
         Height          =   195
         Index           =   4
         Left            =   3660
         TabIndex        =   47
         Top             =   300
         Width           =   1515
      End
      Begin VB.OptionButton optScanType 
         BackColor       =   &H80000005&
         Caption         =   "Digital Output"
         Height          =   195
         Index           =   3
         Left            =   1980
         TabIndex        =   43
         Top             =   600
         Width           =   1635
      End
      Begin VB.OptionButton optScanType 
         BackColor       =   &H80000005&
         Caption         =   "Digital Input"
         Height          =   195
         Index           =   2
         Left            =   1980
         TabIndex        =   42
         Top             =   300
         Width           =   1635
      End
      Begin VB.OptionButton optScanType 
         BackColor       =   &H80000005&
         Caption         =   "Analog Output"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   1635
      End
      Begin VB.OptionButton optScanType 
         BackColor       =   &H80000005&
         Caption         =   "Analog Input"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   300
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.TextBox txtHWVersion 
      Height          =   285
      Left            =   1980
      TabIndex        =   35
      Top             =   1500
      Width           =   255
   End
   Begin VB.ComboBox cmbTestBoard 
      Height          =   315
      ItemData        =   "AInScanRate.frx":0000
      Left            =   180
      List            =   "AInScanRate.frx":0002
      TabIndex        =   32
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1900
   End
   Begin VB.ComboBox cmbMeasBoard 
      Height          =   315
      Left            =   180
      TabIndex        =   31
      Text            =   "Combo1"
      ToolTipText     =   "USB-CTR0x or HP34401"
      Top             =   1980
      Width           =   1900
   End
   Begin VB.TextBox txtFreqInterval 
      Height          =   285
      Left            =   2760
      TabIndex        =   29
      Text            =   "1"
      Top             =   3900
      Width           =   675
   End
   Begin VB.TextBox txtTmrInterval 
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Text            =   "100"
      Top             =   3540
      Width           =   675
   End
   Begin VB.CommandButton cmdSaveLog 
      Appearance      =   0  'Flat
      Caption         =   "Save Log"
      Height          =   396
      Left            =   8700
      TabIndex        =   24
      Top             =   6840
      Width           =   1260
   End
   Begin VB.TextBox txtRateLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6375
      Left            =   5580
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   60
      Width           =   6435
   End
   Begin VB.TextBox txtNumSamples 
      Height          =   285
      Left            =   4500
      TabIndex        =   15
      Text            =   "5"
      Top             =   2340
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      Height          =   396
      Left            =   3480
      TabIndex        =   14
      Top             =   6840
      Width           =   780
   End
   Begin VB.CommandButton cmdStop 
      Appearance      =   0  'Flat
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   396
      Left            =   1260
      TabIndex        =   12
      Top             =   6840
      Width           =   780
   End
   Begin VB.TextBox txtRate 
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Text            =   "10000"
      ToolTipText     =   "300k max using HP34401"
      Top             =   3540
      Width           =   975
   End
   Begin VB.TextBox txtIterations 
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Text            =   "100"
      Top             =   3180
      Width           =   975
   End
   Begin VB.TextBox txtHighChan 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   4500
      TabIndex        =   9
      Text            =   "0"
      Top             =   1980
      Width           =   495
   End
   Begin VB.CommandButton cmdQuit 
      Appearance      =   0  'Flat
      Caption         =   "Quit"
      Height          =   390
      Left            =   10380
      TabIndex        =   1
      Top             =   6840
      Width           =   1260
   End
   Begin VB.Timer tmrCheckStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2100
      Top             =   5940
   End
   Begin VB.CommandButton cmdStartBgnd 
      Appearance      =   0  'Flat
      Caption         =   "Start"
      Height          =   396
      Left            =   240
      TabIndex        =   0
      Top             =   6840
      Width           =   780
   End
   Begin VB.Label lblTimeout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "Timeout"
      Height          =   195
      Left            =   3480
      TabIndex        =   50
      Top             =   2820
      Width           =   915
   End
   Begin VB.Label lblFirmwareVersion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2700
      TabIndex        =   38
      Top             =   1500
      Width           =   2715
   End
   Begin VB.Label lblUniqueID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3540
      TabIndex        =   37
      Top             =   1140
      Width           =   1875
   End
   Begin VB.Label lblHWVersion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "HW ver (A, B, etc)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label lblCurIteration 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   2820
      TabIndex        =   34
      Top             =   3240
      Width           =   1500
   End
   Begin VB.Label lblNumSamples 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "Samples"
      Height          =   195
      Left            =   3480
      TabIndex        =   33
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label lblFreqInterval 
      BackColor       =   &H80000005&
      Caption         =   "Freq Interval"
      Height          =   195
      Left            =   3540
      TabIndex        =   30
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblTmrInterval 
      BackColor       =   &H80000005&
      Caption         =   "Timer Interval"
      Height          =   195
      Left            =   3540
      TabIndex        =   28
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblIterations 
      BackColor       =   &H80000005&
      Caption         =   "Iterations"
      Height          =   195
      Left            =   1260
      TabIndex        =   26
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblStartFreq 
      BackColor       =   &H80000005&
      Caption         =   "Starting Freq"
      Height          =   195
      Left            =   1260
      TabIndex        =   25
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label lblMaxDiffMeas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   180
      TabIndex        =   23
      Top             =   6480
      Width           =   4650
   End
   Begin VB.Label lblMeasBoard 
      BackColor       =   &H80000005&
      Caption         =   "Measure Board"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   1980
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDiffRetVsMeas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Difference Returned vs Measured:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   180
      TabIndex        =   21
      Top             =   5760
      Width           =   3540
   End
   Begin VB.Label lblShowDiffAvM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   3810
      TabIndex        =   20
      Top             =   5760
      Width           =   990
   End
   Begin VB.Label lblMeasRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Measured Rate:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   180
      TabIndex        =   19
      Top             =   4920
      Width           =   1545
   End
   Begin VB.Label lblShowMeasRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1905
      TabIndex        =   18
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label lblBoardNum 
      BackColor       =   &H80000005&
      Caption         =   "Test Board"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label lblMaxDiff 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   180
      TabIndex        =   13
      Top             =   6180
      Width           =   4650
   End
   Begin VB.Label lblMeas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Scan 0 to"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   1980
      Width           =   1035
   End
   Begin VB.Label lblShowCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1890
      TabIndex        =   7
      Top             =   4650
      Width           =   1500
   End
   Begin VB.Label lblReturnedRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rate Returned:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   165
      TabIndex        =   5
      Top             =   4650
      Width           =   1545
   End
   Begin VB.Label lblShowReqRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1890
      TabIndex        =   6
      Top             =   4365
      Width           =   1500
   End
   Begin VB.Label lblReqRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Requested Rate:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   165
      TabIndex        =   4
      Top             =   4365
      Width           =   1545
   End
   Begin VB.Label lblShowReqVsRet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   3840
      TabIndex        =   3
      Top             =   5460
      Width           =   990
   End
   Begin VB.Label lblReqVsRet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Difference Requested vs Returned:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   150
      TabIndex        =   2
      Top             =   5460
      Width           =   3540
   End
   Begin VB.Menu nmuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAppHelp 
         Caption         =   "AInScanRate Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "Help About"
      End
   End
End
Attribute VB_Name = "frmStatusDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const BIFWVERSION = 170
Const BIMFGSERIALNUM = 224

Dim nlIterations As Long
Dim nlRate As Long, nlCount As Long
Dim nmLowChan As Integer, mnHighChan As Integer
Dim mnMaxDifference As Long
Dim mlCurRate As Long, mlCurReturn As Long
Dim mlMaxRateReturnedErr As Long, mlMinRateReturnedErr As Long
Dim mlCurDiff As Long
Dim mdMaxDiffMeas As Double
Dim mdCurMeasDiff As Double
Dim mdCurMeas As Double
Dim mlBoardNum As Long, mlMeasBoardNum As Long
Dim mnMeterAddress As Integer
Dim mlFreqInterval As Long
Dim mbMeasBoardExists As Boolean
Dim mbMeasInstrumentExists As Boolean
Dim CBRange As Long
Dim mnResolution As Long, NumAIChans As Long
Dim NumDIChans As Long, mlPortNum As Long
Dim mlCtrNum As Long, mlFuncType As Long

Dim HighChan As Long, LowChan As Long, MaxChan As Long
Dim mbReset As Boolean

Const mlColorEnabled As Long = &H80000008
Const mlColorDisabled As Long = &H80000011

'Const NumPoints As Long = 6000     ' Number of data points to collect
Const FirstPoint As Long = 0       ' set first element in buffer to transfer to array

Dim mlDataBuffer() As Integer      ' dimension an array to hold the input values
Dim mlDataBuffer32() As Long       ' dimension an array to hold the high resolution input values
Dim mlMemHandle As Long        ' define a variable to contain the handle for
                             ' memory allocated by Windows through cbWinBufAlloc()
Dim ULStat As Long
Dim mbDelayRestart As Boolean

Private Sub chkDelay_Click()

   txtDelay.Enabled = (chkDelay.value = 1)
   mbDelayRestart = (chkDelay.value = 1)
   
End Sub

Private Sub cmbMeasBoard_Click()

   Dim SetColor As Long
   
   mlMeasBoardNum = cmbMeasBoard.ItemData(cmbMeasBoard.ListIndex)
   SetColor = mlColorEnabled
   mbMeasBoardExists = True
   If mlMeasBoardNum = -2 Then
      SetColor = mlColorDisabled
      mbMeasBoardExists = False
   End If
   lblMeasRate.ForeColor = SetColor
   lblDiffRetVsMeas.ForeColor = SetColor
   
End Sub

Private Sub cmbTestBoard_Click()

   mlBoardNum = cmbTestBoard.ItemData(cmbTestBoard.ListIndex)
   GetDeviceInfo

End Sub

Private Sub cmdClear_Click()

   If Not Me.tmrCheckStatus.Enabled Then mbReset = True
   mnMaxDifference = 0
   mlCurRate = 0: mlCurReturn = 0
   mlCurDiff = 0: mdMaxDiffMeas = 0
   mdCurMeasDiff = 0: mdCurMeas = 0
   lblMaxDiff.Caption = ""
   lblMaxDiffMeas.Caption = ""
   txtRateLog.Text = ""
   lblCurIteration.Caption = ""
   If mbMeasBoardExists Then
      If mlMeasBoardNum = -1 Then _
         GPIBWrite mnMeterAddress, "*CLS"
   End If
   
End Sub

Private Sub cmdHelp_Click()
   
   Dim RegNode As String
   Dim NumNodesFound As Long
   Dim Values As Variant
   Dim Node As Long

   RegNode = "Software\Measurement Computing"
   NumNodesFound = FindSubNodes(HKEY_LOCAL_MACHINE, RegNode, Values)
   For Node = 0 To NumNodesFound
      txtRateLog.Text = txtRateLog.Text & Values(Node) & vbCrLf
   Next

End Sub

Private Sub cmdFlash_Click()

   ULStat = cbFlashLED(mlBoardNum)
   If ULStat <> 0 Then ShowErrorDlg ULStat
   
End Sub

Private Sub cmdSaveLog_Click()

   Dim filename As String, Path As String
   Dim ConfigString As String, SerNum As String
   Dim TestBoard As String, FWVersion As String
   Dim ULStat As Long, ConfigVal As Long
   Dim HWVersion As String, ScanType As String
   
   If txtHWVersion.Text = "" Then
      HWVersion = InputBox("Enter hardware version " & _
         "(A, B, etc.) or leave blank.", "Hardware Version", "")
      If Not HWVersion = "" Then
         Me.txtHWVersion.Text = HWVersion
         HWVersion = "(" & HWVersion & ")"
      End If
   Else
      HWVersion = "(" & txtHWVersion.Text & ")"
   End If
   FWVersion = lblFirmwareVersion.Caption
   SerNum = Me.lblUniqueID.Caption
   If FWVersion = "" Then FWVersion = "unknown"
   If SerNum = "" Then SerNum = "unknown"
   
   Select Case mlFuncType
      Case AIFUNCTION
         ScanType = "AInScanRate"
      Case AOFUNCTION
         ScanType = "AOutScanRate"
      Case DIFUNCTION
         ScanType = "DInScanRate"
      Case DOFUNCTION
         ScanType = "DOutScanRate"
   End Select
   TestBoard = cmbTestBoard.Text
   TestBoard = Replace(TestBoard, "/", "_")
   frmFileDialog.txtFileName.Text = ScanType & _
      TestBoard & HWVersion & ".txt"
   frmFileDialog.Show 1, Me
   filename = frmFileDialog.txtFileName.Text
   If Not filename = "" Then
      Path = frmFileDialog.Dir1.Path & "\" & frmFileDialog.txtFileName.Text
      Open Path For Output As #1
      Print #1, cmbTestBoard.Text & " Rate Test"
      Print #1, "Firmware version " & FWVersion & "  Serial number " & SerNum & vbCrLf
      Print #1, nlIterations & " measurements made at " & mlFreqInterval & _
         "Hz intervals between " & nlRate & "Hz and " & lblShowReqRate.Caption & "Hz."
      Print #1, "Pacer output measured using " & cmbMeasBoard.Text & "."
      Print #1, "Scanned channels 0 to " & txtHighChan.Text & _
         " with number of samples set to " & txtNumSamples.Text & " per channel." & vbCrLf & vbCrLf
      Print #1, lblMaxDiff.Caption
      Print #1, "Range of error requested vs. returned: " & _
         mlMinRateReturnedErr & " to " & mlMaxRateReturnedErr & "."
      If mbMeasBoardExists Then _
         Print #1, lblMaxDiffMeas.Caption
      Print #1, vbCrLf
      Print #1, txtRateLog.Text
      Close #1
   End If
   Unload frmFileDialog
   
End Sub

Private Sub cmdStop_Click()

   Me.tmrCheckStatus.Enabled = False
   ULStat = cbStopBackground(mlBoardNum, mlFuncType)
   Me.cmdStartBgnd.Enabled = True
   Me.cmdQuit.Enabled = True
   cmdStop.Enabled = False
   
End Sub

Private Sub Form_Load()
   
   Dim ULStat As Long, AddsFound As Integer
   Dim ReportError As Long, HandleError As Long
   Dim BoardName As String, StatString As String
   
   ' declare revision level of Universal Library

   ULStat = cbDeclareRevision(CURRENTREVNUM)

   ReportError = DONTPRINT
   HandleError = DONTSTOP
   ULStat = cbErrHandling(ReportError, HandleError)
   If ULStat <> 0 Then ShowErrorDlg ULStat
   
   PopulateBoards

End Sub

Private Sub cmdStartBgnd_Click()
   
   Dim NumChans As Integer
   Dim StartRate As Long, NumPoints As Long
   
   If Not mlMemHandle = 0 Then
      ULStat = cbWinBufFree(mlMemHandle)
   End If
   ' Check the resolution of the A/D data and allocate memory accordingly
   NumPoints = Val(txtNumSamples.Text)
   mnHighChan = Val(txtHighChan.Text)
   If (mnHighChan > MaxChan) Then mnHighChan = MaxChan
   txtHighChan.Text = Format(mnHighChan, "0")
   NumChans = ((mnHighChan - LowChan) + 1)
   
   nlCount = NumPoints * NumChans
   If mnResolution > 16 Then
      ' set aside memory to hold high resolution data
      ReDim mlDataBuffer32(nlCount)
      mlMemHandle = cbWinBufAlloc32(nlCount)
   Else
      ' set aside memory to hold data
      ReDim mlDataBuffer(nlCount)
      mlMemHandle = cbWinBufAlloc(nlCount)
   End If
   If mlMemHandle = 0 Then
      ShowErrorDlg BADCOUNT
      Exit Sub
   End If
   If mlFuncType = DOFUNCTION Then
      ULStat& = cbDConfigPort(mlBoardNum, mlPortNum, DIGITALOUT)
      If ULStat <> 0 Then
         ShowErrorDlg ULStat
         Exit Sub
      End If
      If Not (LoadData() = 0) Then Exit Sub
   End If
   
   StartRate = Me.txtRate.Text
   nlRate = StartRate / NumChans
   cmdStartBgnd.Enabled = False
   cmdStop.Enabled = True
   cmdQuit.Enabled = False
   
   If mbMeasBoardExists Then
      If mlMeasBoardNum = -1 Then
         Dim addr As Integer
         addr = cmbMeasBoard.ItemData(cmbMeasBoard.ListIndex)
         GPIBWrite mnMeterAddress, "CONF:FREQ DEF, DEF"
         If CheckGPIBError() Then Exit Sub
      Else
         ULStat = cbCConfigScan(mlMeasBoardNum, 0, _
            PERIOD, CTR_DEBOUNCE_NONE, 0, 0, 0, 0)
         If ULStat <> 0 Then
            ShowErrorDlg ULStat
            Exit Sub
         End If
      End If
   End If
   mlFreqInterval = Val(txtFreqInterval.Text)
   nlIterations = Val(txtIterations.Text)
   tmrCheckStatus.Interval = Val(Me.txtTmrInterval.Text)
   tmrCheckStatus.Enabled = True
   
End Sub

Private Sub mnuHelpAbout_Click()

   frmAbout.Show 1
   
End Sub

Private Sub optScanType_Click(Index As Integer)

   PopulateBoards
   
End Sub

Private Sub tmrCheckStatus_Timer()

   Static sR As Long
   Static sD As Long
   Static sNumReports As Long
   Static sReport As Boolean
   
   Dim ULStat As Long, Difference As Long
   Dim CBRate As Long, CtrVal As Long
   Dim Options As Long, FixedInterval As Boolean
   Dim RequestedRate As Long
   Dim PeriodVal As Double, FreqVal As Double
   Dim MeasDiff As Double, T0 As Single
   Dim CurLine As String, DataRead As String
   Dim Transition As Long, Timeout As Long
   Dim CurCount As Long, CurIndex As Long
   Dim Status As Integer, Spacer As String

   ULStat = cbStopBackground(mlBoardNum, mlFuncType)
   If ULStat <> 0 Then
      ShowErrorDlg ULStat
      Exit Sub
   End If
   'ULStat = cbStopBackground(mlBoardNum, mlFuncType)
   If mbDelayRestart Then
      tmrCheckStatus.Enabled = False
      tmrDelay.Interval = Val(txtDelay.Text)
      tmrDelay.Enabled = True
      Do
         DoEvents
      Loop While tmrDelay.Enabled
      tmrCheckStatus.Enabled = True
   End If
   
   If mbReset Then
      sR = 0
      sD = 0
      sNumReports = 0
   End If
   mbReset = False
   Options = BACKGROUND + CONTINUOUS
   Timeout = Val(txtTimeout.Text)
   CBRate = nlRate - (sR * mlFreqInterval)
   If Not (CBRate > 0) Then
      Me.cmdStop = True
      Exit Sub
   End If
   sR = sR + 1
   lblCurIteration.Caption = sR
   RequestedRate = CBRate
   Select Case mlFuncType
      Case AIFUNCTION
         ULStat = cbAInScan(mlBoardNum, LowChan, HighChan, _
            nlCount, CBRate, CBRange, mlMemHandle, Options)
      Case AOFUNCTION
      Case DIFUNCTION
         ULStat = cbDInScan(mlBoardNum, mlPortNum, nlCount, _
            CBRate, mlMemHandle, Options)
      Case DOFUNCTION
         ULStat = cbDOutScan(mlBoardNum, mlPortNum, nlCount, _
            CBRate, mlMemHandle, Options)
      Case CTRFUNCTION
         ULStat = cbCInScan(mlBoardNum, mlCtrNum, _
            mlCtrNum, nlCount, CBRate, mlMemHandle, Options)
   End Select
   If ULStat <> 0 Then
      ShowErrorDlg ULStat
      Exit Sub
   End If
   If Me.chkLoopTillActive.value = 1 Then
      T0 = Timer()
      Do
         ULStat = cbGetStatus(mlBoardNum, Status, _
            CurCount, CurIndex, mlFuncType)
         If ULStat <> 0 Then
            ShowErrorDlg ULStat, Status, CurCount, CurIndex
            Exit Sub
         End If
         If Timer() - T0 > Timeout Then
            CurLine = mlCurRate & " returned " & mlCurReturn & vbCrLf & vbCrLf
            CurLine = CurLine & vbTab & "Device failed to start scan." & vbCrLf
            CurLine = CurLine & vbTab & "Status = " & Status & _
               ", CurCount = " & CurCount & vbCrLf
            cmdStop_Click
            MsgBox "Timeout occurred", vbCritical, "Device Hung"
            
            CurLine = CurLine & vbCrLf
            txtRateLog.Text = txtRateLog.Text & CurLine
            txtRateLog.SelStart = Len(txtRateLog.Text)
            Exit Sub
         End If
         DoEvents
      Loop While ((CurIndex = -1) And (Status = RUNNING))
   End If
   
   lblShowCount.Caption = Format$(CBRate, "0")
   lblShowReqRate.Caption = Format$(RequestedRate, "0")
   Difference = RequestedRate - CBRate
   If Difference < mlMinRateReturnedErr Then mlMinRateReturnedErr = Difference
   If Difference > mlMaxRateReturnedErr Then mlMaxRateReturnedErr = Difference
   If Abs(Difference) > Abs(mnMaxDifference) Then
      mnMaxDifference = Difference
      lblMaxDiff.Caption = "Max returned vs requested @ " _
         & RequestedRate & ": " & mnMaxDifference
   End If
   
   If mbMeasBoardExists Then
      If mlMeasBoardNum = -1 Then
         Dim BufSize As Integer
         BufSize = 128
         GPIBWrite mnMeterAddress, "READ?"
         If CheckGPIBError() Then Exit Sub
         GPIBRead mnMeterAddress, DataRead, BufSize
         If CheckGPIBError() Then Exit Sub
         FreqVal = Val(DataRead)
      Else
         If CBRate < 30 Then
            Dim TX As Single, TW As Single
            T0 = Timer
            TW = 0.5
            If CBRate < 3 Then TW = 1 / CBRate + 0.5
            Do
               TX = Timer - T0
            Loop While TX < TW
         End If
         ULStat = cbCIn32(mlMeasBoardNum, 0, CtrVal)
         If ULStat <> 0 Then
            ShowErrorDlg ULStat
            Exit Sub
         End If
         PeriodVal = CtrVal * 0.00000002083  '083
         If PeriodVal > 0 Then FreqVal = 1 / PeriodVal
      End If
      MeasDiff = FreqVal - CBRate
      If Abs(MeasDiff) > Abs(mdMaxDiffMeas) Then
         mdMaxDiffMeas = MeasDiff
         lblMaxDiffMeas.Caption = "Max returned vs measured error @ " _
            & RequestedRate & ": " & Format(mdMaxDiffMeas, "0.0")
      End If
      Me.lblShowMeasRate.Caption = Format(FreqVal, "0.0")
      Me.lblShowDiffAvM.Caption = Format(MeasDiff, "0.0")
   Else
      Me.lblShowMeasRate.Caption = ""
      Me.lblShowDiffAvM.Caption = ""
   End If
   Me.lblShowReqVsRet.Caption = RequestedRate - CBRate
   If sR > 1 Then Transition = Abs(Difference - sD)
   FixedInterval = (sR = (nlIterations / 2)) Or (sR = (nlIterations - 1))
   If (Transition > mlFreqInterval) Or sReport Then
      CurLine = mlCurRate & " returned " & mlCurReturn & _
         " (" & mlCurDiff & ") "
      If (mdCurMeas < 1000000) Then Spacer$ = vbTab
      If mbMeasBoardExists Then CurLine = CurLine & _
         Spacer$ & "Meas " & Format(mdCurMeas, "0.0") & _
         " (" & Format(mdCurMeasDiff, "0.0") & ")"
      CurLine = CurLine & vbCrLf
      txtRateLog.Text = txtRateLog.Text & CurLine
      txtRateLog.SelStart = Len(txtRateLog.Text)
      mlCurDiff = 0
      If nlIterations > 60 Then sReport = False
      sNumReports = sNumReports + 1
   Else
      If (Abs(Difference) > Abs(mlCurDiff)) Or FixedInterval Then
         mlCurRate = RequestedRate
         mlCurDiff = Difference
         mlCurReturn = CBRate
         mdCurMeas = FreqVal
         mdCurMeasDiff = MeasDiff
         sReport = FixedInterval
      End If
   End If
   If ULStat <> 0 Then
      ShowErrorDlg ULStat
      Exit Sub
   End If
   If (sR >= nlIterations) Or (RequestedRate = 1) Then
      ULStat = cbStopBackground(mlBoardNum, mlFuncType)
      If ULStat <> 0 Then
         ShowErrorDlg ULStat
         Exit Sub
      End If
      sR = 0
      sReport = False
      Me.tmrCheckStatus.Enabled = False
      cmdQuit.Enabled = True
      cmdStartBgnd.Enabled = True
      cmdStop.Enabled = False
      lblCurIteration.Caption = ""
   Else
      If nlIterations > 60 Then
         sReport = sReport Or ((sNumReports + 1) _
            < (sR / (nlIterations / 10)))
      Else
         sReport = (Not sReport) And (Difference <> 0)
      End If
   End If
   sD = Difference

End Sub

Private Sub cmdQuit_Click()
   
   ULStat = cbWinBufFree(mlMemHandle)      ' Free up memory for use by
                                            ' other programs
   End
   
End Sub

Private Function GetDevName(ByVal BoardNum As Long) As String

   Dim BoardName As String
   
   BoardName = Space$(BOARDNAMELEN)
   ULStat& = cbGetBoardName(BoardNum, BoardName)
   BoardName = Trim(BoardName)
   If Len(BoardName) > 0 Then _
      BoardName = Left(BoardName, Len(BoardName) - 1)
   GetDevName = BoardName

End Function

Private Function CheckForAnalog(ByVal BoardNum As Long) As Boolean
   
   Dim ValidBoard As Boolean
   Dim ReportError As Long, HandleError As Long
   Dim DefaultTrig As Long, ADResolution As Long
   ReportError = DONTPRINT
   HandleError = DONTSTOP
   
   ValidBoard = False
   SetAnalogIODefaults ReportError, HandleError
   Dim ChannelType As Long
   ChannelType = ANALOGINPUT
   NumAIChans = FindAnalogChansOfType(BoardNum, ChannelType, _
      ADResolution, CBRange, LowChan, DefaultTrig)
   If Not (NumAIChans = 0) Then
      cmdStartBgnd.Enabled = True
      txtHighChan.Enabled = True
      ValidBoard = True
      mnResolution = ADResolution
      MaxChan = LowChan + NumAIChans - 1
   End If
   CheckForAnalog = ValidBoard

End Function

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
   ChannelType = PORTINSCAN
   NumDIChans = FindPortsOfType(BoardNum, ChannelType, _
      PROGPORT, DefaultPort, DefaultNumBits, FirstBit)
   If Not (NumDIChans = 0) Then
      cmdStartBgnd.Enabled = True
      txtHighChan.Enabled = True
      ValidBoard = True
      mlPortNum = DefaultPort
      mnResolution = DefaultNumBits
   End If
   CheckForDigital = ValidBoard

End Function

Private Function CheckForCounter(ByVal BoardNum As Long) As Boolean
   
   Dim ValidBoard As Boolean
   Dim ReportError As Long, HandleError As Long
   Dim DefaultCounter As Long
   ReportError = DONTPRINT
   HandleError = DONTSTOP
   
   ValidBoard = False
   'SetDigitalIODefaults ReportError, HandleError
   Dim ChannelType As Long
   ChannelType = CTRSCAN
   NumDIChans = FindCountersOfType(BoardNum, ChannelType, _
      DefaultCounter)
   If Not (NumDIChans = 0) Then
      cmdStartBgnd.Enabled = True
      txtHighChan.Enabled = True
      ValidBoard = True
      mlCtrNum = DefaultCounter
   End If
   CheckForCounter = ValidBoard

End Function

Public Sub ShowErrorDlg(ByVal ErrCode As Long, Optional Status As Integer, _
   Optional CurCount As Long, Optional CurIndex As Long)

   Dim ErrMessage As String
   Dim ErrStat As Long
   Dim StatString As String
   
   Me.tmrCheckStatus.Enabled = False
   If Not ((Status = 0) _
      And (CurCount = 0) _
      And (CurIndex = 0)) Then
         StatString = vbCrLf & vbCrLf & _
         "    " & "Status = " & Status & _
         ", CurCount = " & CurCount & _
         ", CurIndex = " & CurIndex
   Else
      ULStat = cbGetStatus(mlBoardNum, Status, _
         CurCount, CurIndex, mlFuncType)
         StatString = vbCrLf & vbCrLf & _
         "    " & "Status = " & Status & _
         ", CurCount = " & CurCount & _
         ", CurIndex = " & CurIndex
   End If
   ULStat = cbStopBackground(mlBoardNum, mlFuncType)
   ErrMessage$ = Space$(ERRSTRLEN)
   ErrStat = cbGetErrMsg(ErrCode, ErrMessage)
   ErrMessage$ = RTrim$(ErrMessage$)   'Drop the space characters
   ErrMessage$ = Left$(ErrMessage$, Len(ErrMessage$) - 1)
   MsgBox ErrMessage, vbCritical, "Universal Library Error"
      
   txtRateLog.Text = txtRateLog.Text & vbCrLf & _
      vbCrLf & vbTab & "Universal Library Error" & _
      vbCrLf & vbCrLf & "    Iteration " & _
      Me.lblCurIteration.Caption & " of " & txtIterations.Text _
      & ", rate requested " & Me.lblShowReqRate & _
      vbCrLf & vbCrLf & "    " & ErrMessage & StatString
   Me.cmdStartBgnd.Enabled = True
   Me.cmdQuit.Enabled = True
   cmdStop.Enabled = False
   
End Sub

Private Sub PopulateBoards()

   Dim BoardNum As Long, AddsFound As Integer
   Dim BoardName As String, MeasureBoard As Boolean
   Dim SetColor As Long, AddBoard As Boolean
   Dim Result As VbMsgBoxResult
   Dim PromptText As String
   
   'mlBoardNum = 0
   cmbTestBoard.Clear
   cmbMeasBoard.Clear
   cmbMeasBoard.AddItem "None"
   cmbMeasBoard.ItemData(cmbMeasBoard.NewIndex) = -2
   For BoardNum = 0 To 30
      BoardName = GetDevName(BoardNum)
      If Not BoardName = "" Then
         MeasureBoard = ((BoardName = "USB-CTR08") _
         Or (BoardName = "USB-CTR04"))
         If MeasureBoard Then
            cmbMeasBoard.AddItem BoardName
            cmbMeasBoard.ItemData(cmbMeasBoard.NewIndex) = BoardNum
            mbMeasBoardExists = True
            MeasureBoard = False
         End If
         Select Case True
            Case Me.optScanType(0).value
               mlFuncType = AIFUNCTION
               AddBoard = CheckForAnalog(BoardNum)
            Case optScanType(2).value
               mlFuncType = DIFUNCTION
               AddBoard = CheckForDigital(BoardNum)
            Case optScanType(3).value
               mlFuncType = DOFUNCTION
               AddBoard = CheckForDigital(BoardNum)
            Case Me.optScanType(4).value
               mlFuncType = CTRFUNCTION
               AddBoard = CheckForCounter(BoardNum)
         End Select
         If AddBoard Then
            cmbTestBoard.AddItem BoardName
            cmbTestBoard.ItemData(cmbTestBoard.NewIndex) = BoardNum
         End If
      End If
   Next

   If Not mbMeasBoardExists Then
      PromptText = "No MCC measurement device installed. "
   Else
      PromptText = "MCC measurement device is available. "
   End If
   Result = MsgBox(PromptText & "Search for GPIB configured on this PC?", vbYesNo, "Initialize GPIB?")
   If Result = vbNo Then Exit Sub
   
   If InitGPIB Then
      'check for compatible instrument
      Const AddrRange As Integer = 29
      Dim NumFound As Integer, BufSize As Integer
      Dim Command As String, DataRead As String
      Dim ListOfAddresses() As Integer
      
      NumFound = GetAddressList(ListOfAddresses(), AddrRange)
      Command = "*IDN?"
      BufSize = 128
      For AddsFound = 0 To NumFound - 1
         GPIBWrite ListOfAddresses(AddsFound), Command
         If CheckGPIBError() Then Exit Sub
         GPIBRead ListOfAddresses(AddsFound), DataRead, BufSize
         If CheckGPIBError() Then Exit Sub
         If InStr(1, DataRead, "34401") > 0 Then
            mbMeasInstrumentExists = True
            mbMeasBoardExists = True
            mnMeterAddress = ListOfAddresses(AddsFound)
            cmbMeasBoard.AddItem "HP34401"
            cmbMeasBoard.ItemData(cmbMeasBoard.NewIndex) = -1
            Exit For
         End If
      Next
   End If
   If cmbTestBoard.ListCount > 0 Then cmbTestBoard.ListIndex = 0
   If cmbMeasBoard.ListCount > 0 Then _
      cmbMeasBoard.ListIndex = cmbMeasBoard.ListCount - 1
   SetColor = mlColorDisabled
   If mbMeasBoardExists Then _
      SetColor = mlColorEnabled
   lblMeasRate.ForeColor = SetColor
   lblDiffRetVsMeas.ForeColor = SetColor

End Sub

Private Sub GetDeviceInfo()

   Dim ConfigVal As String, CurItem As String
   Dim Version As Single, ConfigLen As Long
   Dim DevNum As Long, VersionString As String
   
   ConfigLen = 64
   DevNum = 0
   ConfigVal = Space(ConfigLen)
   ULStat& = cbGetConfigString(BOARDINFO, mlBoardNum, _
      DevNum, BIDEVUNIQUEID, ConfigVal, ConfigLen)
   If (ULStat& = 0) And (ConfigLen > 0) Then
      lblUniqueID.Caption = Left(ConfigVal, ConfigLen)
   Else
      lblUniqueID.Caption = ""
   End If
   
   lblFirmwareVersion.Caption = ""
   For DevNum& = 0 To 4
      ConfigLen = 32
      ULStat& = cbGetConfigString(BOARDINFO, mlBoardNum, _
         DevNum&, BIDEVVERSION, ConfigVal, ConfigLen)
      If (ULStat& = 0) And (ConfigLen > 0) Then
         CurItem = Choose(DevNum + 1, "Main=", "Meas=", "Radio=", "FPGA=", "Exp=")
         VersionString$ = VersionString$ & CurItem & Left(ConfigVal, ConfigLen) & ", "
      End If
   Next
   VersionString$ = Left(VersionString$, (Len(VersionString$) - 2))
   lblFirmwareVersion.Caption = VersionString$
   
End Sub

Private Function CheckGPIBError() As Boolean

   Dim GPIBErrStat As Boolean
   
   GPIBErrStat = False
   If (ibsta And EERR) = EERR Then
      Dim ErrorMessage As String, StatString As String
      Dim ErrorString As String
      tmrCheckStatus.Enabled = False
      StatString = ParseStatus()
      ErrorString = GetErrorMessage(ErrorMessage)
      txtRateLog.Text = txtRateLog.Text & _
         vbCrLf & vbCrLf & vbTab & "GPIB error: " & ErrorString
      GPIBErrStat = True
      cmdStop_Click
   End If
   CheckGPIBError = GPIBErrStat

End Function

Private Sub tmrDelay_Timer()

   tmrDelay.Enabled = False
   
End Sub

Private Function LoadData() As Long

   Dim x As Integer, LongData As Long
   Dim Element As Integer, Remdr As Integer
   
   For Element = 0 To nlCount - 1
      Remdr = Element Mod 2
      If Remdr = 0 Then
         LongData = 2 ^ mnResolution / 2 - 2 ^ mnResolution
      Else
         LongData = 2 ^ mnResolution / 2 - Remdr
      End If
      mlDataBuffer(Element) = LongData
   Next

   ULStat& = cbWinArrayToBuf(mlDataBuffer(0), mlMemHandle, 0, nlCount)
   LoadData = ULStat&
   
End Function
