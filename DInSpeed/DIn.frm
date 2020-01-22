VERSION 5.00
Begin VB.Form frmIn 
   Caption         =   "Digital Out Speed"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLastBit 
      Height          =   225
      Left            =   3060
      TabIndex        =   12
      Text            =   "-1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtFirstBit 
      Height          =   225
      Left            =   3060
      TabIndex        =   11
      Text            =   "0"
      Top             =   1020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Bit Input"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   1020
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Port Input"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1020
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox chkArray 
      Caption         =   "Use DOutArray"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2460
      TabIndex        =   8
      Top             =   180
      Width           =   2055
   End
   Begin VB.TextBox txtRateEstimate 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Text            =   "1"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtPortIndex 
      Height          =   285
      Left            =   300
      TabIndex        =   3
      Text            =   "0"
      Top             =   540
      Width           =   495
   End
   Begin VB.TextBox txtBoardNum 
      Height          =   285
      Left            =   300
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label lblLastBit 
      Caption         =   "Last Bit"
      Height          =   195
      Left            =   3600
      TabIndex        =   14
      Top             =   1380
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblFirstBit 
      Caption         =   "First Bit"
      Height          =   195
      Left            =   3600
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblRateEstimate 
      Caption         =   "Iterations"
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label lblPortType 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2460
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblPortIndex 
      Caption         =   "Port Index"
      Height          =   195
      Left            =   1020
      TabIndex        =   4
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lblBoardNum 
      Caption         =   "Board Number"
      Height          =   195
      Left            =   1020
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlBoardNum As Long, mlNumBits As Long
Dim mlFirstBit As Long, mlLastBit As Long
Dim mlPortNum As Long, mlNumArrayPorts As Long
Dim mbDoBits As Boolean

Private Sub chkArray_Click()

   Dim PortEnabled As Boolean
   
   PortEnabled = Not (chkArray.Value = 0)
   txtPortIndex.Text = "0"
   txtPortIndex.Enabled = (chkArray.Value = 0)
   If PortEnabled Then
      ULStat& = cbGetConfig(BOARDINFO, mlBoardNum, _
         0, BIDINUMDEVS, ConfigVal&)
      For DevNum& = 0 To ConfigVal& - 1
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
      PortTypeString$ = "Port Type "
      If mlNumArrayPorts > 1 Then
         For PortNum& = mlPortNum To mlPortNum + (mlNumArrayPorts - 1)
            PortTypeString$ = PortTypeString$ & _
               Format(PortNum&, "0") & ", "
         Next
      End If
      PortTypeString$ = Left(PortTypeString$, Len(PortTypeString$) - 2)
   Else
      PortTypeString$ = "Port Type " & _
         Format(mlPortNum, "0")
   End If
   Me.lblPortType.Caption = PortTypeString$

End Sub

Private Sub Command1_Click()

   Dim FirstBit As Long, LastBit As Long
   
   Offset& = GetBitOffset()
   FirstBit = Val(txtFirstBit.Text) + Offset&
   LastBit = Val(txtLastBit.Text) + Offset&
   Me.Command1.Enabled = False
   Iterations& = Val(txtRateEstimate.Text)
   
   ULStat& = cbDConfigPort(mlBoardNum, mlPortNum, DIGITALIN)
   PortNum& = mlPortNum
   If mlPortNum > 10 Then PortNum& = FIRSTPORTA
   If mbDoBits Then
      For I& = 1 To Iterations&
         For CurBit& = FirstBit To LastBit
            ULStat& = cbDBitIn(mlBoardNum, PortNum&, CurBit&, BitVal%)
         Next
      Next
   Else
      DataVal% = 0
      ULStat& = cbDIn(mlBoardNum, mlPortNum, DataVal%)
      For I& = 1 To Iterations&
         ULStat& = cbDIn(mlBoardNum, mlPortNum, DataVal%)
      Next
   End If
   Me.Command1.Enabled = True
   
End Sub

Private Sub Form_Load()

   ULStat& = cbErrHandling(DONTPRINT, DONTSTOP)
   mlBoardNum = 0
   BoardName$ = GetDevName()
   If Not BoardName$ = "" Then
      GetPortType 0
      Command1.Enabled = True
   End If
   
End Sub

Private Sub optPort_Click(Index As Integer)

   mbDoBits = optPort(1).Value
   txtFirstBit.Visible = mbDoBits
   txtLastBit.Visible = mbDoBits
   lblFirstBit.Visible = mbDoBits
   lblLastBit.Visible = mbDoBits
   GetPortType 0
   
End Sub

Private Sub txtBoardNum_Change()

   mlBoardNum = Val(Me.txtBoardNum.Text)
   BoardName$ = GetDevName()
   If Not BoardName$ = "" Then
      Me.txtPortIndex.Text = "0"
      GetPortType 0
   Else
      Command1.Enabled = False
   End If
   
End Sub

Private Function GetDevName() As String

   BoardName$ = Space$(BOARDNAMELEN)
   ULStat& = cbGetBoardName(mlBoardNum, BoardName$)
   BoardName$ = Trim(BoardName$)
   BoardName$ = Left(BoardName$, Len(BoardName$) - 1)
   If BoardName$ = "" Then
      Me.Caption = "Invalid Device"
   Else
      Me.Caption = BoardName$
   End If
   GetDevName = BoardName$

End Function

Private Sub GetPortType(ByVal PortIndex As Long)

   Dim OverBit As Boolean
   
   ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
   PortIndex, DIDEVTYPE, DevType&)
   If Not ULStat& = 0 Then
      Me.lblPortType.Caption = "Invalid Port"
      Command1.Enabled = False
   Else
      mlPortNum = DevType&
      Me.lblPortType.Caption = "Port Type " & _
      Format(mlPortNum, "0")
      Command1.Enabled = True
   End If
   
   ULStat& = cbGetConfig(DIGITALINFO, mlBoardNum, _
      PortIndex, DINUMBITS, NumBits&)
   CurLast& = Val(txtLastBit.Text)
   OverBit = Not (CurLast& < NumBits&)
   If (CurLast& < 0) Or OverBit Then
      mlLastBit = NumBits& - 1
      txtLastBit.Text = Format(mlLastBit, "0")
   End If
   CurFirst& = Val(txtFirstBit.Text)
   If (CurFirst& > mlLastBit) Then
      mlFirstBit = 0
      txtFirstBit.Text = "0"
   End If
   
End Sub

Private Sub txtPortIndex_Change()

   PortIndex& = Val(Me.txtPortIndex.Text)
   GetPortType PortIndex&
   
End Sub

Private Sub ArrayLoop()
   
   Dim OutArrayLow(1) As Long
   Dim OutArrayHigh(1) As Long
   
   Me.Command1.Enabled = False
   Iterations& = Val(txtRateEstimate.Text) * 5
   
   ULStat& = cbDConfigPort(mlBoardNum, mlPortNum, DIGITALIN)
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
      For I& = 0 To Iterations&
         ULStat& = cbDOutArray(mlBoardNum, mlPortNum, _
            LastArrayPort&, OutArrayHigh(0))
         ULStat& = cbDOutArray(mlBoardNum, mlPortNum, _
            LastArrayPort&, OutArrayLow(0))
      Next
   End If
   
End Sub

Function GetBitOffset() As Long

   Select Case mlPortNum
      Case 0, 10
         Offset& = 0
      Case Is > 10
         Offset& = 8
         For CurPort& = 12 To mlPortNum
            Select Case CurPort&
               Case 12, 13, 16, 17, 20, 21, 24, 25
                  Offset& = Offset& + 4
               Case 14, 15, 18, 19, 22, 23, 26, 27
                  Offset& = Offset& + 8
               Case 28, 29, 32, 33, 36, 37, 40, 41
                  Offset& = Offset& + 4
               Case 30, 31, 34, 35, 38, 39
                  Offset& = Offset& + 8
            End Select
         Next
   End Select
   GetBitOffset = Offset&
   
End Function
