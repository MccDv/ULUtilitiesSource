VERSION 5.00
Begin VB.Form frmKeySender 
   Caption         =   "Key Sender"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "KeySender.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Pause"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   180
      Width           =   1035
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2400
      OLEDropMode     =   1  'Manual
      Pattern         =   "*.sdk"
      TabIndex        =   3
      Top             =   780
      Width           =   3075
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1995
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   1995
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3780
      Top             =   240
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2460
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lblCurLine 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   5235
   End
   Begin VB.Label lblCurFile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3420
      Width           =   5235
   End
End
Attribute VB_Name = "frmKeySender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvReslt As Variant
Dim masKeyLines() As String
Dim msKeyFile As String
Dim mnNumKeyLines As Integer, mnCurKeyLine As Integer
Dim mnAppDelay As Integer, mnAutoStop As Integer, mnAutoStart As Integer

Private Sub cmdCancel_Click()

   Timer1.Enabled = False
   Me.cmdStart.Enabled = True
   
End Sub

Private Sub cmdStart_Click()

   On Error GoTo NotFound
   
   If Not (mvReslt = Empty) Then
      AppActivate mvReslt, False
      SendKeys "%{F4}", True
   End If
   If mnCurKeyLine = 0 Then
      AppPath = masKeyLines(mnCurKeyLine)
      If (InStr(1, AppPath, ":") = 0) Then
         'incomplete path - use program files directory
         BasePath$ = Environ("programfiles")
         If Not Left(AppPath, 1) = "\" Then BSlash$ = "\"
         ProgPath$ = BasePath$ & BSlash$ & AppPath
      Else
         ProgPath$ = AppPath
      End If
      mvReslt = Shell(ProgPath$, vbNormalFocus)
      If mnNumKeyLines > 1 Then
         DelSpec$ = masKeyLines(1)
         If LCase(Left(DelSpec$, 6)) = "delay " Then
            DelVal$ = Mid(DelSpec$, 7)
            mnAppDelay = Val(DelVal$)
            mnCurKeyLine = 1
            Me.lblCurLine.Caption = "1)  " & DelSpec$
         Else
            mnAppDelay = 2
         End If
      End If
   End If
   Me.cmdStart.Enabled = False
   Me.Timer1.Enabled = True
   Exit Sub
   
NotFound:
   MsgBox Error(Err) & " (" & ProgPath$ & ")", vbCritical, "Error Opening Application"
   Exit Sub

End Sub

Private Sub Dir1_Change()

   File1.Path = Dir1.Path
   
End Sub

Private Sub Drive1_Change()

   Dir1.Path = Drive1.Drive
   
End Sub

Private Sub File1_DblClick()

   msKeyFile = File1.Path & "\" & File1.FileName
   OpenKeyFile
   
End Sub

Private Sub File1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Data.Files.Count > 0 Then
      msKeyFile = Data.Files(1)
      OpenKeyFile
   End If

End Sub

Private Sub Form_Activate()
   
   If mnAutoStart Then
      Me.lblCurLine = "Autostart set"
      Me.cmdStart = True
   End If

End Sub

Private Sub Form_Load()

   Me.Top = Screen.Height - Me.Height * 1.3
   Me.Left = Screen.Width - Me.Width

End Sub

Private Sub Timer1_Timer()

   If mnAppDelay > 0 Then
      mnAppDelay = mnAppDelay - 1
      Me.lblCurLine.Caption = "1) Delay  " & Format(mnAppDelay, "0")
      Exit Sub
   End If
   mnCurKeyLine = mnCurKeyLine + 1
   If mnCurKeyLine > mnNumKeyLines Then
      Timer1.Enabled = False
      Me.cmdStart.Enabled = True
      Me.lblCurLine.Caption = ""
      mnCurKeyLine = 0
      mvReslt = Empty
      If mnAutoStop Then End
      Exit Sub
   End If
   KeyLine$ = masKeyLines(mnCurKeyLine)
   If Left$(KeyLine$, 1) = "'" Then Exit Sub
   Me.lblCurLine.Caption = Format(mnCurKeyLine, "0") & ")  " & KeyLine$
   ParsedLine = Split(KeyLine$, ";")
   Elements& = UBound(ParsedLine)
   CurKey$ = ParsedLine(0)
   AppActivate mvReslt, False
   SendKeys CurKey$, True
   DoEvents
   If Elements& > 0 Then Repeater& = Val(ParsedLine(1))
   For RepeatLine& = 2 To Repeater&
      AppActivate mvReslt, False
      SendKeys CurKey$, True
      For i& = 0 To 10000
         DoEvents
      Next
   Next
      'Do
      '   Sleep (1000)
      'While AppIsRunning(mvResult)
      'mvReslt = Empty
      'For Delay& = 0 To 10000
      '   DoEvents
      'Next Delay&
      'mvReslt = Shell(AppPath, vbNormalFocus)
   
End Sub

Private Sub OpenKeyFile()

   NumSubs& = FindInString(msKeyFile, "\", Locations)
   Dir1.Path = Left(msKeyFile, Locations(NumSubs&))
   File1.Refresh
   Open msKeyFile For Input As #1
   mnNumKeyLines = 0
   mnCurKeyLine = 0
   ReDim Preserve masKeyLines(mnNumKeyLines)
   Do While Not EOF(1)
      Line Input #1, A1
      If A1 = "" Then Exit Do 'A1 = "; "
      ReDim Preserve masKeyLines(mnNumKeyLines)
      masKeyLines(mnNumKeyLines) = A1
      mnNumKeyLines = mnNumKeyLines + 1
      'Me.List1.AddItem A1
   Loop
   Close #1
   mnNumKeyLines = mnNumKeyLines - 1
   If mnNumKeyLines > 0 Then
      Me.cmdStart.Enabled = True
      Me.lblCurFile.Caption = msKeyFile
   Else
      Me.lblCurFile.Caption = ""
   End If

End Sub

Public Sub SetFileCommand(FilePath As String)

   msKeyFile = FilePath
   lblCurFile.Caption = msKeyFile
   OpenKeyFile

End Sub

Public Sub SetAutoStop()

   mnAutoStop = True
   
End Sub

Public Sub SetAutoStart()

   mnAutoStart = True
   
End Sub

