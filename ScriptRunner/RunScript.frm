VERSION 5.00
Begin VB.Form frmScriptRun 
   Caption         =   "Script Runner"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   Icon            =   "RunScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3660
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtTimoutDelay 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "20"
      Top             =   2460
      Width           =   555
   End
   Begin VB.TextBox txtScript 
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Text            =   "DIOScan.utm"
      Top             =   1020
      Width           =   4635
   End
   Begin VB.TextBox txtLoadDelay 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "4"
      Top             =   2100
      Width           =   555
   End
   Begin VB.TextBox txtAppPath 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Text            =   "utul32.exe"
      Top             =   360
      Width           =   4635
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2640
      Top             =   3060
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Script"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label lblApp 
      Caption         =   "Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label lblStatus 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   1560
      Width           =   4155
   End
   Begin VB.Label Label1 
      Caption         =   "seconds - Script timeout"
      Height          =   255
      Left            =   780
      TabIndex        =   8
      Top             =   2520
      Width           =   3555
   End
   Begin VB.Label lblRunDelay 
      Caption         =   "seconds - Delay before script Run command"
      Height          =   255
      Left            =   780
      TabIndex        =   7
      Top             =   2160
      Width           =   3555
   End
End
Attribute VB_Name = "frmScriptRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvReslt As Variant, mnState As Integer
Dim msScriptDef As String
Dim mlIterations As Long

Private Sub cmdExit_Click()

   End
   
End Sub

Private Sub cmdStart_Click()
   
   Me.lblStatus.Caption = ""
   mlIterations = 0
   If Not (mvReslt = Empty) Then
      CancelRestart
   End If
   mnState = 0
   Me.Timer1.Interval = Val(txtLoadDelay.Text) * 1000
   Me.Timer1.Enabled = True
   Me.cmdStart.Enabled = False
   Me.cmdStop.Enabled = True
   GivenPath = Me.txtAppPath.Text
   If (InStr(GivenPath, "\") = 0) Then AssumedPath$ = CurDir() & "\"
   AppPath = AssumedPath$ & GivenPath
   GivenPath = Me.txtScript.Text
   If (InStr(GivenPath, "\") = 0) Then AssumedPath$ = CurDir() & "\"
   ScriptPath = AssumedPath$ & GivenPath
   msScriptDef = AppPath & " " & ScriptPath

End Sub

Private Sub cmdStop_Click()

   CancelRestart
   
End Sub

Private Sub Timer1_Timer()

   Select Case mnState
      Case 0
         'open the app
         mvReslt = Shell(msScriptDef, vbNormalFocus)
         mnState = 1
         mlIterations = mlIterations + 1
      Case 1
         'run the script
         AppActivate mvReslt, False
         DoEvents
         SendKeys "%R", True
         mnState = 2
      Case 2
         'wait for the app to close or close it if timeout
         t0! = Timer()
         Handle& = OpenProcess(SYNCHRONIZE, 0, mvReslt)
         DelayTime! = Val(txtTimoutDelay.Text)
         Milliseconds& = DelayTime! * 1000
         Result& = WaitForSingleObject(Handle&, Milliseconds&)
         Me.lblStatus.Caption = "Iterations completed = " & Format(mlIterations, "0")
         Et! = (Timer() - t0!) + 1
         If Not (Et! < DelayTime!) Then
            'CancelRestart
            mnState = 3
            Me.SetFocus
            MsgBox "Application did not self-terminate. Restart cancelled.", _
            vbInformation, "Restart Cancelled"
            Me.Timer1.Enabled = False
            Me.cmdStart.Enabled = False
            Me.cmdStop.Enabled = False
         Else
            mnState = 0
            mvReslt = Empty
         End If
   End Select
   
End Sub

Sub CancelRestart()
   
   Me.Timer1.Enabled = False
   Me.cmdStart.Enabled = True
   Me.cmdStop.Enabled = False
   If Not (mvReslt = Empty) Then
      AppActivate mvReslt, False
      SendKeys "{ESC}", True
      DoEvents
      AppActivate mvReslt, False
      SendKeys "{ESC}", True
      SendKeys "%{F4}", True
      DoEvents
      mvReslt = Empty
   End If

End Sub
