VERSION 5.00
Begin VB.Form frmFileDialog 
   Caption         =   "Save File"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5370
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Text            =   "RateLog.txt"
      Top             =   1860
      Width           =   1875
   End
   Begin VB.CommandButton cmdSelect 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4140
      TabIndex        =   4
      Top             =   1860
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      Top             =   1860
      Width           =   1035
   End
   Begin VB.FileListBox flbRateLog 
      Height          =   1455
      Left            =   2400
      Pattern         =   "*.txt"
      TabIndex        =   2
      Top             =   180
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   1875
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1875
   End
End
Attribute VB_Name = "frmFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

   Me.txtFileName.Text = ""
   Me.Hide
   
End Sub

Private Sub cmdSelect_Click()

   Me.Hide
   
End Sub

Private Sub Dir1_Change()

   Me.flbRateLog.Path = Me.Dir1.Path
   
End Sub

Private Sub Drive1_Change()

   Me.Dir1.Path = Me.Drive1.Drive
   
End Sub

Private Sub flbRateLog_Click()

   Me.txtFileName.Text = Me.flbRateLog.Filename
   
End Sub

Private Sub Form_Load()

   Dim KeyFound As Boolean
   Dim RegNode As String
   Dim Value As String
   
   Value = "Personal"
   RegNode = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
   KeyFound = GetKeyValue(HKEY_CURRENT_USER, RegNode, Value)
   If KeyFound Then
      Me.Drive1.Drive = Value
      Me.Dir1.Path = Value
   End If
   Me.flbRateLog.Refresh
   
End Sub
