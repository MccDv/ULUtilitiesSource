VERSION 5.00
Begin VB.Form frmRemoteNetDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remote Network Parameters"
   ClientHeight    =   1110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   3060
      TabIndex        =   6
      Text            =   "5000"
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtHostPort 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "54211"
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtHostName 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "173.76.198.250"
      Top             =   180
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5460
      TabIndex        =   1
      Top             =   600
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   5460
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblTimeout 
      Caption         =   "Timeout"
      Height          =   195
      Left            =   4140
      TabIndex        =   7
      Top             =   660
      Width           =   1035
   End
   Begin VB.Label lblHostPort 
      Caption         =   "Host port"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   660
      Width           =   1035
   End
   Begin VB.Label lblHostName 
      Caption         =   "Host name or IP address"
      Height          =   195
      Left            =   3060
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmRemoteNetDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()

   txtHostName.Text = ""
   
End Sub

Private Sub cmdOK_Click()

   Dim haveParams As Boolean
   
   haveParams = (Len(txtHostName.Text) > 3) _
      And (Len(txtHostPort.Text) > 3) _
      And (Len(txtTimeout.Text) > 0)
   Me.Hide

End Sub
