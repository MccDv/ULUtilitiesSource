VERSION 5.00
Begin VB.Form frmULBoards 
   Caption         =   "List UL Boards"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCaps 
      Caption         =   "All Caps"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdSortBoards 
      Caption         =   "Sort"
      Enabled         =   0   'False
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   1035
   End
   Begin VB.TextBox txtBoardList 
      Height          =   3495
      Left            =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   540
      Width           =   3615
   End
   Begin VB.CommandButton cmdListBoards 
      Caption         =   "List Boards"
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lblNumBoards 
      Height          =   195
      Left            =   1500
      TabIndex        =   3
      Top             =   180
      Width           =   2955
   End
End
Attribute VB_Name = "frmULBoards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msBoardList()
Dim mlMaxBoards As Long
Dim mlNumBoards As Long

Private Sub chkCaps_Click()

   Me.txtBoardList.Text = ""
   Me.cmdSortBoards.Enabled = False
   
End Sub

Private Sub cmdListBoards_Click()

   ReDim msBoardList(mlMaxBoards)
   NumBoards& = -1
   BoardName$ = Space$(BOARDNAMELEN)
   ULStat& = cbGetBoardName(GETFIRST, BoardName$)
   If ULStat& = 126 Then
      txtBoardList.Text = " You shouldn't have to, but" _
         & vbCrLf & " you need to run Instacal first."
      Exit Sub
   End If
   If Len(BoardName$) Then
      'Drop the space characters
      BoardName$ = RTrim$(BoardName$)
      StringSize% = Len(BoardName$)
      'lop off the null
      If StringSize% > 0 Then BoardName$ = Left$(BoardName$, StringSize% - 1)
      NumBoards& = NumBoards& + 1
      If chkCaps.Value Then
         msBoardList(NumBoards&) = UCase(BoardName$)
      Else
         msBoardList(NumBoards&) = BoardName$
      End If
   End If
   Do
      BoardName$ = Space$(BOARDNAMELEN)
      ULStat = cbGetBoardName(GETNEXT, BoardName$)
      If Len(BoardName$) Then
         NumBoards& = NumBoards& + 1
         'Drop the space characters
         BoardName$ = RTrim$(BoardName$)
         StringSize% = Len(BoardName$)
         'lop off the null
         If StringSize% > 0 Then BoardName$ = Left$(BoardName$, StringSize% - 1)
         If chkCaps.Value Then
            msBoardList(NumBoards&) = UCase(BoardName$)
         Else
            msBoardList(NumBoards&) = BoardName$
         End If
      End If
   Loop While Len(BoardName$) > 3
   mlNumBoards = NumBoards& - 1
   ReDim Preserve msBoardList(mlNumBoards)
   
   Me.txtBoardList.Text = ""
   For ThisBoard& = 0 To mlNumBoards
      Me.txtBoardList.Text = Me.txtBoardList.Text & msBoardList(ThisBoard&) & vbCrLf
   Next ThisBoard&
   If Not (mlNumBoards < 0) Then
      cmdSortBoards.Enabled = True
      lblNumBoards.Caption = mlNumBoards & " boards in list."
   Else
      cmdSortBoards.Enabled = False
      lblNumBoards.Caption = "No boards in list."
   End If
   
End Sub

Private Sub cmdSortBoards_Click()
   
   Me.txtBoardList.Text = ""
   QuickSortVariants msBoardList, 0, mlNumBoards
   For ThisBoard& = 0 To mlNumBoards
      Me.txtBoardList.Text = Me.txtBoardList.Text & msBoardList(ThisBoard&) & vbCrLf
   Next ThisBoard&

End Sub

Private Sub Form_Load()

   mlMaxBoards = 400
   
End Sub
