VERSION 5.00
Begin VB.Form frmScriptFiles 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Script Files"
   ClientHeight    =   2985
   ClientLeft      =   4050
   ClientTop       =   4650
   ClientWidth     =   7410
   ControlBox      =   0   'False
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
   HelpContextID   =   503
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbPattern 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   2520
      Width           =   2475
   End
   Begin VB.Frame fraSearch 
      BackColor       =   &H80000005&
      Caption         =   "Search for Compatible Scripts"
      Height          =   3255
      Left            =   180
      TabIndex        =   10
      Top             =   2220
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Height          =   315
         HelpContextID   =   502
         Left            =   5820
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkSubs 
         BackColor       =   &H80000005&
         Caption         =   "Search subdirectories"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   900
         Width           =   2355
      End
      Begin VB.ComboBox cmbProducts 
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find Compatible Scripts"
         Height          =   315
         HelpContextID   =   502
         Left            =   3180
         TabIndex        =   12
         Top             =   360
         Width           =   2475
      End
      Begin VB.ListBox lstScripts 
         Height          =   1815
         ItemData        =   "ScrFiles.frx":0000
         Left            =   180
         List            =   "ScrFiles.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   11
         Top             =   1260
         Width           =   6735
      End
      Begin VB.Label lblNumFound 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   6300
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   6300
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   180
      TabIndex        =   2
      Top             =   540
      Width           =   2475
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1605
      Left            =   2820
      TabIndex        =   1
      Top             =   540
      Width           =   3315
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   180
      Width           =   1755
   End
   Begin VB.TextBox txtPattern 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Text            =   "*.uts;*.utm"
      Top             =   180
      Width           =   2475
   End
   Begin VB.TextBox txtHeader 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   5955
   End
   Begin VB.Label lblDrives 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Drives:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblTypes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "List files of type:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "frmScriptFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mnFileExists As Integer
Dim mnSaveAs As Integer
Dim msAsFile As String

Private WithEvents dd As CDirDrill
Attribute dd.VB_VarHelpID = -1
Private m_Cancel As Boolean

Private Sub cmbPattern_Change()

   filespec$ = cmbPattern.Text
   txtPattern.Text = cmbPattern.Text
   File1.Pattern = filespec$
   cmdOK.Enabled = (Left$(filespec$, 1) <> "*") And ((Right$(LCase$(filespec$), 4) = ".uts") Or (Right$(LCase$(filespec$), 4) = ".utm"))

End Sub

Private Sub cmbPattern_Click()

   CommentLoc& = InStr(cmbPattern.Text, " (") - 1
   If CommentLoc& > 1 Then
      filespec$ = Left$(cmbPattern.Text, InStr(cmbPattern.Text, " (") - 1)
   Else
      filespec$ = cmbPattern.Text
   End If
   txtPattern.Text = filespec$
   File1.Pattern = filespec$
   cmdOK.Enabled = (Left$(filespec$, 1) <> "*") And ((Right$(LCase$(filespec$), 4) = ".uts") Or (Right$(LCase$(filespec$), 4) = ".utm"))

End Sub

Private Sub cmdCancel_Click()

   Me.txtHeader.Text = ""
   txtPattern.Text = ""
   Me.Hide

End Sub

Private Sub cmdCopy_Click()

   NumItems& = lstScripts.ListCount - 1
   Clipboard.Clear
   For Item& = 0 To NumItems&
      If lstScripts.Selected(Item&) Then
         ClipText$ = ClipText$ & lstScripts.List(Item&) & vbCrLf
      End If
   Next Item&
   Clipboard.SetText ClipText$
   
End Sub

Private Sub cmdFind_Click()

   Dim SubDirectories() As String
   
   cmdFind.Enabled = False
   lstScripts.Clear
   Me.lblNumFound.Caption = ""

   NumDirs& = 0
   ScriptDir$ = Dir1.Path
   
   If chkSubs.Value Then
      dd.Folder = ScriptDir$
      dd.Pattern = Me.txtPattern.Text
      dd.AttributeMask = vbHidden Or vbSystem Or vbArchive Or vbReadOnly
      ' Clear cancel flag.
      m_Cancel = False
      ' Let it rip!
      dd.BeginSearch
      Exit Sub
   End If
   
   ScriptDir$ = Dir1.Path & "\"
   File1.Path = ScriptDir$
   GoSub FindFiles
   cmdFind.Enabled = True
   Comatibles& = Me.lstScripts.ListCount
   Me.lblNumFound.Caption = "Found " & Comatibles& & " compatible files."
   Exit Sub
   
FindFiles:
   NumFiles& = File1.ListCount - 1
   For FileNum& = 0 To NumFiles&
      Filename$ = File1.List(FileNum&)
      ScriptPath$ = ScriptDir$ & Filename$
      FileStatus% = EvaluateScript(ScriptPath$)
      If FileStatus% = 2 Then
         Me.lstScripts.AddItem ScriptPath$
      End If
   Next FileNum&
   Return
   
End Sub

Private Sub cmdOK_Click()

   If Me.txtHeader.Visible Then
      Me.Hide
      Exit Sub
   End If
   If mnFileExists Then
      Resp = MsgBox("Replace " & File1.Filename & "?", 4, "Replace Existing File?")
      If Resp = 7 Then
         txtPattern.Text = "*.uts"
         mnFileExists = False
         Exit Sub
      Else
         Me.Hide
         Exit Sub
      End If
   End If
   If Me.Caption = "Open New Script" Then
      For ListNum% = 0 To File1.ListCount - 1
         If txtPattern.Text = File1.List(ListNum%) Then
            Resp = MsgBox("Replace " & File1.Filename & "?", 4, "Replace Existing File?")
            If Resp = 7 Then
               txtPattern.Text = "*.uts"
               Exit Sub
            End If
         End If
      Next ListNum%
   End If
   Me.Hide

End Sub

Private Sub Dir1_Change()

   File1.Path = Dir1.Path
   If mnSaveAs Then txtHeader.Text = Dir1.Path & "\" & msAsFile

End Sub

Private Sub Dir1_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Then File1.Path = Dir1.Path
   If mnSaveAs Then txtHeader.Text = Dir1.Path & "\" & msAsFile

End Sub

Private Sub Drive1_Change()

   On Error GoTo BadDrive
   Dir1.Path = Drive1.Drive
   If mnSaveAs Then txtHeader.Text = Dir1.Path & "\" & msAsFile
   Exit Sub

BadDrive:
   MsgBox "Select another drive."
   Resume Next

End Sub

Private Sub File1_Click()

   If (Me.Caption = "Open New Script") Or mnSaveAs Then mnFileExists = True
   txtPattern.Text = File1.Filename
   If mnSaveAs Then txtHeader.Text = Dir1.Path & "\" & msAsFile

End Sub

Private Sub File1_DblClick()

   If (Me.Caption = "Open New Script") Or mnSaveAs Then
      Resp = MsgBox("Replace " & File1.Filename & "?", 4, "Replace Existing File?")
      If Resp = 7 Then Exit Sub
   End If
   txtPattern.Text = File1.Filename
   If mnSaveAs Then txtHeader.Text = Dir1.Path & "\" & msAsFile
   Me.Hide

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   ' Set cancel flag if Escape key was pressed.
   m_Cancel = (KeyAscii = vbKeyEscape)

End Sub

Private Sub Form_Load()

   File1.Pattern = txtPattern.Text  'txtPattern.Text
   mnFileExists = False

End Sub

Private Sub Form_Activate()
   
   Me.File1.SetFocus
   CurCaption$ = Me.Caption
   Select Case CurCaption$
      Case "Save Script As"
         Me.txtHeader.Text = Me.Dir1.Path & "\NewScript.uts"
         Me.cmdOK.Enabled = True
         mnSaveAs = True
         msAsFile = "NewScript.uts"
         Me.txtPattern.Text = msAsFile
      Case "Find Compatible Scripts"
         Me.Height = 6105
         cmbPattern.Visible = False
         Me.fraSearch.Visible = True
         Set dd = New CDirDrill
         File1.Pattern = txtPattern.Text
         FillProductList
         ' Enable watching for an escape key.
         Me.KeyPreview = True
         Me.HelpContextID = 502
      Case "Save Data As Text"
         Me.txtHeader.Text = Me.Dir1.Path & "\DataFile.txt"
         Me.cmdOK.Enabled = True
         mnSaveAs = True
         msAsFile = "DataFile.txt"
         Me.txtPattern.Text = msAsFile
   End Select
   
End Sub

Private Sub lstScripts_Click()

   Me.txtHeader.Text = lstScripts.Text
   Me.cmdOK.Enabled = True
   Me.cmdCopy.Enabled = lstScripts.SelCount > 0
   
End Sub

Private Sub lstScripts_DblClick()

   Me.txtHeader.Text = lstScripts.Text
   Me.cmdOK.Enabled = True
   Me.cmdOK.Value = True

End Sub

Private Sub txtPattern_Change()

   filespec$ = txtPattern.Text
   cmdOK.Enabled = (Left$(filespec$, 1) <> "*") 'And ((Right$(LCase$(filespec$), 4) = ".uts") Or (Right$(LCase$(filespec$), 4) = ".utm"))
   If mnSaveAs Then
      msAsFile = txtPattern.Text
      Me.txtHeader.Text = Me.Dir1.Path & "\" & msAsFile
   End If

End Sub

Private Sub txtPattern_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      filespec$ = txtPattern.Text
      File1.Pattern = filespec$
      If mnSaveAs Then
         msAsFile = txtPattern.Text
         txtHeader.Text = Dir1.Path & "\" & msAsFile
      End If
   End If

End Sub

Sub FillProductList()
   
   ListFile$ = GetBoardFile()
   If Not ListFile$ = "" Then
      Open ListFile$ For Input As #4
      Do While Not EOF(4)
         Line Input #4, A1$
         cmbProducts.AddItem A1$
      Loop
      Close #4
      If cmbProducts.ListCount > 122 Then cmbProducts.ListIndex = 122
   End If
   
End Sub

Function EvaluateScript(ScriptPath As String) As Integer

   Open ScriptPath For Input As #4

   ProdSearch$ = UCase(Me.cmbProducts.Text)
   If LCase(Right(ProdSearch$, 3)) = "uts" Then
      'search for master scripts calling subscript
      If LCase(Right(ScriptPath, 3)) = "utm" Then
         FindSubscript% = True
         CompListFound% = True
      End If
   ElseIf LCase(Right(ProdSearch$, 3)) = "utm" Then
      'search for master scripts calling subscript
      FindSubscript% = True
      CompListFound% = True
   Else
      Do While Not EOF(4)
         Line Input #4, A1
         If A1 = "'Compatibility list" Then
            CompListFound% = True
            EvaluateScript = 1   'compatibility list found
            Exit Do
         End If
      Loop
   End If
   
   If CompListFound% Then
      Do While Not EOF(4)
         Line Input #4, A1
         CapName$ = UCase(A1)
         If Not FindSubscript% Then
            If InStr(1, CapName$, "ALL DEVICES") Then
               ProdFound% = True
               EvaluateScript = 2   'all products are compatible
               Exit Do
            ElseIf InStr(1, CapName$, "ALL EXCEPT") Then
               SearchException% = True
            End If
         End If
         'Else
            If InStr(1, CapName$, ProdSearch$) Then
               If SearchException% Then
                  'board listed is not compatible
                  Exit Do
               Else
                  ProdFound% = True
                  EvaluateScript = 2   'product listed as compatible
                  Exit Do
               End If
            End If
         'End If
      Loop
   Else
      EvaluateScript = 0   'no compatibility list in file
   End If
   Close #4

End Function

Private Sub dd_Done(ByVal TotalFiles As Long, ByVal TotalFolders As Long)
   
   Me.cmdFind.Enabled = True
   Comatibles& = Me.lstScripts.ListCount
   Me.lblNumFound.Caption = "Found " & Comatibles& & " compatible files."

End Sub

Private Sub dd_NewFile(ByVal filespec As String, Cancel As Boolean)

   'FileName$ = dd.ExtractName(filespec)
   FileStatus% = EvaluateScript(filespec)
   If FileStatus% = 2 Then Me.lstScripts.AddItem filespec

End Sub

Private Sub dd_NewFolder(ByVal FolderSpec As String, Cancel As Boolean)

   ' Take a breath, bail if Escape was pressed.
   DoEvents
   Cancel = m_Cancel

End Sub

