Attribute VB_Name = "KSMain"

Sub Main()

   CmdString$ = Command()
   If Not (CmdString$ = "") Then
      CmdSegments = Split(CmdString$, ";")
      NumSegs& = UBound(CmdSegments)
      FileSpec$ = CmdSegments(0)
      Quotes& = FindInString(FileSpec$, Chr(34), Locs)
      If Quotes& = 0 Then
         If Locs(0) = 1 Then
            FileSpec$ = Mid(FileSpec$, 2)
         Else
            FileSpec$ = Left(FileSpec$, Quotes& - 1)
         End If
      Else
         If Not Quotes& < 0 Then FileSpec$ = Mid(FileSpec$, Locs(0) + 1, Locs(1) - (Locs(0) + 1))
      End If
      'frmKeySender.lblCurFile.Caption = FileSpec$
      If Not (InStr(1, FileSpec$, ".") = 0) Then
         frmKeySender.SetFileCommand FileSpec$
      End If
      For Segment& = 1 To NumSegs&
         CurCommand$ = LCase(CmdSegments(Segment&))
         Select Case CurCommand$
            Case "autostart"
               frmKeySender.SetAutoStart
            Case "autostop"
               frmKeySender.SetAutoStop
         End Select
      Next
   End If
   frmKeySender.Show
   DoEvents
   
End Sub
