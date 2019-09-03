Attribute VB_Name = "ULDiscovery"
Dim mfNoForm As Form
Dim mlNumDevices As Long
Dim mxDescriptors() As DaqDeviceDescriptor, mnNumDescriptors As Integer
Dim mxDemoDescriptors() As DaqDeviceDescriptor
Dim mnNumDemoDescriptors As Integer

Public Function DiscoverDevices(ByVal Interface As DaqDeviceInterface, _
   Optional CreateDev As Boolean, Optional HostString As Variant, Optional HostPort As Long, _
   Optional Timeout As Long, Optional AddDemo As Boolean) As Long

   Dim NamesAdded As String, NumsAdded As String
   Dim UIDStrings() As String, NewName As String
   Dim Inventory(100) As DaqDeviceDescriptor
   Dim RemoteInventory As DaqDeviceDescriptor
   Dim os As Long, NumIDs As Long
   Dim MatchFound As Boolean
   
   'mnNumDescriptors = 0
   If CreateDev Then
      EnumIndex% = 0
      'to do - move to after discovery and
      'remove only devices in UL whose UID
      'is not in the inventory
      For x = 0 To gnNumBoards - 1
         'check created against detected devices
         BoardNum& = gnBoardEnum(EnumIndex%)
         BoardName$ = GetNameOfBoard(BoardNum&)
         Prefix$ = Left(BoardName$, 3)
         If Not ((Prefix$ = "PCI") Or (Prefix$ = "WLS")) Then
            StLen& = ERRSTRLEN   '64
            ReturnString$ = Space$(StLen&)
            ConfigLen& = StLen&
            'BIFACTORYID
            ULS& = cbErrHandling(DONTPRINT, gnErrHandling)
            ULStat = GetConfigString573(BOARDINFO, BoardNum&, DevNum%, _
               BIDEVUNIQUEID, ReturnString$, ConfigLen&)
            ULS& = cbErrHandling(gnErrReporting, gnErrHandling)
            If ConfigLen& = 0 Then
               ReturnString$ = BoardName$
            Else
               ReturnString$ = Left$(ReturnString$, ConfigLen&)
            End If
            ReDim Preserve UIDStrings(NumIDs)
            UIDStrings(NumIDs) = ReturnString$ & "|" & Format(BoardNum&, "0")
            NumIDs = NumIDs + 1
            If Prefix$ = "DEM" Then NumDemoBoards% = NumDemoBoards% + 1
         End If
         EnumIndex% = EnumIndex% + 1
      Next
   End If
   mlNumDevices = 100
   
   If IsMissing(HostString) Then
      If Interface > 0 Then
         ULStat = cbGetDaqDeviceInventory(Interface, Inventory(0), mlNumDevices)
      Else
         mlNumDevices = 0
      End If
      If Not SaveFunc(mfNoForm, GetDaqDeviceInventory, ULStat, Interface, _
         0, mlNumDevices, Timeout, A5, A6, A7, A8, A9, A10, A11, 0) Then
         os = 0
         If (Interface = ANY_IFC) And (NumDemoBoards% > 0) Then
            'add existing Demo boards to the inventory
            For ExistingDesc% = 0 To mnNumDescriptors - 1
               If mxDescriptors(ExistingDesc%).ProductID = 45 Then
                  'add existing Demo to beginning of local inventory
                  For NuidItem% = 7 To 0 Step -1
                     NUID$ = NUID$ & mxDescriptors(ExistingDesc%).NUID(NuidItem%)
                  Next
                  If Not (NUID$ = PrevNUID$) Then
                     For NewDesc% = mlNumDevices To 1 Step -1
                        Inventory(NewDesc%) = Inventory(NewDesc% - 1)
                     Next
                     Inventory(0) = mxDescriptors(ExistingDesc%)
                     mlNumDevices = mlNumDevices + 1
                  End If
                  PrevNUID$ = NUID$
                  NUID$ = ""
               End If
            Next
         End If
         
         mnNumDescriptors = 0
         If GetDiscoverOption() Then
            For IDNum& = 0 To NumIDs - 1
               SplitID = Split(UIDStrings(IDNum&), "|")
               InstalledBoardNum& = SplitID(1)
               MatchFound = False
               For x = 0 To mlNumDevices - 1
                  BoardInInventory& = cbGetBoardNumber(Inventory(x))
                  If InstalledBoardNum& = BoardInInventory& Then
                     MatchFound = True
                     Exit For
                  End If
               Next
               If Not MatchFound Then RemoveBoardFromLibrary InstalledBoardNum&
            Next
         End If
         If gnInitializing Then RemoveDiscoveryForm
         'For DemoDesc% = 0 To mnNumDemoDescriptors - 1
         '   ReDim Preserve mxDescriptors(mnNumDescriptors)
         '   mxDescriptors(mnNumDescriptors) = mxDescriptors(DemoDesc%)
         '   mnNumDescriptors = mnNumDescriptors + 1
         'Next
         'If mlNumDevices > 0 Then _
         '   ReDim mxDescriptors(mlNumDevices - 1)

         If AddDemo Then
            Dim DemoDescriptor As DaqDeviceDescriptor
            Dim NameData() As Byte
            Dim UIDData() As Byte
            'to create a demo board, assign the PID
            'and create a unique ID for NUID
            DemoDescriptor.ProductID = 45
            NewName = "DEMO-BOARD"
            UIDString$ = Format(Rnd(3) * 100000000, "00000000")
            NameData = StrConv(NewName, vbFromUnicode)
            UIDData = StrConv(UIDString$, vbFromUnicode)
            For Element% = 0 To UBound(NameData)
               DemoDescriptor.ProductName(Element%) = NameData(Element%)
            Next
            For Element% = 0 To 7
               DemoDescriptor.NUID(Element%) = UIDData(Element%)
            Next
            If CreateDev Then
               CurBoardNum& = GetFirstAvailableBoardNum(os)
               AddBoardToLibrary CurBoardNum&, DemoDescriptor
               NewName = "DEMO-BOARD"
               NamesAdded = NamesAdded & NewName & ", "
               NumsAdded = NumsAdded & Format(CurBoardNum&, "0") & ", "
            End If
            ReDim Preserve mxDescriptors(mnNumDescriptors)
            mxDescriptors(mnNumDescriptors) = DemoDescriptor
            mnNumDescriptors = mnNumDescriptors + 1
         End If
         For x = 0 To mlNumDevices - 1
            BoardInInventory& = cbGetBoardNumber(Inventory(x))
            If BoardInInventory& = -1 Then
               If CreateDev Then
                  CurBoardNum& = GetFirstAvailableBoardNum(os)
                  AddBoardToLibrary CurBoardNum&, Inventory(x)
                  RawName$ = Inventory(x).ProductName
                  ConvName$ = StrConv(RawName$, vbUnicode)
                  NewName = Left(ConvName$, InStr(1, ConvName$, Chr(0)) - 1)
                  NamesAdded = NamesAdded & NewName & ", "
                  NumsAdded = NumsAdded & Format(CurBoardNum&, "0") & ", "
               End If
            End If
            ReDim Preserve mxDescriptors(mnNumDescriptors)
            mxDescriptors(mnNumDescriptors) = Inventory(x)
            os = CurBoardNum& + 1
            mnNumDescriptors = mnNumDescriptors + 1
         Next
      End If
   Else
      ULStat = cbGetNetDeviceDescriptor(HostString, HostPort, RemoteInventory, Timeout)
      If Not SaveFunc(mfNoForm, GetNetDeviceDescriptor, ULStat, HostString, _
         HostPort, 0, Timeout, A5, A6, A7, A8, A9, A10, A11, 0) Then
         os = 0
         BoardInInventory& = cbGetBoardNumber(RemoteInventory)
         If BoardInInventory& = -1 Then
            DevicesFound& = DevicesFound& + 1
            CurBoardNum& = GetFirstAvailableBoardNum(os)
            If CreateDev Then _
               AddBoardToLibrary CurBoardNum&, RemoteInventory
            RawName$ = RemoteInventory.ProductName
            ConvName$ = StrConv(RawName$, vbUnicode)
            NewName = Left(ConvName$, InStr(1, ConvName$, Chr(0)) - 1)
            NamesAdded = NamesAdded & NewName & ", "
            NumsAdded = NumsAdded & Format(os, "0") & ", "
            'DevicesFound& = DevicesFound& + 1
            ReDim Preserve mxDescriptors(mnNumDescriptors)
            mxDescriptors(mnNumDescriptors) = RemoteInventory
            mnNumDescriptors = mnNumDescriptors + 1
         End If
      End If
   End If
   
   If Not (NamesAdded = "") Then
      NamesAdded = Left(NamesAdded, Len(NamesAdded) - 2)
      NumsAdded = Left(NumsAdded, Len(NumsAdded) - 2)
      DevStat$ = " available to add to UL"
      If CreateDev Then DevStat$ = " added as board(s) "
      If Not CreateDev Then NumsAdded = ""
      PrintMain NamesAdded$ & DevStat$ & NumsAdded
   Else
      PrintMain "No devices added to library."
   End If
   DiscoverDevices = mnNumDescriptors

End Function

Public Function GetFirstAvailableBoardNum(ByVal CurrentNum As Long) As Long

   Dim Searching As Boolean
   Dim os As Long
   
   os = CurrentNum
   Searching = True
   Do
      If gnNumBoards = 0 Then Searching = False
      For i = 0 To gnNumBoards - 1
         BoardStored& = gnBoardEnum(i)
         If BoardStored& = os Then
            os = os + 1
            Searching = True
            Exit For
         End If
         Searching = False
      Next
   Loop While Searching
   GetFirstAvailableBoardNum = os

End Function

Public Sub AddBoardToLibrary(ByVal BoardNum As Long, ByRef Descriptor As DaqDeviceDescriptor)

   ULStat = cbCreateDaqDevice(BoardNum, Descriptor)
   If Not SaveFunc(mfNoForm, CreateDaqDevice, ULStat, BoardNum, _
      Descriptor.ProductID, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
      ReDim Preserve gnBoardEnum(gnNumBoards)
      gnBoardEnum(gnNumBoards) = BoardNum
      gnNumBoards = gnNumBoards + 1
      'mnSavedDescriptors = mnSavedDescriptors + 1
   End If
   'BoardInInventory& = cbGetBoardNumber(Descriptor)

End Sub

Public Sub RemoveBoardFromLibrary(ByVal BoardNumber As Long)

   BoardName$ = GetNameOfBoard(BoardNumber)
   ULStat = cbReleaseDaqDevice(BoardNumber)
   If Not SaveFunc(mfNoForm, ReleaseDaqDevice, ULStat, _
      BoardNumber, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, 0) Then
      Dim TempList As Variant
      TempList = gnBoardEnum()
      ReDim gnBoardEnum(0)
      NumBoards% = 0
      For i% = 0 To gnNumBoards - 1
         If Not TempList(i%) = BoardNumber Then
            ReDim Preserve gnBoardEnum(NumBoards%)
            gnBoardEnum(NumBoards%) = TempList(i%)
            NumBoards% = NumBoards% + 1
         End If
      Next
      gnNumBoards = NumBoards%
      'mnSavedDescriptors = mnSavedDescriptors - 1
      PrintMain BoardName$ & " (board " & Format(BoardNumber, "0") & ") removed from library"
   End If

End Sub

Public Function GetInventoryItem(ByVal ItemIndex As Integer, _
   ByRef UniqueIdentifier As String) As DaqDeviceDescriptor

   If Not UniqueIdentifier = "" Then
      For i& = 0 To mnNumDescriptors - 1
         UniqueID$ = mxDescriptors(i&).UniqueID
         UIDString$ = NullTermByteToString(UniqueID$)
         If UIDString$ = UniqueIdentifier Then
            GetInventoryItem = mxDescriptors(i&)
            Exit For
         End If
      Next
   Else
      GetInventoryItem = mxDescriptors(ItemIndex)
   End If
   
End Function

Public Function GetInventorySize() As Long

   GetInventorySize = mnNumDescriptors
   
End Function
