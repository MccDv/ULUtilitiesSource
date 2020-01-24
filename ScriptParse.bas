Attribute VB_Name = "ScriptParse"
Global gsCommand As String
Global Const UNITPERCENT = 1
Global Const UNITVOLTS = 2
Global Const UNITDEGREES = 3
Global Const UNITCOUNTS = 4
Global Const UNITFLOAT = 5
Global Const UNITLSBS = 6

'Scripting constants
Global Const SSetBoardName = 2000
Global Const SShowDiag = 2001
Global Const SContPlot = 2002
Global Const SConvData = 2003
Global Const SConvPT = 2004
Global Const SSetData = 2005
Global Const SSetAmplitude = 2006
Global Const SSetOffset = 2007
Global Const SCalCheck = 2008
Global Const SCountSet = 2009
Global Const SAddPTBuf = 2010
Global Const SSetDevName = 2011
Global Const SSetPlotOpts = 2012
Global Const SBufInfo = 2013
Global Const SSetPlotChan = 2014
Global Const SNextBlock = 2015
Global Const SSetBlock = 2016
Global Const SCalData = 2017
Global Const SSetResolution = 2018
Global Const SShowText = 2019
Global Const SGetTC = 2020
Global Const SToEng = 2021
Global Const SPlotType = 2022
Global Const SCalcNoise = 2023
Global Const SLogOutput = 2024
Global Const SLoadStringList = 2025
Global Const SGetStringFromList = 2026
Global Const SSetPlotScaling = 2027
Global Const SSetFirstPlotPoint = 2028
Global Const SLoadCSVList = 2029

Global Const SEvalEnable = 2030
Global Const SEvalDelta = 2031
Global Const SEvalMaxMin = 2032
Global Const SGetCSVsFromList = 2033
Global Const SCopyFile = 2036
Global Const SRunApp = 2037
Global Const SEndApp = 2038
Global Const SResetConfig = 2039
Global Const SEvalChannel = 2040

Global Const SGetFormRef = 2041
Global Const SSelPortRange = 2042
Global Const SSetPortDirection = 2043
Global Const SReadPortRange = 2044
Global Const SWritePortRange = 2045
Global Const SSelBitRange = 2046
Global Const SReadBitRange = 2047
Global Const SWriteBitRange = 2048
Global Const SSetBitDirection = 2052
Global Const SWaitForIdle = 2053
Global Const SWaitForEvent = 2054
Global Const SWaitStatusChange = 2055
Global Const SStopOnCount = 2056
Global Const SPlotOnCount = 2057
Global Const SSetBitsPerPort = 2058
Global Const SCounterArm = 2060

Global Const SGenerateData = 2049
Global Const SPlotGenData = 2050
Global Const SPlotAcqData = 2051

Global Const SDelay = 3000
Global Const SErrorPrint = 3001
Global Const SGetStatus = 3002
Global Const SErrorFlow = 3003
Global Const SULErrFlow = 3004
Global Const SULErrReport = 3005
Global Const SSetStaticOption = 3006
Global Const SScriptRate = 3008
Global Const SSetVariable = 3009
Global Const SGetFormProps = 3010
Global Const SGetStaticOptions = 3011
Global Const SSetFormProps = 3012
Global Const SCloseApp = 3013
Global Const SSetVarDefault = 3014
Global Const SPauseScript = 3015
Global Const SGetParameterString = 3016
Global Const SPicklist = 3017
Global Const SGenRndVal = 3018
Global Const SPickGroup = 3019
Global Const SSetLibType = 3020
Global Const SCalcMaxSinDelta = 3021
Global Const SIsListed = 3022
Global Const SPeriodCalc = 3023
Global Const SPulseWidthCalc = 3024
Global Const SMapAISwitch = 3025
Global Const SStopScript = 3026
Global Const SSetMCCControl = 3027
Global Const SGetDP8200Cmd = 3028
Global Const SEvalParamRev = 3029

Global Const SLoadSubScript = 4001
Global Const SCloseSubScript = 4002
Global Const SOpenWindow = 5001
Global Const SCloseWindow = 5002

'evaluation constants
Global Const EStatus = 10001
Global Const ETimeStamp = 10002
Global Const EEventType = 10003
Global Const EDataDC = 10010
Global Const EDataPulse = 10011
Global Const EDataDelta = 10012
Global Const EDataAmplitude = 10013
Global Const SEvalTrigPoint = 10015
Global Const EDataTime = 10021
Global Const EDataOutVsIn = 10022
Global Const EDataSkew = 10023
Global Const EData32Delta = 10030
Global Const EError = 10050
Global Const EHistogram = 10051
Global Const EStore488Value = 10052
Global Const ECompareStoredValue = 10053

Public Function ParseScriptLine(ScriptLine As String, Parameters As Variant) As Long

   Dim Arguments() As String
   
   ParamsFound& = FindInString(ScriptLine, ",", Locations)
   If Not ParamsFound& < 0 Then
      ReDim Arguments(ParamsFound& + 1)
      StartLoc& = 1
      For StringLoc& = 0 To ParamsFound& + 1
         If Not StringLoc& > ParamsFound& Then
            EndLoc& = Locations(StringLoc&) - StartLoc&
         Else
            EndLoc& = Len(ScriptLine)
         End If
         CurArg$ = Mid$(ScriptLine, StartLoc&, EndLoc&)
         Arguments(StringLoc&) = Trim$(CurArg$)
         If Not StringLoc& > ParamsFound& Then StartLoc& = Locations(StringLoc&) + 1
      Next StringLoc&
      Parameters = Arguments()
   Else
      ' "'" & ScriptLine & "' is not a valid script line. " & _
      "It should be changed to a comment or no-op.", vbCritical, "Invalid Script Line"
   End If
   ParseScriptLine = ParamsFound&

End Function

Public Function ParseUnits(Expression As String, TypeOfUnit As Integer) As String

   OrgExp$ = Expression
   For UnitType% = 1 To 4
      UnitString$ = Choose(UnitType%, "%", "V", "�", "C")
      UnitLoc& = InStr(Expression, UnitString$)
      If UnitLoc& > 0 Then
         OrgExp$ = Left(Expression, UnitLoc& - 1)
         TypeOfUnit = UnitType%
         'Exit For
      End If
   Next
   If TypeOfUnit = 0 Then
      If Not (InStr(1, Expression, ".") = 0) Then
         TypeOfUnit = UNITFLOAT
      Else
         TypeOfUnit = UNITLSBS
      End If
   End If
   ParseUnits = Trim(OrgExp$)
   
End Function

Public Function GetWindowType(WindowCode As Integer) As String

   Select Case WindowCode
      Case MAIN_FORM
         GetWindowType = "Main Form"
      Case ANALOG_IN
         GetWindowType = "Analog Input Form"
      Case ANALOG_OUT
         GetWindowType = "Analog Output Form"
      Case DIGITAL_IN
         GetWindowType = "Digital Input Form"
      Case DIGITAL_OUT
         GetWindowType = "Digital Output Form"
      Case COUNTERS
         GetWindowType = "Counter Form"
      Case UTILITIES
         GetWindowType = "Miscellaneous Form"
      Case Config
         GetWindowType = "Configuration Form"
      Case GPIB_CTL
         GetWindowType = "GPIB Control Form"
      Case LOGFUNC
         GetWindowType = "Log Functions Form"
      Case DATA_EVAL
         GetWindowType = "Data Evaluation Form"
      Case Else
         GetWindowType = "Scripting Command"
   End Select
   
End Function

Function GetScriptString(ScriptCode As Integer) As String

   Select Case ScriptCode
      Case ScriptTime
         ''ConstToVal = 0
      Case AIn
         GetScriptString = "AIn(, Range, , LowChan, HighChan, NumAInCalls)"
      Case AInScan
         GetScriptString = "AInScan(LowChan, HighChan, TotalCount, Rate, Range, , Options)"
      Case ALoadQueue
         GetScriptString = "ALoadQueue(QueueSize, QueueElement, ChannelType, Channel, Range, SetpointValue, [ChanMode], [DataRate])"
      Case APretrig
         GetScriptString = "APretrig(LowChan, HighChan, PreTrigCount, TotalCount, Rate, Range, , Options)"
      Case ATrig
         GetScriptString = "ATrig(LowChan, TrigType, , Range, , LowChan, HighChan, " & _
         "NumAInCalls, LowThreshold, HighThreshold)"
      Case FileAInScan
         GetScriptString = "FileAInScan(LowChan, HighChan, TotalCount, Rate, Range, FileName, Options)"
      Case FileGetInfo
         'ConstToVal = 7
      Case FilePretrig
         GetScriptString = "FilePretrig(LowChan, HighChan, PreTrigCount, TotalCount, Rate, Range, FileName, Options)"
      Case TIn
         GetScriptString = "TIn(, Scale, , Options, LowChan, HighChan, NumTInCalls)"
      Case TInScan
         GetScriptString = "TInScan(LowChan, HighChan, Scale, , Options, NumTInCalls)"
      Case AOut
         GetScriptString = "AOut(, Range, DACValue, LowChan, HighChan, NumAOutCalls)"
      Case AOutScan
         GetScriptString = "AOutScan(LowChan, HighChan, TotalCount, Rate, Range, , Options)"
      Case C8536Init
         'ConstToVal = 13
      Case C9513Init
         GetScriptString = "C9513Init(ChipNum, FOutDivider, FOutSource, Compare1, Compare2, TimeOfDay)"
      Case C8254Config
         GetScriptString = "C8254Config(CounterNum, 8254Configuration)"
      Case C8536Config
         'ConstToVal = 16
      Case C9513Config
         GetScriptString = "C9513Config(CounterNum, GateControl, CounterEdge, CountSource, SpecialGate, " & _
         "Reload, RecycleMode, BCDMode, CountDirec, OutputCtrl)"
      Case CLoad
         GetScriptString = "CLoad(CounterNum, LoadValue)"
      Case CIn
         GetScriptString = "CIn(CounterNum,,NumberOfReads)"
      Case CStoreOnInt
         GetScriptString = "CStoreOnInt(IntCount, , , MultiCounters)"
      Case CFreqIn
         GetScriptString = "CFreqIn(ChipNum, GateInterval)"
      Case DConfigPort
         GetScriptString = "DConfigPort(Port, Direction)"
      Case DBitIn
         GetScriptString = "DBitIn()"
      Case DIn
         GetScriptString = "DIn(Port)"
      Case DInScan
         GetScriptString = "DInScan(Port, TotalCount, Rate, , Options)"
      Case DBitOut
         GetScriptString = "DBitOut(Port, , Value, XPosition, YPosition)"
      Case DOut
         GetScriptString = "DOut(Port, Value)"
      Case DOutScan
         GetScriptString = "DOutScan(Port, TotalCount, Rate, , Options)"
      Case AConvertData
         'ConstToVal = 29
      Case ACalibrateData
         'ConstToVal = 30
      Case AConvertPretrigData
         'ConstToVal = 31
         GetScriptString = "AConvertPretrigData(True/False)"
      Case ToEngUnits
         'ConstToVal = 32
      Case FromEngUnits
         'ConstToVal = 33
      Case FileRead
         'ConstToVal = 34
      Case MemSetDTMode
         'ConstToVal = 35
      Case MemReset
         'ConstToVal = 36
      Case MemRead
         'ConstToVal = 37
      Case MemWrite
         'ConstToVal = 38
      Case MemReadPretrig
         'ConstToVal = 39
      Case RS485
         'ConstToVal = 40
      Case WinBufToArray
         'ConstToVal = 41
      Case WinArrayToBuf
         'ConstToVal = 42
      Case WinBufAlloc
         'ConstToVal = 43
      Case WinBufFree
         'ConstToVal = 44
      Case InByte
         GetScriptString = "InByte(Register, Composite, Consecutive, MaskValue, MaskFirst, " & _
         "MaskSecond, SurrogateEnable, SurrogateBoard, DevNum)"
      Case OutByte
         GetScriptString = "OutByte(Register, Value, Composite, Consecutive, MaskValue, MaskFirst, " & _
         "MaskSecond, SurrogateEnable, SurrogateBoard, DevNum)"
      Case InWord
         GetScriptString = "InWord(Register, Composite, Consecutive, MaskValue, MaskFirst, " & _
         "MaskSecond, SurrogateEnable, SurrogateBoard, DevNum)"
      Case OutWord
         GetScriptString = "OutWord(Register, Value, Composite, Consecutive, MaskValue, MaskFirst, " & _
         "MaskSecond, SurrogateEnable, SurrogateBoard, DevNum)"
      Case SetTrigger
         GetScriptString = "SetTrigger(TrigType, LowThreshold, HighThreshold, [TriggerOption])"
      Case DeclareRevision
         'ConstToVal = 50
      Case GetRevision
         'ConstToVal = 51
      Case LoadConfig
         'ConstToVal = 52
      Case GetBoardName
         'ConstToVal = 53
      Case GetConfig
         'ConstToVal = 54
      Case SetConfig
         GetScriptString = "SetConfig(InfoType, DevNum, ,ConfigItem, ConfigVal)"
      Case GetErrMsg
         'ConstToVal = 56
      Case ErrHandling
         'ConstToVal = 57
      Case GetStatus
         GetScriptString = "GetStatus()"
      Case StopBackground
         GetScriptString = "StopBackground()"
'new library
      Case AddBoard
         'ConstToVal = 60
      Case AddExp
         'ConstToVal = 61
      Case AddMem
         'ConstToVal = 62
      Case AIGetPcmCalCoeffs
         'ConstToVal = 63
      Case CreateBoard
         'ConstToVal = 64
      Case DeleteBoard
         'ConstToVal = 65
      Case SaveConfig
         GetScriptString = "SaveConfig(FileName)"
      Case C7266Config
         GetScriptString = "C7266Config(CounterNum, Quadrature, " & _
         "CountingMode, DataEncoding, IndexMode, InvertIndex, " & _
         "FlagPins, GateEnable)"
      Case CIn32
         GetScriptString = "CIn32(CounterNum, , NumberOfReads)"
      Case CLoad32
         GetScriptString = "CLoad32(CounterNum, LoadValue)"
      Case CStatus
         'ConstToVal = 70
      Case EnableEvent
         GetScriptString = "EnableEvent(EventType, EventSize)"
      Case DisableEvent
         GetScriptString = "DisableEvent(EventType)"
      Case CallbackFunc
         'ConstToVal = 73
      Case GetSubSystemStatus
         'ConstToVal = 74
      Case StopSubSystemBackground
         'ConstToVal = 75
      Case DConfigBit
         GetScriptString = "DConfigBit(Port, Direction, , XPosition, YPosition)"
      Case SelectSignal
         GetScriptString = "SelectSignal(SignalDirection, Signal, Connection, Polarity)"
      Case GetSignal
         'ConstToVal = 78
      Case FlashLED
         'ConstToVal = 79
      Case LogGetFileName
         'ConstToVal = 80
      Case LogGetFileInfo
         'ConstToVal = 81
      Case LogGetSampleInfo
         'ConstToVal = 82
      Case LogGetAIInfo
         'ConstToVal = 83
      Case LogGetCJCInfo
         'ConstToVal = 84
      Case LogGetDIOInfo
         'ConstToVal = 85
      Case LogReadTimeTags
         'ConstToVal = 86
      Case LogReadAIChannels
         'ConstToVal = 87
      Case LogReadCJCChannels
         'ConstToVal = 88
      Case LogReadDIOChannels
         'ConstToVal = 89
      Case LogConvertFile
         'ConstToVal = 90
      Case LogSetPreferences
         'ConstToVal = 91
      Case LogGetPreferences
         'ConstToVal = 92
      Case LogGetAIChannelCount
         'ConstToVal = 93
      Case CInScan
         GetScriptString = "CInScan(LowCounter, HighCounter, Samples, Rate, Options)"
      Case CConfigScan
         GetScriptString = "CConfigScan(, Mode, Debounce, DebounceTrig, Edge, MapChannel, , CounterNum)"
      Case CClear
         GetScriptString = "CClear(, CounterNum)"
      Case TimerOutStart
         GetScriptString = "TimerOutStart(TimerNum, Frequency)"
      Case TimerOutStop
         GetScriptString = "TimerOutStop(TimerNum)"
      Case WinBufAlloc32
         'ConstToVal = 99
      Case WinBufToArray32
         'ConstToVal = 100
      Case DaqInScan
         GetScriptString = "DaqInScan(, , , , Rate, , TotalCount, , Options, PretrigCount)"
      Case DaqSetTrigger
         GetScriptString = "DaqSetTrigger(Source, TrigType, TrigChan, ChanType, Range, " & _
         "TrigLevel, LevelVariance, TrigEvent)"
      Case DaqOutScan
         GetScriptString = "DaqOutScan(, , , , Rate, TotalCount, , Options)"
      Case GetTCValues
         'ConstToVal = 104
      Case VIn
         'ConstToVal = 105
      Case GetConfigString
         'ConstToVal = 106
      Case SetConfigString
         'ConstToVal = 107
      Case VOut
         GetScriptString = "VOut(, Range, Value, Options, LowChan, HighChan, NumVOutCalls)"
      Case DaqSetSetpoints
         GetScriptString = "DaqSetSetpoints(QueueSize, Element, LimitA, LimitB, Output1, Output2, " & _
         "Mask1, Mask2, Latch, OutputType)"
      Case DeviceLogin
         'ConstToVal = 110
      Case DeviceLogout
         'ConstToVal = 111
      Case AIn32
         GetScriptString = "AIn32(, Range, , Options, LowChan, HighChan, NumAInCalls)"
      Case VIn32  '116
         GetScriptString = "VIn32(, Range, , Options, LowChan, HighChan, NumVInCalls)"
      Case WinBufAlloc64   '117
         GetScriptString = "WinBufAlloc64(BufferSize)"
      Case ScaledWinBufToArray '118
         GetScriptString = "ScaledWinBufToArray(MemHandle&, DataBuffer#, FirstPoint&, CBCount&)"
      Case ScaledWinBufAlloc '119
         GetScriptString = "ScaledWinBufAlloc(NumPoints&)"
      Case TEDSRead  '120
         GetScriptString = "TEDSRead(BoardNum, Chan, DataBuffer(), CBCount, Options)"
      Case ScaledWinArrayToBuf  '121
         GetScriptString = "ScaledWinArrayToBuf(DataArray#, MemHandle&, FirstPoint&, CBCount&)"
      Case CIn64     '122
         GetScriptString = "CIn64(CounterNum,,NumberOfReads)"
      Case CLoad64   '123
         GetScriptString = "CLoad64(CounterNum, LoadValue)"
      Case WinBufToArray64 '124
         GetScriptString = "WinBufToArray64(DataArray#, MemHandle&, FirstPoint&, CBCount&)"
      Case IgnoreInstaCal   '125
         GetScriptString = "cbIgnoreInstaCal()"
      Case GetDaqDeviceInventory   '126
         GetScriptString = "cbGetDaqDeviceInventory(InterfaceType, DeviceDescriptor, NumberOfDevices&)"
      Case CreateDaqDevice   '127
         GetScriptString = "cbCreateDaqDevice(BoardNum&, DevDesc)"
      Case ReleaseDaqDevice   '128
         GetScriptString = "cbReleaseDaqDevice(BoardNum&)"
      Case GetBoardNumber   '129
         GetScriptString = "cbGetBoardNumber(DeviceDescriptor)"
      Case GetNetDeviceDescriptor   '130
         GetScriptString = "cbGetNetDeviceDescriptor(host$, Port&, DeviceDescriptor, Timeout&)"
      Case AInputMode   '131
         GetScriptString = "AInputMode(SetSingleEnded)"
      Case AChanInputMode   '132
         GetScriptString = "AChanInputMode(BoardNum, Chan, InputMode)"
      Case WinArrayToBuf32   '133
         GetScriptString = "cbWinArrayToBuf32(DataBuffer, MemHandle, FirstPoint, CBCount)"
      Case DInArray   '134
         GetScriptString = "cbDInArray(BoardNum&, FirstPort&, LastPort&, DataArray&)"
      Case DOutArray   '135
         GetScriptString = "cbDOutArray(BoardNum&, FirstPort&, LastPort&, DataArray&)"
      'Case DaqDeviceVersion   '136
      '   GetScriptString = "cbDaqDeviceVersion(BoardNum, VersionType, Version!)"
      Case DIn32     '137
         GetScriptString = "cbDIn32(BoardNum, PortNum, DataValue)"
      Case DOut32     '138
         GetScriptString = "cbDOut32(BoardNum, PortNum, DataValue)"
      Case DClearAlarm  '139
         GetScriptString = "DClearAlarm(BoardNum, PortNum, AlarmMask)"
      
'auxillary functions
      Case SelectCounters
         'ConstToVal = 200
   
'GPIB functions
      Case GPFind
         'ConstToVal = 201
      Case GPSend
         GetScriptString = "GPSend(Device, Command, [Return], [Conditional])"
      Case GPReceive
         GetScriptString = "GPReceive(Device)"
      Case GPTrigger
         GetScriptString = "GPTrigger(Device)"
      Case GPDevClear
         GetScriptString = "GPDevClear(Device, [Conditional])"
      Case GPIBAsk
         'ConstToVal = 206
      Case GPInit
         'ConstToVal = 207
      Case GPPtrs
         'ConstToVal = 208
      Case GPSelDevClear
         GetScriptString = "GPSelDevClear()"
      Case GPIBSre
         'ConstToVal = 210
      Case GPIBReturn
         GetScriptString = "GPReturnVal(ReturnString)"
   
   
  'Scripting constants
      Case SSetBoardName
         GetScriptString = "SSetBoardName(BoardName)"
      Case SShowDiag '2001
         GetScriptString = "SShowDiag(DialogText, DialogType, VarName$, Title$, Default, [Conditional])"
      Case SContPlot
         GetScriptString = "SContPlot(Enable)"
      Case SConvData
         GetScriptString = "SConvData(Enable)"
      Case SConvPT
         GetScriptString = "SConvPT(Enable)"
      Case SSetData
         GetScriptString = "SSetData(DataType)"
      Case SSetAmplitude
         GetScriptString = "SSetAmplitude(, , Amplitude)"
      Case SSetOffset   '2007
         GetScriptString = "SSetOffset(, , Offset)"
      Case SCalCheck
         'ConstToVal = 2008
      Case SCountSet
         GetScriptString = "SCountSet(NumPoints)"
      Case SAddPTBuf
         GetScriptString = "SAddPTBuf(Enable)"
      Case SSetDevName
         'ConstToVal = 2011
      Case SSetPlotOpts
         GetScriptString = "SSetPlotOpts(RetainPlot, ShowSource)"
      Case SBufInfo
         GetScriptString = "SBufInfo()"
      Case SSetPlotChan
         GetScriptString = "SSetPlotChan(ChannelToView)"
      Case SNextBlock
         GetScriptString = "SNextBlock()"
      Case SSetBlock '2016
         GetScriptString = "SSetBlock(NumPoints, [Conditional])"
      Case SCalData
         GetScriptString = "SCalData(, Enable)"
      Case SSetResolution
         GetScriptString = "SSetResolution(BitsResolution)"
      Case SShowText
         GetScriptString = "SShowText(, Enable)"
      Case SGetTC
         GetScriptString = "SGetTC(, Enable)"
      Case SToEng '2021
         GetScriptString = "SUseEngUnits(TF%)"
      Case SPlotType '2022
         GetScriptString = "SSetPlotType(TypeOfPlot)"
      Case SCalcNoise '2023
         GetScriptString = "SCalcNoise(TF%)"
      Case SLogOutput   '2024
         GetScriptString = "SLogOutput(ScreenOutput, FileOutput, FileName$)"
      Case SLoadStringList   '2025
         GetScriptString = "SLoadStringList(FileName$, ListSizeVariable$)"
      Case SGetStringFromList '2026
         GetScriptString = "SGetStringFromList(ListIndex, VariableToStore)"
      Case SSetPlotScaling    '2027
         GetScriptString = "SSetPlotScaling(PlotScaleMode)"
      Case SSetFirstPlotPoint '2028
         GetScriptString = "SSetFirstPlotPoint(FirstPlotPoint)"
      Case SLoadCSVList       '2029
         GetScriptString = "SLoadCSVList(FileName$, ListSizeVariable$, NumValuesVariable$, [ListName])"
      Case SEvalEnable '2030
         GetScriptString = "SEvalEnable(Enable)"
      Case SEvalDelta '2031
         GetScriptString = "SEvalData(Delta&)"
      Case SEvalMaxMin '2032
         GetScriptString = "SEvalDelta(Min&, Max&)"
      Case SGetCSVsFromList '2033
         GetScriptString = "SGetCSVsFromList(ListIndex, VarIndex, VariableToStore, [ListName])"
      Case SCopyFile    '2036
         GetScriptString = "SCopyFile(FileSource$, FileDestination$)"
      Case SRunApp    '2037
         GetScriptString = "SRunApp(CommandLine$, Wait%)"
      Case SEndApp    '2038
         GetScriptString = "SEndApp()"
      Case SResetConfig    '2039
         GetScriptString = "SResetConfig()"
      Case SEvalChannel '2040
         GetScriptString = "SEvalChannel(Channel%)"
      Case SGetFormRef  '2041
         GetScriptString = "SGetFormRef(RefNum)"
      Case SSelPortRange   '2042
         GetScriptString = "SSelPortRange(FirstPortIndex, LastPortIndex)"
      Case SSetPortDirection  '2043
         GetScriptString = "SSetPortDirection(Direction)"
      Case SReadPortRange  '2044
         GetScriptString = "SReadPortRange(NumberOfBlocks)"
      Case SWritePortRange '2045
         GetScriptString = "SWritePortRange(NumberOfBlocks)"
      Case SSelBitRange    '2046
         GetScriptString = "SSelBitRange(FirstBit, LastBit)"
      Case SReadBitRange   '2047
         GetScriptString = "SReadBitRange(NumberOfBlocks)"
      Case SWriteBitRange  '2048
         GetScriptString = "SWriteBitRange(NumberOfBlocks)"
      Case SGenerateData   '2049
         GetScriptString = "SGenerateData(DataType, Cycles, NumPoints, " & _
         "NumChans, Amplitude [p-p], Offset, SigType, NewData, Channel, FirstPoint)"
      Case SPlotGenData    '2050
         GetScriptString = "SPlotGenData()"
      Case SPlotAcqData    '2051
         GetScriptString = "SPlotAcqData()"
      Case SSetBitDirection '2052
         GetScriptString = "SSetBitDirection(Direction)"
      Case SWaitForIdle    '2053
         GetScriptString = "SWaitForIdle(StopOnCount, Timeout)"
      Case SWaitForEvent '2054
         GetScriptString = "SWaitForEvent(EventType, WaitData, Timeout)"
      Case SWaitStatusChange '2055
         GetScriptString = "SWaitStatusChange(StopDelta, WaitCondition, Timeout)"
      Case SStopOnCount '2056
         GetScriptString = "SStopOnCount(StopCount, Timeout)"
      Case SPlotOnCount '2057
         GetScriptString = "SPlotOnCount(PlotCount, TimeLimit)"
      Case SSetBitsPerPort ' 2058
         GetScriptString = "SSetBitsPerPort(PortNum, CumBits, BitsInPort)"
      Case SCounterArm ' 2060
         GetScriptString = "SCounterArm(CtrNum, EnableDisable, Conditional)"
      
      Case SDelay    '3000
         GetScriptString = "SDelay(DelayTime)"
      Case SErrorPrint
         GetScriptString = "SErrorPrint(Enable)"
      Case SGetStatus
         GetScriptString = "SGetStatus(Enable)"
      Case SErrorFlow   '3003
         GetScriptString = "SErrorFlow(LocalErrHandling)"
      Case SULErrFlow   '3004
         GetScriptString = "SULErrFlow(ULErrHandling)"
      Case SULErrReport   '3005
         GetScriptString = "SULErrReport(ULErrReporting)"
      Case SSetStaticOption   '3006
         GetScriptString = "SSetStaticOption(StaticOptions, [Conditional])"
      Case SScriptRate    '3008
         GetScriptString = "SScriptRate(MilliSeconds)"
      Case SSetVariable    '3009
         GetScriptString = "SSetVariable(VarName, VarValue, [Conditional])"
      Case SGetFormProps   '3010
         GetScriptString = "SGetFormProps(PropName, StoreVariable)"
      Case SGetStaticOptions '3011
         GetScriptString = "SGetStaticOptions()"
      Case SSetFormProps   '3012
         GetScriptString = "SSetFormProps(PropName, PropVal)"
      Case SCloseApp    '3013
         GetScriptString = "SCloseApp()"
      Case SSetVarDefault  '3014
         GetScriptString = "SSetVarDefault(VarName, DefaultValue)"
      Case SPauseScript    '3015
         GetScriptString = "SPauseScript([Conditional])"
      Case SGetParameterString '3016
         GetScriptString = "SGetParameterString(Function, ParamNumber, ParamValue, ReturnedString)"
      Case SPicklist       '3017
         GetScriptString = "SPicklist(ListIndex, ListToPick, VariableToSet, [ListSize], [NumericDefault])"
      Case SGenRndVal      '3018
         GetScriptString = "SGenRndVal(SeedValue, VariableToSet)"
      Case SPickGroup      '3019
         GetScriptString = "SPickGroup(GroupIndicator$, GroupIndex&, GroupList$, ListToSet$, NumGroups&)"
      Case SSetLibType     '3020
         GetScriptString = "SSetLibType(LibType)"
      Case SCalcMaxSinDelta     '3021
         GetScriptString = "SCalcMaxSinDelta(Amplitude, SourceRate, IsPerChan, Result, [XClockRate])"
      Case SIsListed       '3022
         GetScriptString = "SIsListed(ListToSearch, ValueToSearch, IsListed, [AtIndex], [ListSeparator])"
      Case SPeriodCalc       '3023
         GetScriptString = "SPeriodCalc(Rate, Period)"
      Case SPulseWidthCalc   '3024
         GetScriptString = "SPulseWidthCalc(Time, PulseWidth)"
      Case SMapAISwitch   '3025
         GetScriptString = "SMapAISwitch(MapString)"
      Case SStopScript   '3026
         GetScriptString = "SStopScript([Conditional])"
      Case SSetMCCControl  '3027
         GetScriptString = "SSetMCCControl(TrueFalse)"
      Case SGetDP8200Cmd   '3028
         GetScriptString = "SGetDP8200Cmd(Value, Command)"
      Case SEvalParamRev   '3029
         GetScriptString = "SEvalParamRev(Value, EvalCondition)"
      Case SLoadSubScript
         GetScriptString = "SLoadSubScript(Duplicate Board, FileName, BoardName, [Condition])"
      Case SCloseSubScript
         GetScriptString = "SCloseSubScript()"
      Case SOpenWindow
         GetScriptString = "SOpenWindow()"
      Case SCloseWindow
         GetScriptString = "SCloseWindow()"
   
      Case EStatus ' 10001
         GetScriptString = "EStatus(Condition, StatusLow, StatusHigh, FailTimeout)"
      Case ETimeStamp '10002
         GetScriptString = "ETimeStamp()"
      Case EEventType    '10003
         GetScriptString = "EEventType(EventType, EventData, " & _
         "FailIfData, NoEvent, FailIfTimeout)"
      Case EDataDC ' 10010
         GetScriptString = "EDataDC(DataPoints, EvalChan, " & _
         "DCValue, Tolerance [켏SBs], DCOption)"
      Case EDataPulse   '10011
         GetScriptString = "EDataPulse(DataPoints, EvalChan, " & _
         "Pulse HiVal [V], Pulse LoVal [V], Tolerance [켞olts], HiBy Sample, " & _
         "LoBy Sample, Time Tolerance, Evaluation Option)"
      Case EDataDelta ' 10012
         GetScriptString = "EDataDelta(DataPoints, EvalChan, " & _
         "ValueOfChange [V], FailIfDelta, Evaluation Option)"
      Case EDataAmplitude  '10013
         GetScriptString = "EDataAmplitude(DataPoints, EvalChan, " & _
         "Amplitude [p-p], ATolerance [켏SBs], EvaluationOption)"
      Case SEvalTrigPoint  '10015
         GetScriptString = "SEvalTrigPoint(DataPoints, EvalChan, " & _
         "TrigPolarity, Threshold [Volts], Guardband, Tolerance [S], [EvalOption])"
      Case EDataTime '10021
         GetScriptString = "EDataTime(DataPoints, EvalChan, " & _
         "Threshold [Volts], Guardband [켏SBs], SourceFreq, " & _
         "Ext Clock Rate [S/s], Tolerance [S/~], EvaluationOption)"
      Case EDataOutVsIn '10022
         GetScriptString = "EDataOutVsIn(DataPoints, EvalChan, " & _
         "FailIfComparison, OutputFormRef, NumberOfBits)"
      Case EDataSkew    '10023
         GetScriptString = "EDataSkew(DataPoints, NumChans, " & _
         "Threshold [Volts], Guardband [켏SBs], " & _
         "Tolerance [S], EvaluationOption)"
      Case EData32Delta ' 10030
         GetScriptString = "EData32Delta(DataPoints, EvalChan, " & _
         "ValueOfChange, FailIfDelta, DeltaOption)"
      Case EError    '10050
         GetScriptString = "EError(, Function, ExpectedError, " & _
         "Alternate1, Alternate2, Alternate3, Action)"
      Case EHistogram   '10051
         GetScriptString = "EHistogram(BinSpread, MaxRMSValue, " & _
         "AverageValue, AvgValTol)"
      Case EStore488Value '10052
         GetScriptString = "EStore488Value()"
      Case ECompareStoredValue '10053
         GetScriptString = "ECompareStoredValue(ExpectedValue, ErrorUnits, " & _
         "Tolerance)"
   End Select
   
End Function

Function ParseScriptString(ScriptString As String, Elements As Variant) As Long
   
   Dim Arguments() As String
   
   EndFuncName& = InStr(1, ScriptString, "(")
   EndFunction& = InStr(1, ScriptString, ")")
   FuncLength& = EndFunction& - EndFuncName&
   If EndFuncName& > 1 Then
      ReDim Arguments(0)
      Arguments(0) = Left(ScriptString, EndFuncName& - 1)
      NumArgs& = NumArgs& + 1
      If FuncLength& > 1 Then
         ParamsFound& = FindInString(ScriptString, ",", Locations)
         'If Not ParamsFound& < 0 Then ReDim Preserve Arguments(ParamsFound& + 1)
         StartLoc& = EndFuncName& + 1
         For StringLoc& = 0 To ParamsFound& + 1
            If Not StringLoc& > ParamsFound& Then
               EndLoc& = Locations(StringLoc&) - StartLoc&
            Else
               EndLoc& = Len(ScriptString)
            End If
            CurArg$ = Mid$(ScriptString, StartLoc&, EndLoc&)
            If Right(CurArg$, 1) = ")" Then CurArg$ = Left(CurArg$, Len(CurArg$) - 1)
            ReDim Preserve Arguments(NumArgs&)
            Arguments(NumArgs&) = Trim$(CurArg$)
            If Not StringLoc& > ParamsFound& Then StartLoc& = Locations(StringLoc&) + 1
            NumArgs& = NumArgs& + 1
         Next StringLoc&
      End If
      Elements = Arguments()
   End If
   ParseScriptString = NumArgs& - 1

End Function

Function ConstToVal(Constant As String) As Long

   Select Case Constant
      Case "ScriptTime"
         ConstValue& = 0
      Case "AIn"
         ConstValue& = 1
      Case "AInScan"
         ConstValue& = 2
      Case "ALoadQueue"
         ConstValue& = 3
      Case "APretrig"
         ConstValue& = 4
      Case "ATrig"
         ConstValue& = 5
      Case "FileAInScan"
         ConstValue& = 6
      Case "FileGetInfo"
         ConstValue& = 7
      Case "FilePretrig"
         ConstValue& = 8
      Case "TIn"
         ConstValue& = 9
      Case "TInScan"
         ConstValue& = 10
      Case "AOut"
         ConstValue& = 11
      Case "AOutScan"
         ConstValue& = 12
      Case "C8536Init"
         ConstValue& = 13
      Case "C9513Init"
         ConstValue& = 14
      Case "C8254Config"
         ConstValue& = 15
      Case "C8536Config"
         ConstValue& = 16
      Case "C9513Config"
         ConstValue& = 17
      Case "CLoad"
         ConstValue& = 18
      Case "CIn"
         ConstValue& = 19
      Case "CStoreOnInt"
         ConstValue& = 20
      Case "CFreqIn"
         ConstValue& = 21
      Case "DConfigPort"
         ConstValue& = 22
      Case "DBitIn"
         ConstValue& = 23
      Case "DIn"
         ConstValue& = 24
      Case "DInScan"
         ConstValue& = 25
      Case "DBitOut"
         ConstValue& = 26
      Case "DOut"
         ConstValue& = 27
      Case "DOutScan"
         ConstValue& = 28
      Case "AConvertData"
         ConstValue& = 29
      Case "ACalibrateData"
         ConstValue& = 30
      Case "AConvertPretrigData"
         ConstValue& = 31
      Case "ToEngUnits"
         ConstValue& = 32
      Case "FromEngUnits"
         ConstValue& = 33
      Case "FileRead"
         ConstValue& = 34
      Case "MemSetDTMode"
         ConstValue& = 35
      Case "MemReset"
         ConstValue& = 36
      Case "MemRead"
         ConstValue& = 37
      Case "MemWrite"
         ConstValue& = 38
      Case "MemReadPretrig"
         ConstValue& = 39
      Case "RS485"
         ConstValue& = 40
      Case "WinBufToArray"
         ConstValue& = 41
      Case "WinArrayToBuf"
         ConstValue& = 42
      Case "WinBufAlloc"
         ConstValue& = 43
      Case "WinBufFree"
         ConstValue& = 44
      Case "InByte"
         ConstValue& = 45
      Case "OutByte"
         ConstValue& = 46
      Case "InWord"
         ConstValue& = 47
      Case "OutWord"
         ConstValue& = 48
      Case "SetTrigger"
         ConstValue& = 49
   
      Case "DeclareRevision"
         ConstValue& = 50
      Case "GetRevision"
         ConstValue& = 51
      Case "LoadConfig"
         ConstValue& = 52
      Case "GetBoardName"
         ConstValue& = 53
      Case "GetConfig"
         ConstValue& = 54
      Case "SetConfig"
         ConstValue& = 55
      Case "GetErrMsg"
         ConstValue& = 56
      Case "ErrHandling"
         ConstValue& = 57
      Case "GetStatus"
         ConstValue& = 58
      Case "StopBackground"
         ConstValue& = 59
'new library
      Case "AddBoard"
         ConstValue& = 60
      Case "AddExp"
         ConstValue& = 61
      Case "AddMem"
         ConstValue& = 62
      Case "AIGetPcmCalCoeffs"
         ConstValue& = 63
      Case "CreateBoard"
         ConstValue& = 64
      Case "DeleteBoard"
         ConstValue& = 65
      Case "SaveConfig"
         ConstValue& = 66
      Case "C7266Config"
         ConstValue& = 67
      Case "CIn32"
         ConstValue& = 68
      Case "CLoad32"
         ConstValue& = 69
      Case "CStatus"
         ConstValue& = 70
      Case "EnableEvent"
         ConstValue& = 71
      Case "DisableEvent"
         ConstValue& = 72
      Case "CallbackFunc"
         ConstValue& = 73
      Case "GetSubSystemStatus"
         ConstValue& = 74
      Case "StopSubSystemBackground"
         ConstValue& = 75
      Case "DConfigBit"
         ConstValue& = 76
      Case "SelectSignal"
         ConstValue& = 77
      Case "GetSignal"
         ConstValue& = 78
      Case "FlashLED"
         ConstValue& = 79
      Case "LogGetFileName"
         ConstValue& = 80
      Case "LogGetFileInfo"
         ConstValue& = 81
      Case "LogGetSampleInfo"
         ConstValue& = 82
      Case "LogGetAIInfo"
         ConstValue& = 83
      Case "LogGetCJCInfo"
         ConstValue& = 84
      Case "LogGetDIOInfo"
         ConstValue& = 85
      Case "LogReadTimeTags"
         ConstValue& = 86
      Case "LogReadAIChannels"
         ConstValue& = 87
      Case "LogReadCJCChannels"
         ConstValue& = 88
      Case "LogReadDIOChannels"
         ConstValue& = 89
      Case "LogConvertFile"
         ConstValue& = 90
      Case "LogSetPreferences"
         ConstValue& = 91
      Case "LogGetPreferences"
         ConstValue& = 92
      Case "LogGetAIChannelCount"
         ConstValue& = 93
      Case "CInScan"
         ConstValue& = 94
      Case "CConfigScan"
         ConstValue& = 95
      Case "CClear"
         ConstValue& = 96
      Case "TimerOutStart"
         ConstValue& = 97
      Case "TimerOutStop"
         ConstValue& = 98
      Case "WinBufAlloc32"
         ConstValue& = 99
      Case "WinBufToArray32"
         ConstValue& = 100
      Case "DaqInScan"
         ConstValue& = 101
      Case "DaqSetTrigger"
         ConstValue& = 102
      Case "DaqOutScan"
         ConstValue& = 103
      Case "GetTCValues"
         ConstValue& = 104
      Case "VIn"
         ConstValue& = 105
      Case "GetConfigString"
         ConstValue& = 106
      Case "SetConfigString"
         ConstValue& = 107
      Case "VOut"
         ConstValue& = 108
      Case "DaqSetSetpoints"
         ConstValue& = 109
      Case "DeviceLogin"
         ConstValue& = 110
      Case "DeviceLogout"
         ConstValue& = 111
      Case "AIn32"
         ConstValue& = 114
      Case "ToEngUnits32"
         ConstValue& = 115
      Case "VIn32"  '116
         ConstValue& = 116
      Case "WinBufAlloc64"
         ConstValue& = 117
      Case "ScaledWinBufToArray"
         ConstValue& = 118
      Case "ScaledWinBufAlloc"
         ConstValue& = 119
      Case "TEDSRead"  '120
         ConstValue& = 120
      Case "ScaledWinArrayToBuf"  '121
         ConstValue& = 121
      Case "CIn64"     '122
         ConstValue& = 122
      Case "CLoad64"
         ConstValue& = 123
      Case "WinBufToArray64"
         ConstValue& = 124
      Case "IgnoreInstaCal"
         ConstValue& = 125
      Case "GetDaqDeviceInventory"
         ConstValue& = 126
      Case "CreateDaqDevice"
         ConstValue& = 127
      Case "ReleaseDaqDevice"
         ConstValue& = 128
      Case "GetBoardNumber"
         ConstValue& = 129
      Case "GetNetDeviceDescriptor"
         ConstValue& = 130
      Case "AInputMode"
         ConstValue& = 131
      Case "AChanInputMode"
         ConstValue& = 132
      Case "WinArrayToBuf32"
         ConstValue& = 133
      Case "DInArray"
         ConstValue& = 134
      Case "DOutArray"
         ConstValue& = 135
      Case "DaqDeviceVersion"
         ConstValue& = 136
      Case "DIn32"
         ConstValue& = 137
      Case "DOut32"
         ConstValue& = 138
      Case "DClearAlarm"
         ConstValue& = 139

'auxillary functions
      Case "SelectCounters"
         ConstValue& = 200
   
'GPIB functions
      Case "GPFind"
         ConstValue& = 201
      Case "GPSend"
         ConstValue& = 202
      Case "GPReceive"
         ConstValue& = 203
      Case "GPTrigger"
         ConstValue& = 204
      Case "GPDevClear"
         ConstValue& = 205
      Case "GPIBAsk"
         ConstValue& = 206
      Case "GPInit"
         ConstValue& = 207
      Case "GPPtrs"
         ConstValue& = 208
      Case "GPSelDevClear"
         ConstValue& = 209
      Case "GPIBSre"
         ConstValue& = 210
      Case "GPReturnVal"
         ConstValue& = 211
   
   
  'Scripting constants
      Case "SSetBoardName"
         ConstValue& = 2000
      Case "SShowDiag"
         ConstValue& = 2001
      Case "SContPlot"
         ConstValue& = 2002
      Case "SConvData"
         ConstValue& = 2003
      Case "SConvPT"
         ConstValue& = 2004
      Case "SSetData"
         ConstValue& = 2005
      Case "SSetAmplitude"
         ConstValue& = 2006
      Case "SSetOffset"
         ConstValue& = 2007
      Case "SCalCheck"
         ConstValue& = 2008
      Case "SCountSet"
         ConstValue& = 2009
      Case "SAddPTBuf"
         ConstValue& = 2010
      Case "SSetDevName"
         ConstValue& = 2011
      Case "SSetPlotOpts"
         ConstValue& = 2012
      Case "SBufInfo"
         ConstValue& = 2013
      Case "SSetPlotChan"
         ConstValue& = 2014
      Case "SNextBlock"
         ConstValue& = 2015
      Case "SSetBlock"
         ConstValue& = 2016
      Case "SCalData"
         ConstValue& = 2017
      Case "SSetResolution"
         ConstValue& = 2018
      Case "SShowText"
         ConstValue& = 2019
      Case "SGetTC"
         ConstValue& = 2020
      Case "SToEng"
         ConstValue& = 2021
      Case "SSetPlotType"
         ConstValue& = 2022
      Case "SCalcNoise"
         ConstValue& = 2023
      Case "SLogOutput"
         ConstValue& = 2024
      Case "SLoadStringList"
         ConstValue& = 2025
      Case "SGetStringFromList"
         ConstValue& = 2026
      Case "SSetPlotScaling"
         ConstValue& = 2027
      Case "SSetFirstPlotPoint"
         ConstValue& = 2028
      Case "SLoadCSVList"
         ConstValue& = 2029
      Case "SEvalEnable"
         ConstValue& = 2030
      Case "SEvalDelta"
         ConstValue& = 2031
      Case "SEvalMaxMin"
         ConstValue& = 2032
      Case "SGetCSVsFromList"
         ConstValue& = 2033
      Case "SCopyFile"
         ConstValue& = 2036
      Case "SRunApp"
         ConstValue& = 2037
      Case "SEndApp"
         ConstValue& = 2038
      Case "SResetConfig"
         ConstValue& = 2039
      Case "SEvalChannel"
         ConstValue& = 2040
      Case "SGetFormRef"
         ConstValue& = 2041
      Case "SSelPortRange"
         ConstValue& = 2042
      Case "SSetPortDirection"
         ConstValue& = 2043
      Case "SReadPortRange"
         ConstValue& = 2044
      Case "SWritePortRange"
         ConstValue& = 2045
      Case "SSelBitRange"
         ConstValue& = 2046
      Case "SReadBitRange"
         ConstValue& = 2047
      Case "SWriteBitRange"
         ConstValue& = 2048
      Case "SGenerateData"
         ConstValue& = 2049
      Case "SPlotGenData"
         ConstValue& = 2050
      Case "SPlotAcqData"
         ConstValue& = 2051
      Case "SSetBitDirection"
         ConstValue& = 2052
      Case "SWaitForIdle"
         ConstValue& = 2053
      Case "SWaitForEvent"
         ConstValue& = 2054
      Case "SWaitStatusChange"
         ConstValue& = 2055
      Case "SStopOnCount"
         ConstValue& = 2056
      Case "SPlotOnCount"
         ConstValue& = 2057
      Case "SSetBitsPerPort"
         ConstValue& = 2058
      Case "SCounterArm"
         ConstValue& = 2060
      
      Case "SDelay"
         ConstValue& = 3000
      Case "SErrorPrint"
         ConstValue& = 3001
      Case "SGetStatus"
         ConstValue& = 3002
      Case "SErrorFlow"
         ConstValue& = 3003
      Case "SULErrFlow"
         ConstValue& = 3004
      Case "SULErrReport"
         ConstValue& = 3005
      Case "SSetStaticOption"
         ConstValue& = 3006
      Case "STimeStamp"
         ConstValue& = 3007
      Case "SScriptRate"
         ConstValue& = 3008
      Case "SSetVariable"
         ConstValue& = 3009
      Case "SGetFormProps"
         ConstValue& = 3010
      Case "SGetStaticOptions"
         ConstValue& = 3011
      Case "SSetFormProps"
         ConstValue& = 3012
      Case "SCloseApp"
         ConstValue& = 3013
      Case "SSetVarDefault"
         ConstValue& = 3014
      Case "SPauseScript"
         ConstValue& = 3015
      Case "SGetParameterString"
         ConstValue& = 3016
      Case "SPicklist"
         ConstValue& = 3017
      Case "SGenRndVal"
         ConstValue& = 3018
      Case "SPickGroup"
         ConstValue& = 3019
      Case "SSetLibType"
         ConstValue& = 3020
      Case "SCalcMaxSinDelta"
         ConstValue& = 3021
      Case "SIsListed"
         ConstValue& = 3022
      Case "SPeriodCalc"
         ConstValue& = 3023
      Case "SPulseWidthCalc"
         ConstValue& = 3024
      Case "SMapAISwitch"
         ConstValue& = 3025
      Case "SStopScript"
         ConstValue& = 3026
      Case "SSetMCCControl"
         ConstValue& = 3027
      Case "SGetDP8200Cmd"
         ConstValue& = 3028
      Case "SEvalParamRev"
         ConstValue& = 3029
      
      Case "SLoadSubScript"
         ConstValue& = 4001
      Case "SCloseSubScript"
         ConstValue& = 4002
      Case "SOpenWindow"
         ConstValue& = 5001
      Case "SCloseWindow"
         ConstValue& = 5002
'evaluation constants
      Case "EStatus"
         ConstValue& = 10001
      Case "ETimeStamp"
         ConstValue& = 10002
      Case "EEventType"
         ConstValue& = 10003
      Case "EDataDC"
         ConstValue& = 10010
      Case "EDataPulse"
         ConstValue& = 10011
      Case "EDataDelta"
         ConstValue& = 10012
      Case "EDataAmplitude"
         ConstValue& = 10013
      Case "SEvalTrigPoint"
         ConstValue& = 10015
      Case "EDataTime"
         ConstValue& = 10021
      Case "EDataOutVsIn"
         ConstValue& = 10022
      Case "EDataSkew"
         ConstValue& = 10023
      Case "EData32Delta"
         ConstValue& = 10030
      Case "EError"
         ConstValue& = 10050
      Case "EHistogram"
         ConstValue& = 10051
      Case "EStore488Value"
         ConstValue& = 10052
      Case "ECompareStoredValue"
         ConstValue& = 10053
      Case Else
         ConstValue& = -1
   End Select
   ConstToVal = ConstValue&

End Function

Function GetParamInfo(ParamName As String, Param As String, Optional Qualifier As Variant) As String

   If Not IsMissing(Qualifier) Then CheckQualifier% = True
   
   Select Case ParamName
      Case "[ListName]"
         Reply$ = "Optional name of list file (required for multiple lists in one session)."
      Case "[TriggerOption]"
         Reply$ = "Optional bitmapped value for various trigger options."
      Case "8254Configuration"
         ArgVal& = Val(Param)
         Reply$ = Get8254ConfigString(ArgVal&)
      Case "Action"
         Reply$ = "Bitfield: 8 = Include 'no error'"
      Case "Amplitude [p-p]"
         Reply$ = "Volts if decimal point is used, temperature " & _
         "if � is used, percentage of FS if % is used, otherwise, counts."
      Case "ATolerance [켏SBs]"
         Reply$ = "Counts if value is followed by 'C', otherwise, LSBs."
      Case "ChannelType"
         ArgVal& = Val(Param)
         Reply$ = GetChannelTypeString(ArgVal&)
      Case "Command"
         Reply$ = GetGPIBCmdText(Param)
      Case "Condition"
         ArgVal& = Val(Param)
         Reply$ = GetStatCondString(ArgVal&)
      Case "ConfigItem"
         ArgVal& = Val(Param)
         Reply$ = GetCfgItemString(ArgVal&)
      Case "Connection"
         ArgVal& = Val(Param)
         Reply$ = GetConnectionString(ArgVal&)
      Case "CounterNum"
         Reply$ = "" '"Must be one based for ALL products (even zero based products)."
      Case "CountingMode"
         ArgVal& = Val(Param)
         Reply$ = GetCountingModeString(ArgVal&)
      Case "DACValue"
         Reply$ = "Counts or, if '%' is used, percentage of full scale."
      Case "DataEncoding"
         ArgVal& = Val(Param)
         Select Case ArgVal&
            Case BCD_ENCODING
               Reply$ = "BCDENCODING"
            Case BINARY_ENCODING
               Reply$ = "BINARYENCODING"
            Case Else
               Reply$ = "Invalid parameter value"
         End Select
      Case "DataType"
         ArgVal& = Val(Param)
         Reply$ = GetDataTypeString(ArgVal&)
      Case "DCValue"
         Reply$ = "Volts if decimal point is used, otherwise, counts."
      Case "DeltaOption"
         If CheckQualifier% Then
            If Qualifier < 3 Then
               Reply$ = "MovingAverage (default if only a number is supplied), " & _
               "FirstPoint (use the format 'OptionString = Value')"
            Else
               Reply$ = "Number of failing transitions to ignore."
            End If
         Else
            Reply$ = "For FailIfDelta < 3: MovingAverage (default if only a number is supplied), " & _
            "FirstPoint (use the format 'OptionString = Value') - " & _
            "Otherwise, number of transitions to ignore."
         End If
      Case "Device"
         Select Case Param
            Case "HP8112"
               Reply$ = "Uses a 9513 counter device at specified board number " & _
               "if using MCC boards for control."
         End Select
      Case "DialogType"
         Reply$ = "0 for simple dialog, 1 to set run-time variable."
      Case "Direction"
         ArgVal& = Val(Param)
         Reply$ = GetDirectionString(ArgVal&)
      Case "ErrorUnits"
         ArgVal& = Val(Param)
         Reply$ = GetErrorUnits(ArgVal&)
      Case "EvaluationOption"
         Reply$ = "MovingAverage (default if only a number is supplied), " & _
         "FirstPoint (use the format 'OptionString = Value')"
      Case "Edge"
         ArgVal& = Val(Param)
         Reply$ = GetCtrEdgeString(ArgVal&)
      Case "EventType"
         ArgVal& = Val(Param)
         Reply$ = GetEventTypeString(ArgVal&)
      Case "ExpectedError", "Alternate1", "Alternate2", "Alternate3"
         ArgVal& = Val(Param)
         Reply$ = GetErrorText(ArgVal&)
      Case "Ext Clock Rate [S/s]"
         Reply$ = "Required for EXTCLOCK or cbAIn - otherwise, may be zero"
      Case "FailIfComparison"
         Reply$ = "Number of bits to compare (0 = all), negative values for inverted values"
      Case "FailIfData"
         Reply$ = "IsEqual(0), IsLessThan(1), NotEqual(2), GreaterThan(3)"
      Case "FailIfDelta"
         ArgVal& = Val(Param)
         Reply$ = GetFailIfDeltaString(ArgVal&)
      Case "FailIfTimeout"
         Reply$ = "If non-zero, fails if a timeout occurs."
      Case "FileOutput"
         Reply$ = "Bitfield: 1=LogAll, 2=LogErrors, 4=LogComments"
      Case "FirstBit"
         ArgVal& = Val(Param)
         If ArgVal& < 0 Then
            Reply$ = "Unselect all bits"
         End If
      Case "FirstPortIndex"
         Port% = Val(Param)
         PortType& = GetPortFromIndex(Port%)
         If PortType& = 0 Then
            Reply$ = "Unselect all ports"
         Else
            IntPort% = PortType&
            Reply$ = GetPortStringEx(IntPort%)
         End If
      Case "FlagPins"
         ArgVal& = Val(Param)
         Reply$ = GetFlagPinString(ArgVal&)
      Case "Function"
         Argument% = Val(Param)
         Reply$ = GetFunctionName(Argument%)
      Case "HiBy Sample"
         Reply$ = "Set > TotalCount to ignore."
      Case "HighThreshold"
         Reply$ = "If float (contains decimal point), converts from voltage."
      Case "IndexMode"
         ArgVal& = Val(Param)
         Reply$ = GetIndexModeString(ArgVal&)
      Case "InfoType"
         ArgVal& = Val(Param)
         Reply$ = GetCfgInfoTypeString(ArgVal&)
      Case "LastPortIndex"
         Port% = Val(Param)
         PortType& = GetPortFromIndex(Port%)
         IntPort% = PortType&
         Reply$ = GetPortStringEx(IntPort%)
      Case "LibType"
         Reply$ = "0 = UniLib, 1 = .NetLib, 2 = MsgLib"
      Case "LoBy Sample"
         Reply$ = "Set > TotalCount to ignore."
      Case "LocalErrHandling"
         ArgVal& = Val(Param)
         Reply$ = GetLocalErrHandlingString(ArgVal&)
      Case "LowCounter", "HighCounter"
         Reply$ = "Use actual value that UL uses."
      Case "LowThreshold"
         Reply$ = "If float (contains decimal point), converts from voltage."
      Case "Mode"
         ArgVal& = Val(Param)
         Reply$ = GetCtrModeString(ArgVal&)
      Case "NoEvent"
         Reply$ = "If non-zero, fails if an event occurs."
      Case "NumberOfBits"
         Reply$ = "Number of bits to evaluate or 0 if evaluating port."
      Case "NumberOfReads"
         Reply$ = "Negative values return one sample for each script line."
      Case "Offset"
         Reply$ = "Volts if decimal point is used, " & _
         "percentage of FS if % is used, otherwise, counts."
      Case "Options", "StaticOptions"
         ArgVal& = Val(Param)
         Reply$ = GetOptionsString(ArgVal&, ANALOG_IN)
      Case "PlotScaleMode"
         Reply$ = "0=Full Scale, 1=Auto-scale, 2=Fixed"
      Case "Polarity"
         ArgVal& = Val(Param)
         Reply$ = GetPolarityString(ArgVal&)
      Case "Port"
         ArgVal& = Val(Param)
         Reply$ = GetPortStringEx(ArgVal&)
      Case "Quadrature"
         ArgVal& = Val(Param)
         Reply$ = GetQuadString(ArgVal&)
      Case "Range"
         ArgVal& = Val(Param)
         Reply$ = GetRangeString(ArgVal&)
      Case "Scale"
         ArgVal& = Val(Param)
         Reply$ = GetScaleString(ArgVal&)
      Case "ScreenOutput"
         Reply$ = "Bitfield: 1=PrintAll, 2=PrintErrors, 4=PrintComments, 8=Pause"
      Case "SetConfig"
         ArgVal& = Val(Param)
         Reply$ = GetConfigGlobalString(ArgVal&)
      Case "Signal"
         ArgVal& = Val(Param)
         Reply$ = GetSignalString(ArgVal&)
      Case "SignalDirection"
         ArgVal& = Val(Param)
         Reply$ = GetSigDirString(ArgVal&)
      Case "SigType"
         ArgVal& = Val(Param)
         Reply$ = GetSigTypeString(ArgVal&)
      Case "StatusLow"
         ArgVal& = Val(Param)
         Reply$ = GetStatLowString(ArgVal&)
      Case "StatusHigh"
         ArgVal& = Val(Param)
         Reply$ = GetStatHighString(ArgVal&)
      Case "StopDelta"
         ArgVal& = Val(Param)
         Reply$ = "No change (0)."
         If ArgVal& = 1 Then Reply$ = "Not static (1)."
      Case "StopOnCount"
         Reply$ = "Values greater than 0 end scan when CurCount reaches value."
      Case "Timeout"
         Reply$ = "If non-zero, aborts after Timeout iterations of script timer."
      Case "Tolerance [S/~]"
         Reply$ = "Use negative values for AIn or multi-A/D products."
      Case "TrigPolarity"
         ArgVal& = Val(Param)
         Reply$ = GetTrigPolarity(ArgVal&)
      Case "TrigType"
         ArgVal& = Val(Param)
         Reply$ = GetTrigTypeString(ArgVal&)
      Case "TypeOfPlot"
         ArgVal& = Val(Param)
         Reply$ = GetPlotTypeString(ArgVal&)
      Case "ULErrHandling"
         ArgVal& = Val(Param)
         Reply$ = GetLibraryErrHandlingString(ArgVal&)
      Case "ULErrReporting"
         ArgVal& = Val(Param)
         Reply$ = GetLibraryErrReportingString(ArgVal&)
      Case "WaitData"
         Reply$ = "If non-zero, waits for EventData to " & _
         "increment beyond this value."
   End Select
   GetParamInfo = Reply$
   
End Function

Public Function LocateScriptFile(FileSpecification As String, ByVal ScriptDir As String, _
ByVal MasterDir As String, Attempts As Integer) As String

   Static stDirLoc As Long
   
   'get the actual file name from the FileSpecification
   If FileSpecification = "" Then
      Attempts = 5
   Else
      SpecParts = Split(FileSpecification, "\")
      Filename$ = SpecParts(UBound(SpecParts))
      TryFile$ = Filename$
   End If
   
   Select Case Attempts
      Case 0
         stDirLoc = 0
         If Left(FileSpecification, 3) = "..\" Then
            'if relative path
            NumDirs& = FindInString(MasterDir, "\", Locations)
            RelDirs& = FindInString(FileSpecification, "..\", Positions)
            RelReDirs& = FindInString(FileSpecification, "\", RePositions)
            If RelReDirs& - RelDirs& > 0 Then
               'there are directories above the relative directory specified
               StartRelSpec& = RePositions(RelReDirs& - (RelReDirs& - RelDirs&)) + 1
               LenRelSpec& = RePositions(RelReDirs&) - RePositions(RelDirs&)
               RelDirSpec$ = Mid(FileSpecification, StartRelSpec&, LenRelSpec&)
            Else
               RelDirSpec$ = ""
            End If
            If Not (NumDirs& - (RelDirs& + 1)) < 0 Then
               TryPath$ = Left(MasterDir, Locations(NumDirs& - RelDirs&)) & RelDirSpec$
            End If
         Else
            TryPath$ = ""
            TryFile$ = FileSpecification
         End If
         Attempts = Attempts + 1
      Case 1
         Attempts = Attempts + 1
         TryPath$ = ScriptDir
         If Not (InStr(1, FileSpecification, ":") = 0) Then
            TryFile$ = Filename$   'FileSpecification
         Else
            TryFile$ = FileSpecification
         End If
      Case 2
         Attempts = Attempts + 1
         TryPath$ = MasterDir
         TryFile$ = Filename$
      Case 3
         'try parent directories of directory master script is running from
         NumDirs& = FindInString(MasterDir, "\", Locations)
         If stDirLoc = 0 Then
            stDirLoc = NumDirs&
            TryPath$ = MasterDir
            If Not Right(TryPath$, 1) = "\" Then TryPath$ = TryPath$ & "\"
         Else
            If NumDirs& > 0 Then
               TryPath$ = Left(MasterDir, Locations(stDirLoc + 1))
            End If
         End If
         TryFile$ = Filename$
         stDirLoc = stDirLoc - 1
         If stDirLoc <= 0 Then Attempts = Attempts + 1
      Case 4
         Attempts = Attempts + 1
         TryPath$ = App.Path & "\"
         TryFile$ = Filename$
      Case 5
         TryPath$ = ""
         TryFile$ = ""
   End Select
   'CurDir() & "\"
   LocateScriptFile = TryPath$ & TryFile$

End Function
