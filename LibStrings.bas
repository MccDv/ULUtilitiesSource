Attribute VB_Name = "LibStrings"
   Global Const ScriptTime = 0
   Global Const AIn = 1
   Global Const AInScan = 2
   Global Const ALoadQueue = 3
   Global Const APretrig = 4
   Global Const ATrig = 5
   Global Const FileAInScan = 6
   Global Const FileGetInfo = 7
   Global Const FilePretrig = 8
   Global Const TIn = 9
   Global Const TInScan = 10
   
   Global Const AOut = 11
   Global Const AOutScan = 12
   
   Global Const C8536Init = 13
   Global Const C9513Init = 14
   Global Const C8254Config = 15
   Global Const C8536Config = 16
   Global Const C9513Config = 17
   Global Const CLoad = 18
   Global Const CIn = 19
   Global Const CStoreOnInt = 20
   Global Const CFreqIn = 21
   
   Global Const DConfigPort = 22
   Global Const DBitIn = 23
   Global Const DIn = 24
   Global Const DInScan = 25
   
   Global Const DBitOut = 26
   Global Const DOut = 27
   Global Const DOutScan = 28
   
   Global Const AConvertData = 29
   Global Const ACalibrateData = 30
   Global Const AConvertPretrigData = 31
   Global Const ToEngUnits = 32
   Global Const FromEngUnits = 33
   Global Const FileRead = 34
   
   Global Const MemSetDTMode = 35
   Global Const MemReset = 36
   Global Const MemRead = 37
   Global Const MemWrite = 38
   Global Const MemReadPretrig = 39
   
   Global Const RS485 = 40
   Global Const WinBufToArray = 41
   Global Const WinArrayToBuf = 42
   Global Const WinBufAlloc = 43
   Global Const WinBufFree = 44
   Global Const InByte = 45
   Global Const OutByte = 46
   Global Const InWord = 47
   Global Const OutWord = 48
   Global Const SetTrigger = 49
   
   Global Const DeclareRevision = 50
   Global Const GetRevision = 51
   Global Const LoadConfig = 52

   Global Const GetBoardName = 53
   Global Const GetConfig = 54
   Global Const SetConfig = 55
   Global Const GetErrMsg = 56

   Global Const ErrHandling = 57
   Global Const GetStatus = 58
   Global Const StopBackground = 59

   'new library
   Global Const AddBoard = 60
   Global Const AddExp = 61
   Global Const AddMem = 62
   Global Const AIGetPcmCalCoeffs = 63
   Global Const CreateBoard = 64
   Global Const DeleteBoard = 65
   Global Const SaveConfig = 66

   Global Const C7266Config = 67
   Global Const CIn32 = 68
   Global Const CLoad32 = 69
   Global Const CStatus = 70
   Global Const EnableEvent = 71
   Global Const DisableEvent = 72
   Global Const CallbackFunc = 73
   Global Const GetSubSystemStatus = 74
   Global Const StopSubSystemBackground = 75
   Global Const DConfigBit = 76
   Global Const SelectSignal = 77
   Global Const GetSignal = 78
   Global Const FlashLED = 79
   Global Const LogGetFileName = 80
   Global Const LogGetFileInfo = 81
   Global Const LogGetSampleInfo = 82
   Global Const LogGetAIInfo = 83
   Global Const LogGetCJCInfo = 84
   Global Const LogGetDIOInfo = 85
   Global Const LogReadTimeTags = 86
   Global Const LogReadAIChannels = 87
   Global Const LogReadCJCChannels = 88
   Global Const LogReadDIOChannels = 89
   Global Const LogConvertFile = 90
   Global Const LogSetPreferences = 91
   Global Const LogGetPreferences = 92
   Global Const LogGetAIChannelCount = 93
   Global Const CInScan = 94
   Global Const CConfigScan = 95
   Global Const CClear = 96
   Global Const TimerOutStart = 97
   Global Const TimerOutStop = 98
   Global Const WinBufAlloc32 = 99
   Global Const WinBufToArray32 = 100
   Global Const DaqInScan = 101
   Global Const DaqSetTrigger = 102
   Global Const DaqOutScan = 103
   Global Const GetTCValues = 104
   Global Const VIn = 105
   Global Const GetConfigString = 106
   Global Const SetConfigString = 107
   Global Const VOut = 108
   Global Const DaqSetSetpoints = 109
   Global Const DeviceLogin = 110
   Global Const DeviceLogout = 111
   Global Const PulseOutStart = 112
   Global Const PulseOutStop = 113
   Global Const AIn32 = 114
   Global Const ToEngUnits32 = 115
   Global Const VIn32 = 116
   Global Const WinBufAlloc64 = 117
   Global Const ScaledWinBufToArray = 118
   Global Const ScaledWinBufAlloc = 119
   Global Const TEDSRead = 120
   Global Const ScaledWinArrayToBuf = 121
   Global Const CIn64 = 122
   Global Const CLoad64 = 123
   Global Const WinBufToArray64 = 124
   Global Const IgnoreInstaCal = 125
   Global Const GetDaqDeviceInventory = 126
   Global Const CreateDaqDevice = 127
   Global Const ReleaseDaqDevice = 128
   Global Const GetBoardNumber = 129
   Global Const GetNetDeviceDescriptor = 130
   Global Const AInputMode = 131
   Global Const AChanInputMode = 132
   Global Const WinArrayToBuf32 = 133
   Global Const DInArray = 134
   Global Const DOutArray = 135
   Global Const DaqDeviceVersion = 136
   Global Const DIn32 = 137
   Global Const DOut32 = 138
   Global Const DClearAlarm = 139
   
   'auxillary functions
   Global Const SelectCounters = 200

   'GPIB functions
   Global Const GPFind = 201
   Global Const GPSend = 202
   Global Const GPReceive = 203
   Global Const GPTrigger = 204
   Global Const GPDevClear = 205
   Global Const GPIBAsk = 206
   Global Const GPInit = 207
   Global Const GPPtrs = 208
   Global Const GPSelDevClear = 209
   Global Const GPIBSre = 210
   Global Const GPIBReturn = 211
   
   Dim mnNoErrorFile As Integer, mnNoBoardFile As Integer
   
Function GetFunctionName(FunctionNum As Integer) As String

   Func$ = GetFunctionString(FunctionNum)
   Loca = InStr(1, Func$, "(")
   If Not Loca = 0 Then Reply$ = Left(Func$, Loca) & ")"
   GetFunctionName = Reply$

End Function

Function GetFunctionString(FunctionNum As Integer)

   Select Case FunctionNum
      Case 0
         FunctionString$ = "ScriptTime (Seconds)"
      Case AIn
         FunctionString$ = "cbAIn(BoardNum, Chan, Gain, DataValue)  ScriptVals(FirstChan, LastChan, NumPoints)"
      Case AInScan
         FunctionString$ = "cbAInScan(BoardNum, LowChan, HighChan, CBCount, CBRate, Gain, MemHandle, Options)"
      Case ALoadQueue
         FunctionString$ = "cbALoadQueue(BoardNum, ChanArray, GainArray, NumChans, [ChanMode], [DataRate])"
      Case APretrig
         FunctionString$ = "cbAPretrig(BoardNum, LowChan, HighChan, PretrigCount, CBCount, CBRate, Gain, MemHandle, Options)"
      Case ATrig
         FunctionString$ = "cbATrig(BoardNum, Chan, TrigType, TrigValue, Gain, DataValue)"
      Case FileAInScan
         FunctionString$ = "cbFileAInScan(BoardNum, LowChan, HighChan, CBCount, CBRate, Gain, FileName, Options)"
      Case FileGetInfo
         FunctionString$ = "cbFileGetInfo(FileName, LowChan, HighChan, PretrigCount, TotalCount, CBRate, Gain)"
      Case FilePretrig
         FunctionString$ = "cbFilePretrig(BoardNum, LowChan, HighChan, PretrigCount, CBCount, CBRate, Gain, FileName, Options)"
      Case TIn
         FunctionString$ = "cbTIn(BoardNum, Chan, CBScale, TempValue, Options)"
      Case TInScan   '10
         FunctionString$ = "cbTInScan(BoardNum, LowChan, HighChan, CBScale, DataBuffer, Options)"
      Case AOut
         FunctionString$ = "cbAOut(BoardNum, Chan, Gain, DataValue)"
      Case AOutScan
         FunctionString$ = "cbAOutScan(BoardNum, LowChan, HighChan, CBCount, CBRate, Gain, MemHandle, Options)"
      Case C8536Init
         FunctionString$ = "cbC8536Init(BoardNum, ChipNum, Ctr1Output)"
      Case C9513Init
         FunctionString$ = "cbC9513Init(BoardNum, ChipNum, FOutDivider, FOutSource, Compare1, Compare2, TimeOfDay)"
      Case C8254Config
         FunctionString$ = "cbC8254Config(BoardNum, CounterNum, Config)"
      Case C8536Config
         FunctionString$ = "cbC8536Config(BoardNum, CounterNum, OutputControl, RecycleMode, Retrigger)"
      Case C9513Config
         FunctionString$ = "cbC9513Config(BoardNum, CounterNum, GateControl, CounterEdge, CountSource, SpecialGate, Reload, RecycleMode, BCDMode, CountDirec, OutputCtrl)"
      Case CLoad
         FunctionString$ = "cbCLoad(BoardNum, RegNum, LoadValue)"
      Case CIn
         FunctionString$ = "cbCIn(BoardNum, CounterNum, CBCount)"
      Case CStoreOnInt  '20
         FunctionString$ = "cbCStoreOnInt(BoardNum, IntCount, CntrControl, DataBuffer) ScriptVals(CtrBitFeild)"
      Case CFreqIn
         FunctionString$ = "cbCFreqIn(BoardNum, SigSource, GateInterval, CBCount, Freq)"
      Case DConfigPort
         FunctionString$ = "cbDConfigPort(BoardNum, PortNum, Direction)"
      Case DBitIn
         FunctionString$ = "cbDBitIn(BoardNum, PortType, BitNum, BitValue)"
      Case DIn
         FunctionString$ = "cbDIn(BoardNum, PortNum, DataValue)"
      Case DInScan
         FunctionString$ = "cbDInScan(BoardNum, PortNum, CBCount, CBRate, MemHandle, Options)"
      Case DBitOut
         FunctionString$ = "cbDBitOut(BoardNum, PortType, BitNum, BitValue)"
      Case DOut
         FunctionString$ = "cbDOut(BoardNum, PortNum, DataValue)"
      Case DOutScan
         FunctionString$ = "cbDOutScan(BoardNum, PortNum, CBCount, CBRate, MemHandle, Options)"
      Case AConvertData
         FunctionString$ = "cbAConvertData(BoardNum, NumPoints, ADData, ChanTags)"
      Case ACalibrateData  '30
         FunctionString$ = "cbACalibrateData(BoardNum, NumPoints, Gain, ADData)"
      Case AConvertPretrigData
         FunctionString$ = "cbAConvertPretrigData(BoardNum, PretrigCount, TotalCount, ADData, ChanTags)"
      Case ToEngUnits
         FunctionString$ = "cbToEngUnits(BoardNum, Range, DataVal, EngUnits)"
      Case FromEngUnits
         FunctionString$ = "cbFromEngUnits(BoardNum, Range, EngUnits, DataVal)"
      Case FileRead
         FunctionString$ = "cbFileRead(FileName, FirstPoint, NumPoints, DataBuffer)"
      Case MemSetDTMode
         FunctionString$ = "cbMemSetDTMode(BoardNum, Mode)"
      Case MemReset
         FunctionString$ = "cbMemReset(BoardNum)"
      Case MemRead
         FunctionString$ = "cbMemRead(BoardNum, DataBuffer, FirstPoint, CBCount)"
      Case MemWrite
         FunctionString$ = "cbMemWrite(BoardNum, DataBuffer, FirstPoint, CBCount)"
      Case MemReadPretrig
         FunctionString$ = "cbMemReadPretrig(BoardNum, DataBuffer, FirstPoint, CBCount)"
      Case RS485  '40
         FunctionString$ = "cbRS485(BoardNum, Transmit, Receive)"
      Case WinBufToArray
         FunctionString$ = "cbWinBufToArray(MemHandle, DataBuffer, FirstPoint, CBCount)"
      Case WinArrayToBuf
         FunctionString$ = "cbWinArrayToBuf(DataBuffer, MemHandle, FirstPoint, CBCount)"
      Case WinBufAlloc
         FunctionString$ = "cbWinBufAlloc(NumPoints)"
      Case WinBufFree
         FunctionString$ = "cbWinBufFree(MemHandle)"
      Case InByte
         FunctionString$ = "cbInByte(BoardNum, PortNum)"
      Case OutByte
         FunctionString$ = "cbOutByte(BoardNum, PortNum, PortVal)"
      Case InWord
         FunctionString$ = "cbInWord(BoardNum, PortNum)"
      Case OutWord
         FunctionString$ = "cbOutWord(BoardNum, PortNum, PortVal)"
      Case SetTrigger
         FunctionString$ = "cbSetTrigger(BoardNum, TrigType, LowThreshold, HighThreshold, [TriggerOption])"
      Case DeclareRevision '50
         FunctionString$ = "cbDeclareRevision(RevNum)"
      Case GetRevision
         FunctionString$ = "cbGetRevision(DLLRevNum, VXDRevNum)"
      Case LoadConfig
         FunctionString$ = "cbLoadConfig(FileName)"
      Case GetBoardName
         FunctionString$ = "cbGetBoardName(BoardNum, BoardName)"
      Case GetConfig
         FunctionString$ = "cbGetConfig(InfoType, BoardNum, DevNum, ConfigItem, ConfigVal)"
      Case SetConfig
         FunctionString$ = "cbSetConfig(InfoType, BoardNum, DevNum, ConfigItem, ConfigVal)"
      Case GetErrMsg
         FunctionString$ = "cbGetErrMsg(ErrCode, ErrMsg)"
      Case ErrHandling
         FunctionString$ = "cbErrHandling(ErrReporting, ErrHandling)"
      Case GetStatus
         If WIN32APP And (CURRENTREVNUM > 5.19) Then
            FunctionString$ = "cbGetStatus(BoardNum, Status, CurCount, CurIndex, FunctionType)"
         Else
            FunctionString$ = "cbGetStatus(BoardNum, Status, CurCount, CurIndex)"
         End If
      Case StopBackground
         If WIN32APP And (CURRENTREVNUM > 5.19) Then
            FunctionString$ = "cbStopBackground(BoardNum, FunctionType)"
         Else
            FunctionString$ = "cbStopBackground(BoardNum)"
         End If
      Case AddBoard  '60
         FunctionString$ = "cbAddBoard()"
      Case AddExp
         FunctionString$ = "cbAddExp()"
      Case AddMem
         FunctionString$ = "cbAddMem()"
      Case AIGetPcmCalCoeffs
         FunctionString$ = "cbAIGetPcmCalCoeffs()"
      Case CreateBoard
         FunctionString$ = "cbCreateBoard()"
      Case DeleteBoard
         FunctionString$ = "cbDeleteBoard()"
      Case SaveConfig
         FunctionString$ = "cbSaveConfig(FileName)"
      Case C7266Config
         FunctionString$ = "cbC7266Config(BoardNum, CounterNum, Quadrature, CountingMode, DataEncoding, IndexMode, InvertIndex, FlagPins, GateEnable)"
      Case CIn32
         FunctionString$ = "cbCIn32(BoardNum, CounterNum, CBCount)"
      Case CLoad32
         FunctionString$ = "cbCLoad32(BoardNum, RegNum, LoadValue)"
      Case CStatus   '70
         FunctionString$ = "cbCStatus(BoardNum, CounterNum, StatusBits)"
      Case EnableEvent
         FunctionString$ = "cbEnableEvent(BoardNum, EventType, EventSize, ProcAddress, UserData)"
      Case DisableEvent
         FunctionString$ = "cbDisableEvent(BoardNum, EventType)"
      Case CallbackFunc
         FunctionString$ = "MyCallbackFunc(BoardNum, EventType, EventData, UserData)"
      Case GetSubSystemStatus
         FunctionString$ = "cbGetSubSystemStatus (BoardNum&, Subsystem&, Status%, CurCount&, CurIndex&)"
      Case StopSubSystemBackground
         FunctionString$ = "cbStopSubSystemBackground (BoardNum&, Subsystem&)"
      Case DConfigBit
         FunctionString$ = "cbDConfigBit (BoardNum&, PortNum&, BitNum&, Direction&)"
      Case SelectSignal
         FunctionString$ = "cbSelectSignal(BoardNum&, Direction&, Signal&, Connection&, Polarity&)"
      Case GetSignal
         FunctionString$ = "cbGetSignal(BoardNum&, Direction&, Signal&, Index&, Connection&, Polarity&)"
      Case FlashLED
         FunctionString$ = "cbFlashLED(BoardNum&)"
      Case LogGetFileName  '80
         FunctionString$ = "cbLogGetFileName(FileNum&, Path$, Filename$)"
      Case LogGetFileInfo
         FunctionString$ = "cbLogGetFileInfo(Filename$, Version&, Size&)"
      Case LogGetSampleInfo
         FunctionString$ = "cbLogGetSampleInfo(Filename$, SampleInterval&, SampleCount&, StartDate&, StartTime&)"
      Case LogGetAIInfo
         FunctionString$ = "cbLogGetAIInfo(Filename$, ChannelMask&, UnitMask&, AIChannelCount&)"
      Case LogGetCJCInfo
         FunctionString$ = "cbLogGetCJCInfo(Filename$, CJCChannelCount&)"
      Case LogGetDIOInfo
         FunctionString$ = "cbLogGetDIOInfo(Filename$, DIOChannelCount&)"
      Case LogReadTimeTags
         FunctionString$ = "cbLogReadTimeTags(Filename$, StartSample&, SampleCount&, Dates&, Times&)"
      Case LogReadAIChannels
         FunctionString$ = "cbLogReadAIChannels(Filename$, StartSample&, SampleCount&, AIChannelData!)"
      Case LogReadCJCChannels
         FunctionString$ = "cbLogReadCJCChannels(Filename$, StartSample&, SampleCount&, CJCChannelData!)"
      Case LogReadDIOChannels
         FunctionString$ = "cbLogReadDIOChannels(Filename$, StartSample&, SampleCount&, DIOChannelData&)"
      Case LogConvertFile  '90
         FunctionString$ = "cbLogConvertFile(Filename$, DestPath$, FileType&, StartSample&, SampleCount&, Delimiter&)"
      Case LogSetPreferences
         FunctionString$ = "cbLogSetPreferences(TimeFormat&, TimeZone&, Units&)"
      Case LogGetPreferences
         FunctionString$ = "cbLogGetPreferences(TimeFormat&, TimeZone&, Units&)"
      Case LogGetAIChannelCount
         FunctionString$ = "cbLogGetAIChannelCount(Filename$, AIChannelCount&)"
      Case CInScan
         FunctionString$ = "cbCInScan(BoardNum&, LowChan&, HighChan&, CBCount&, CBRate&, MemHandle&, Options&)"
      Case CConfigScan
         FunctionString$ = "cbCConfigScan(BoardNum&, Chan&, Mode&, DebounceTime&, DebounceTrigger&, EdgeDetection&, TickSize&, MapChannel&)"
      Case CClear
         FunctionString$ = "cbCClear(BoardNum&, CounterNum&)"
      Case TimerOutStart
         FunctionString$ = "cbTimerOutStart(BoardNum&, TimerNum&, Frequency)"
      Case TimerOutStop
         FunctionString$ = "cbTimerOutStop(BoardNum&, TimerNum&)"
      Case WinBufAlloc32
         FunctionString$ = "cbWinBufAlloc32(NumPoints&)"
      Case WinBufToArray32 '100
         FunctionString$ = "cbWinBufToArray32(MemHandle&, DataBuffer&, FirstPoint&, CBCount&)"
      Case DaqInScan
         FunctionString$ = "cbDaqInScan(BoardNum&, ChanArray%, ChanTypeArray%, GainArray%, ChanCount&, CBRate&, PretrigCount&, CBCount&, MemHandle&, Options&)"
      Case DaqSetTrigger
         FunctionString$ = "cbDaqSetTrigger(BoardNum&, TrigSource&, TrigSense&, TrigChan&, ChanType&, Gain&, Level!, Variance!, TrigEvent&)"
      Case DaqOutScan
         FunctionString$ = "cbDaqOutScan(BoardNum&, ChanArray%, ChanTypeArray%, GainArray%, ChanCount&, CBRate&, CBCount&, MemHandle&, Options&)"
      Case GetTCValues  '104
         FunctionString$ = "cbGetTCValues(BoardNum&, ChanArray%, ChanTypeArray%, ChanCount&, MemHandle&, FirstPoint&, count&, CBScale&, TempValArray!)"
      Case VIn
         FunctionString$ = "cbVIn(BoardNum, Chan, CBRange, Voltage, Options)"
      Case GetConfigString '= 106
         FunctionString$ = "cbGetConfigString(InfoType%, mnBoardNum, DevNum%, ConfigItem%, ReturnString$, ConfigLen&)"
      Case SetConfigString '= 107
         FunctionString$ = "cbSetConfigString(InfoType%, mnBoardNum, DevNum%, ConfigItem%, ReturnString$, ConfigLen&)"
      Case VOut '= 108
         FunctionString$ = "cbVOut(BoardNum, Chan, CBRange, Voltage, Options)"
      Case DaqSetSetpoints '= 109
         FunctionString$ = "cbDaqSetSetpoints(BoardNum, LimitAArray!, LimitBArray!, Reserved!, SetpointFlagsArray&, SetpointOutputArray&, Output1Array!, Output2Array!, OutputMask1Array!, OutputMask2Array!, SetpointCount&)"
      Case DeviceLogin     '= 110
         FunctionString$ = "cbDeviceLogin(BoardNum, Account$, Password$)"
      Case DeviceLogout    '= 111
         FunctionString$ = "cbDeviceLogout(BoardNum)"
      Case PulseOutStart
         FunctionString$ = "cbPulseOutStart(BoardNum&, TimerNum&, Frequency, " & _
         "DutyCycle, PulseCount, InitialDelay, IdleState, Options)"
      Case PulseOutStop
         FunctionString$ = "cbPulseOutStop(BoardNum&, TimerNum&)"
      Case AIn32     '= 114
         FunctionString$ = "AIn32(BoardNum, Chan, Gain, DataValue, Options)  ScriptVals(FirstChan, LastChan, NumPoints)"
      Case ToEngUnits32 '115
         FunctionString$ = "cbToEngUnits32(BoardNum, Range, DataVal, EngUnits)"
      Case VIn32 '116
         FunctionString$ = "cbVIn32(BoardNum, Chan, Range, DataVal, Options)"
      Case WinBufAlloc64   '117
         FunctionString$ = "cbWinBufAlloc64(BufferSize)"
      Case ScaledWinBufToArray '118
         FunctionString$ = "cbScaledWinBufToArray(MemHandle&, DataBuffer#, FirstPoint&, CBCount&)"
      Case ScaledWinBufAlloc '119
         FunctionString$ = "cbScaledWinBufAlloc(NumPoints&)"
      Case TEDSRead  '120
         FunctionString$ = "cbTEDSRead(BoardNum, Chan, DataBuffer(), CBCount, Options)"
      Case ScaledWinArrayToBuf   '121
         FunctionString$ = "cbScaledWinArrayToBuf(DataArray#, MemHandle&, FirstPoint&, CBCount&)"
      Case CIn64     '122
         FunctionString$ = "cbCIn64(BoardNum, CounterNum, CBCount)"
      Case CLoad64   '123
         FunctionString$ = "cbCLoad64(BoardNum, RegNum, LoadValue)"
      Case WinBufToArray64   '124
         FunctionString$ = "cbWinBufToArray64(DataArray#, MemHandle&, FirstPoint&, CBCount&)"
      Case IgnoreInstaCal   '125
         FunctionString$ = "cbIgnoreInstaCal()"
      Case GetDaqDeviceInventory   '126
         FunctionString$ = "cbGetDaqDeviceInventory(InterfaceType, DeviceDescriptor, NumberOfDevices&)"
      Case CreateDaqDevice   '127
         FunctionString$ = "cbCreateDaqDevice(BoardNum&, DevDesc)"
      Case ReleaseDaqDevice   '128
         FunctionString$ = "cbReleaseDaqDevice(BoardNum&)"
      Case GetBoardNumber   '129
         FunctionString$ = "cbGetBoardNumber(DeviceDescriptor)"
      Case GetNetDeviceDescriptor   '130
         FunctionString$ = "cbGetNetDeviceDescriptor(host$, Port&, DeviceDescriptor, Timeout&)"
      Case AInputMode   '131
         FunctionString$ = "cbAInputMode(BoardNum&, InputMode&)"
      Case AChanInputMode   '132
         FunctionString$ = "cbAChanInputMode(BoardNum&, Chan&, InputMode&)"
      Case WinArrayToBuf32   '133
         FunctionString$ = "cbWinArrayToBuf32(DataBuffer, MemHandle, FirstPoint, CBCount)"
      Case DInArray     '134
         FunctionString$ = "cbDInArray(BoardNum, FirstPort, LastPort, DataArray)"
      Case DOutArray   '135
         FunctionString$ = "cbDOutArray(BoardNum, FirstPort, LastPort, DataArray)"
      'Case DaqDeviceVersion   '136
      '   FunctionString$ = "cbDaqDeviceVersion(BoardNum, VersionType, Version!)"
      Case DIn32     '137
         FunctionString$ = "cbDIn32(BoardNum, PortNum, DataValue)"
      Case DOut32     '138
         FunctionString$ = "cbDOut32(BoardNum, PortNum, DataValue)"
      Case DClearAlarm  '139
         FunctionString$ = "DClearAlarm(BoardNum, PortNum, AlarmMask)"
      Case GPFind          '201
         FunctionString$ = "ibFind()"
      Case GPSend
         FunctionString$ = "Send(Device, Command, [Return], [Conditional])"
      Case GPReceive
         FunctionString$ = "Receive(Device)"
      Case GPTrigger
         FunctionString$ = "Trigger(Device)"
      Case GPDevClear
         FunctionString$ = "DevClear (DCL, [Conditional])"
      Case GPIBAsk
         FunctionString$ = "ibAsk()"
      Case GPInit
         FunctionString$ = "ibInit()"
      Case GPPtrs
         FunctionString$ = "ibPtrs()"
      Case GPSelDevClear
         FunctionString$ = "SelDevClear (SDC)"
      Case GPIBSre            '210
         FunctionString$ = "ibsre()"
      Case GPIBReturn         '211
         FunctionString$ = "GPReturnVal(ReturnString)"
         
      Case SSetBoardName      '2000
         FunctionString$ = "SetBoardName(BoardName$)"
      Case SShowDiag
         FunctionString$ = "ShowDialog(DiagText$, DialogType, VarName$, Title$, Default, [Conditional])"
      Case SContPlot
         FunctionString$ = "ContinuousPlot(TF%)"
      Case SConvData
         FunctionString$ = "Convert(TF%)"
      Case SConvPT
         FunctionString$ = "ConvertPT(TF%)"
      Case SSetData  '2005
         FunctionString$ = "DataType(MenuIndex&)"
      Case SSetAmplitude
         FunctionString$ = "SetAmplitude(Amplitude&)"
      Case SSetOffset
         FunctionString$ = "SetOffset(Offset&)"
      Case SCalCheck
         FunctionString$ = "SetCalMode(TF%)"
      Case SCountSet
         FunctionString$ = "SetCalCount(Points%)"
      Case SAddPTBuf '2010
         FunctionString$ = "AddPTBuffer(TF%)"
      Case SSetDevName
         FunctionString$ = "SetDeviceName(DevName$)"
      Case SSetPlotOpts
         FunctionString$ = "SetPlotOptions(PlotOptions%, TitleType%)"
      Case SBufInfo  '2013
         FunctionString$ = "UpdateBuffer()"
      Case SSetPlotChan '2014
         FunctionString$ = "SetPlotChannels(PlotChans%)"
      Case SNextBlock '2015
         FunctionString$ = "PlotNextBlock(Chan%)"
      Case SSetBlock  '2016
         FunctionString$ = "SetBlockSize(Size&, [Conditional])"
      Case SSetResolution '2018
         FunctionString$ = "SetPlotResolution(Res%)"
      Case SShowText  '2019
         FunctionString$ = "PlotText(TF%)"
      Case SGetTC  '2020
         FunctionString$ = "GetTCValues(TF%)"
      Case SToEng '2021
         FunctionString$ = "UseEngUnits(TF%)"
      Case SPlotType '2022
         FunctionString$ = "SetPlotType(PlotType%)"
      Case SCalcNoise '2023
         FunctionString$ = "CalcNoise(TF%)"
      Case SLogOutput   '2024
         FunctionString$ = "SLogOutput(ScreenOutput, FileOutput, FileName$)"
      Case SLoadStringList   '2025
         FunctionString$ = "SLoadStringList(FileName$, ListSizeVariable$)"
      Case SGetStringFromList '2026
         FunctionString$ = "SGetStringFromList(ListIndex, VariableToStore)"
      Case SSetPlotScaling    '2027
         FunctionString$ = "SSetPlotScaling(PlotScaleMode)"
      Case SSetFirstPlotPoint '2028
         FunctionString$ = "SSetFirstPlotPoint(FirstPlotPoint)"
      Case SLoadCSVList       '2029
         FunctionString$ = "SLoadCSVList(FileName$, ListSizeVariable$, NumCSVValues%, [ListName])"
      Case SEvalEnable '2030
         FunctionString$ = "EvaluateData(Enable%)"
      Case SEvalDelta '2031
         FunctionString$ = "EvaluateDelta(MaxDelta&)"
      Case SEvalMaxMin '2032
         FunctionString$ = "EvaluateMaxMin(MinValue&, MaxValue&)"
      Case SGetCSVsFromList '2033
         FunctionString$ = "SGetStringFromList(ListIndex, ValueIndex, VariableToStore, [ListName])"
      Case SCopyFile '2036
         FunctionString$ = "SCopyFile(FileSource, FileDesination)"
      Case SRunApp '2037
         FunctionString$ = "SRunApp(CommandLine$, Wait)"
      Case SEndApp '2038
         FunctionString$ = "SEndApp()"
      Case SResetConfig '2039
         FunctionString$ = "SResetConfig()"
      Case SEvalChannel '2040
         FunctionString$ = "EvaluationChannel(Channel%)"
      Case SGetFormRef  '2041
         FunctionString$ = "SGetFormRef(FormNumber)"
      Case SSelPortRange   '2042
         FunctionString$ = "SSelPortRange(FirstPortIndex, LastPortIndex)"
      Case SSetPortDirection  '2043
         FunctionString$ = "SSetPortDirection(Direction)"
      Case SReadPortRange  '2044
         FunctionString$ = "SReadPortRange(NumberOfReads)"
      Case SWritePortRange '2045
         FunctionString$ = "SWritePortRange(NumberOfBlocks)"
      Case SSelBitRange   '2046
         FunctionString$ = "SSelBitRange(FirstBit, LastBit)"
      Case SReadBitRange   '2047
         FunctionString$ = "SReadBitRange(NumberOfReads)"
      Case SWriteBitRange  '2048
         FunctionString$ = "SWriteBitRange(NumberOfBlocks)"
      Case SGenerateData   '2049
         FunctionString$ = "SGenerateData(DataType, Cycles, NumPoints, " & _
         "NumChans, Amplitude, Offset, SigType, NewData, Channel, FirstPoint)"
      Case SPlotGenData    '2050
         FunctionString$ = "SPlotGenData()"
      Case SPlotAcqData    '2051
         FunctionString$ = "SPlotAcqData()"
      Case SSetBitDirection  '2052
         FunctionString$ = "SSetBitDirection(Direction)"
      Case SWaitForIdle '2053
         FunctionString$ = "SWaitForIdle(StopOnCount, Timeout)"
      Case SWaitForEvent '2054
         FunctionString$ = "SWaitForEvent(EventType, WaitData, Timeout)"
      Case SWaitStatusChange  '2055
         FunctionString$ = "SWaitStatusChange(StopDelta, WaitCondition, Timeout)"
      Case SStopOnCount  '2056
         FunctionString$ = "SStopOnCount(StopCount, Timeout)"
      Case SPlotOnCount '2057
         FunctionString$ = "SPlotOnCount(PlotCount, Timeout)"
      Case SSetBitsPerPort ' 2058
         FunctionString$ = "SSetBitsPerPort(PortNum, CumBits, BitsInPort)"
      Case SCounterArm ' 2060
         FunctionString$ = "SCounterArm(CtrNum, EnableDisable, Conditional)"
         
      Case SDelay '3000
         FunctionString$ = "Delay(NumSeconds&)"
      Case SErrorPrint
         FunctionString$ = "PrintErrors(TF%)"
      Case SGetStatus
         FunctionString$ = "GetStatus(TF%)"
      Case SErrorFlow   '3003
         FunctionString$ = "ErrorFlow(Flow%)"
      Case SULErrFlow   '3004
         FunctionString$ = "SULErrFlow(ErrorHandling)"
      Case SULErrReport   '3005
         FunctionString$ = "SULErrReport(ErrorReporting)"
      Case SSetStaticOption   '3006
         FunctionString$ = "SSetStaticOption(StaticOptions, [Conditional])"
      Case SScriptRate    '3008
         FunctionString$ = "SScriptRate(MilliSeconds)"
      Case SSetVariable    '3009
         FunctionString$ = "SSetVariable(VarName, VarValue, [Conditional])"
      Case SGetFormProps   '3010
         FunctionString$ = "SGetFormProps(PropName, StoreVariable)"
      Case SGetStaticOptions  '3011
         FunctionString$ = "SGetStaticOptions()"
      Case SSetFormProps   '3012
         FunctionString$ = "SSetFormProps(PropName, PropVal)"
      Case SCloseApp    '3013
         FunctionString$ = "SCloseApp()"
      Case SSetVarDefault  '3014
         FunctionString$ = "SSetVarDefault(VarName, DefaultValue)"
      Case SPauseScript    '3015
         FunctionString$ = "SPauseScript([Conditional])"
      Case SGetParameterString    '3016
         FunctionString$ = "SGetParameterString(Function, ParamNumber, ParamValue, ReturnedString)"
      Case SPicklist    '3017
         FunctionString$ = "SPicklist(ListIndex, ListToPick, VariableToSet, [ListSize], [NumericDefault])"
      Case SGenRndVal   '3018
         FunctionString$ = "SGenRndVal(SeedValue, VariableToSet)"
      Case SPickGroup   '3019
         FunctionString$ = "SPickGroup(GroupIndicator$, GroupIndex&, GroupList$, ListToSet$, NumGroups&)"
      Case SSetLibType     '3020
         FunctionString$ = "SSetLibType(LibType)"
      Case SCalcMaxSinDelta     '3021
         FunctionString$ = "SCalcMaxSinDelta(Amplitude, SourceRate, IsPerChan, Result, [XClockRate])"
      Case SIsListed    '3022
         FunctionString$ = "SIsListed(ListToSearch, ValueToSearch, IsListed, [AtIndex], [ListSeparator])"
      Case SPeriodCalc    '3023
         FunctionString$ = "SPeriodCalc(Rate, Period)"
      Case SPulseWidthCalc  '3024
         FunctionString$ = "SPulseWidthCalc(Time, PulseWidth)"
      Case SMapAISwitch    '3025
         FunctionString$ = "SMapAISwitch(MapString)"
      Case SStopScript    '3026
         FunctionString$ = "SStopScript([Conditional])"
      Case SSetMCCControl  '3027
         FunctionString$ = "SSetMCCControl(TrueFalse)"
      Case SGetDP8200Cmd  '3028
         FunctionString$ = "SGetDP8200Cmd(Value, Command)"
      Case SEvalParamRev  '3029
         FunctionString$ = "SEvalParamRev(Value, EvalCondition)"
      Case SLoadSubScript '4001
         FunctionString$ = "OpenSubscript(DupeBoard, ScriptName, BoardName, [Condition])"
      Case SCloseSubScript '4002
         FunctionString$ = "CloseSubScript(ScriptName)"

      Case SOpenWindow '5001
         FunctionString$ = "OpenWindow(WindowType, BoardName)"
      Case SCloseWindow '5002
         FunctionString$ = "CloseWindow(WindowType)"
      Case WinAPIGlobalAlloc
         FunctionString$ = "GlobalAlloc(wFlags, dwBytes)"
      Case WinAPIGlobalFree
         FunctionString$ = "GlobalFree(MemHandle)"
      Case WinAPICreateFileMapping
         FunctionString$ = "CreateFileMapping(hFile&, lpAttributes, flProtect&, dwMaxSizeHigh&, dwMaxSizeLow&, lpName$)"
      Case WinAPIMapViewOfFile
         FunctionString$ = "MapViewOfFile(hFileMappingObject&, dwDesiredAccess&, dwFileOffsetHigh&, dwFileOffsetLow&, dwNumberOfBytesToMap&)"
      Case WinAPIUnmapViewOfFile
         FunctionString$ = "UnmapViewOfFile(Handle&)"
      Case WinAPIOpenFileMapping
         FunctionString$ = "OpenFileMapping(dwDesiredAccess&, bInheritHandle&, lpName$)"
      
      Case USBBlink '= 8001
         FunctionString$ = "cbUSBBlink(DeviceNum&)"
      Case USBReset '= 8002
         FunctionString$ = "cbUSBReset(DeviceNum&)"
      Case USBGetSerialNum '= 8003
         FunctionString$ = "cbUSBGetSerialNum(DeviceNum&, SerialNum$)"
      Case USBSetSerialNum '= 8004
         FunctionString$ = "cbUSBSetSerialNum(DeviceNum&, SerialNum$)"
      Case USBMemRead '= 8005
         FunctionString$ = "cbUSBMemRead(DeviceNum&, address&, mccData&, count&)"
      Case USBMemWrite '= 8006
         FunctionString$ = "cbUSBMemWrite(DeviceNum&, address&, mccData&, count&)"
      Case USBWatchdog '= 8007
         FunctionString$ = "cbUSBWatchdog(DeviceNum&, Status&, timeout&, Action&, Channel&, State&)"
      Case USBAInScan ' 8008
         FunctionString$ = "cbUSBAInScan(DeviceNum&, StartChan&, EndChan&, Count&, Rate&, Gain&, Data%, Options&)"
      Case USBALoadQueue ' 8009
         FunctionString$ = "cbUSBALoadQueue(DeviceNum&, ChanArray%, GainArray%, Count&)"
      Case USBAIn ' 8010
         FunctionString$ = "cbUSBAIn(DeviceNum&, Channel&, Gain&, DataValue%)"
      Case USBAOut '= 8011
         FunctionString$ = "cbUSBAOut(DeviceNum&, Channel&, Gain&, Data%)"
      Case USBDConfigPort '= 8012
         FunctionString$ = "cbUSBDConfigPort(DeviceNum&, PortNum&, Direction&)"
      Case USBDIn '= 8013
         FunctionString$ = "cbUSBDIn(DeviceNum&, PortNum&, DataValue%)"
      Case USBDOut '= 8014
         FunctionString$ = "cbUSBDOut(DeviceNum&, PortNum&, DataValue%)"
      Case USBDBitIn '= 8015
         FunctionString$ = "cbUSBDBitIn(DeviceNum&, PortNum&, BitNum&, BitValue%)"
      Case USBDBitOut '= 8016
         FunctionString$ = "cbUSBDBitOut(DeviceNum&, PortNum&, BitNum&, BitValue%)"
      Case USBGetStatus '= 8017
         FunctionString$ = "cbUSBGetStatus(DeviceNum&, Status%, CurCount&, CurIndex&)"
      Case USBGetErrMsg '= 8018
         FunctionString$ = "cbUSBGetErrMsg(ErrCode&, ErrMsg$)"
      Case USBDConfigBit '= 8019
         FunctionString$ = "cbUSBDConfigBit(DeviceNum&, PortNum&, BitNum&, Direction&)"
      Case USBStopBackground '= 8021
         FunctionString$ = "cbUSBStopBackground(DeviceNum&)"
      Case USBCInit '= 8022
         FunctionString$ = "cbUSBCInit(DeviceNum&)"
      Case USBCIn32 '= 8023
         FunctionString$ = "cbUSBCIn32(DeviceNum&, CounterNum&, Data&)"
      Case USBFromEngUnits '= 8024
         FunctionString$ = "cbUSBFromEngUnits(DeviceNum&, Gain&, EngUnits!, DataVal%)"
      Case USBToEngUnits '= 8025
         FunctionString$ = "cbUSBToEngUnits(DeviceNum&, Gain&, DataVal%, EngUnits!)"
      Case USBDSetTrig '= 8026
         FunctionString$ = "cbUSBDSetTrig(DeviceNum&, TrigType&, Channel&)"
      Case USBSaveConfig '= 8027
         FunctionString$ = "cbUSBSaveConfig(FileName$)"
      Case EStatus   ' 10001
         FunctionString$ = "EStatus(Condition, Limit1, Limit2, FailIfTimeout)"
      Case ETimeStamp '10002
         FunctionString$ = "ETimeStamp()"
      Case EEventType   ' 10003
         FunctionString$ = "EEventType(EventType, EventData, FailIfData, " & _
         "NoEvent, FailIfTimeout)"
      Case EDataDC   ' 10010
         FunctionString$ = "EDataDC(DataPoints, EvalChan%, DCValue, ATol&, DCOption)"
      Case EDataPulse   ' 10011
         FunctionString$ = "EDataPulse(DataPoints, " & _
         "EvalChan, HiVal, LoVal, ATol, HiBy, LoBy, TTol, Repeat)"
      Case EDataDelta   ' 10012
         FunctionString$ = "EDataDelta(DataPoints, EvalChan, " & _
         "ValueOfChange, FailIfDelta, MovingAverage)"
      Case EDataAmplitude  '10013
         FunctionString$ = "EDataAmplitude(DataPoints, EvalChan, " & _
         "AmplitudeValue, ATolerance, MovingAverage)"
      Case SEvalTrigPoint  '10015
         FunctionString$ = "SEvalTrigPoint(DataPoints, EvalChan, " & _
         "TrigPolarity, Threshold [Volts], Guardband, Tolerance [S], [EvalOption])"
      Case EDataTime '10021
         FunctionString$ = "EDataTime(DataPoints, EvalChan, Threshold, " & _
         "Guardband, SourceFreq, Rate [S/s], Tolerance [S/~], EvaluationOption)"
      Case EDataOutVsIn '10022
         FunctionString$ = "EDataOutVsIn(DataPoints, EvalChan, " & _
         "FailIfComparison, OutputFormRef, NumberOfBits)"
      Case EDataSkew '10023
         FunctionString$ = "EDataSkew(DataPoints, NumChans, " & _
         "Threshold [Volts], Guardband [±LSBs], " & _
         "Tolerance [S], EvaluationOption)"
      Case EData32Delta  '10030
         FunctionString$ = "EData32Delta(DataPoints, EvalChan, " & _
         "ValueOfChange, FailIfDelta, DeltaOption)"
      Case EError
         FunctionString$ = "EError(Function, ExpectedError, " & _
         "Alternate1, Alternate2, Alternate3, Action)"
      Case EHistogram   '10051
         FunctionString$ = "EHistogram(BinSpread, MaxRMSValue, " & _
         "AverageValue, AvgValTol)"
      Case EStore488Value '10052
         FunctionString$ = "EStore488Value()"
      Case ECompareStoredValue '10053
         FunctionString$ = "ECompareStoredValue(ExpectedValue, ErrorUnits, " & _
         "Tolerance)"
      Case Else
         FunctionString$ = "Undefined function"
   End Select
   If gnThreading Then
      If Left(FunctionString$, 2) = "cb" Then
         GetFunctionString = Mid(FunctionString$, 3)
      End If
   Else
      GetFunctionString = FunctionString$
   End If

End Function

Function GetRangeString(ByVal Range As Long) As String
   
   Select Case Range
      Case -2
         RangeString$ = ""
      Case NOTUSED - 1
         RangeString$ = "NOTUSED"
      Case BIP5VOLTS '0
         RangeString$ = "BIP5VOLTS"
      Case BIP10VOLTS   '1
         RangeString$ = "BIP10VOLTS"
      Case BIP2PT5VOLTS '2
         RangeString$ = "BIP2PT5VOLTS"
      Case BIP1PT25VOLTS   '3
         RangeString$ = "BIP1PT25VOLTS"
      Case BIP1VOLTS '4
         RangeString$ = "BIP1VOLTS"
      Case BIPPT625VOLTS   '5
         RangeString$ = "BIPPT625VOLTS"
      Case BIPPT5VOLTS  '6
         RangeString$ = "BIPPT5VOLTS"
      Case BIPPT1VOLTS  '7
         RangeString$ = "BIPPT1VOLTS"
      Case BIPPT05VOLTS '8
         RangeString$ = "BIPPT05VOLTS"
      Case BIPPT01VOLTS '9
         RangeString$ = "BIPPT01VOLTS"
      Case BIPPT005VOLTS   '10
         RangeString$ = "BIPPT005VOLTS"
      Case BIP1PT67VOLTS   '11
         RangeString$ = "BIP1PT67VOLTS"
      Case BIPPT312VOLTS   '17
         RangeString$ = "BIPPT312VOLTS"
      Case BIPPT156VOLTS   '18
         RangeString$ = "BIPPT156VOLTS"
      Case BIPPT078VOLTS   '19
         RangeString$ = "BIPPT078VOLTS"
      Case BIP60VOLTS   '20
         RangeString$ = "BIP60VOLTS"
      Case BIP15VOLTS   '21
         RangeString$ = "BIP15VOLTS"
      Case BIPPT125VOLTS   '22
         RangeString$ = "BIPPT125VOLTS"
      Case BIP30VOLTS   '23
         RangeString$ = "BIP30VOLTS"
      Case BIPPT25VOLTS '12
         RangeString$ = "BIPPT25VOLTS"
      Case BIPPT2VOLTS  '13
         RangeString$ = "BIPPT2VOLTS"
      Case BIP2VOLTS '14
         RangeString$ = "BIP2VOLTS"
      Case BIP20VOLTS   '15
         RangeString$ = "BIP20VOLTS"
      Case BIP4VOLTS '16
         RangeString$ = "BIP4VOLTS"
      Case BIPPT073125VOLTS   '73
         RangeString$ = "BIPPT073125VOLTS"
      Case UNI10VOLTS   '100
         RangeString$ = "UNI10VOLTS"
      Case UNI5VOLTS '101
         RangeString$ = "UNI5VOLTS"
      Case UNI2PT5VOLTS '102
         RangeString$ = "UNI2PT5VOLTS"
      Case UNI2VOLTS '103
         RangeString$ = "UNI2VOLTS"
      Case UNI1PT25VOLTS   '104
         RangeString$ = "UNI1PT25VOLTS"
      Case UNI1VOLTS '105
         RangeString$ = "UNI1VOLTS"
      Case UNIPT1VOLTS  '106
         RangeString$ = "UNIPT1VOLTS"
      Case UNIPT01VOLTS '107
         RangeString$ = "UNIPT01VOLTS"
      Case UNIPT02VOLTS '108
         RangeString$ = "UNIPT02VOLTS"
      Case UNI1PT67VOLTS   '109
         RangeString$ = "UNI1PT67VOLTS"
      Case UNIPT5VOLTS  '110
         RangeString$ = "UNIPT5VOLTS"
      Case UNIPT25VOLTS '111
         RangeString$ = "UNIPT25VOLTS"
      Case UNIPT2VOLTS  '112
         RangeString$ = "UNIPT2VOLTS"
      Case UNIPT05VOLTS '113
         RangeString$ = "UNIPT05VOLTS"
      Case UNI4VOLTS '114
         RangeString$ = "UNI4VOLTS"
      Case MA4TO20   '200
         RangeString$ = "MA4TO20"
      Case MA2to10   '201
         RangeString$ = "MA2to10"
      Case MA1TO5 '202
         RangeString$ = "MA1TO5"
      Case MAPT5TO2PT5  '203
         RangeString$ = "MAPT5TO2PT5"
      Case MA0TO20   '204
         RangeString$ = "MA0TO20"
      Case BIPPT025AMPS '205
         RangeString$ = "BIPPT025AMPS"
      Case BIPPT025VOLTSPERVOLT  '400
         RangeString$ = "BIPPT025VOLTSPERVOLT"
      Case Else
         RangeString$ = "Custom"
   End Select
   GetRangeString = RangeString$

End Function

Function GetOptionsString(OptionVal As Long, FormType As Integer, Optional FormFunction As Variant) As String

   'count changes if options are added
   For i% = 0 To 9
      If (OptionVal And 2 ^ i%) = 2 ^ i% Then
         If ((OptionVal And 2 ^ i%) = SINGLEIO) Or ((OptionVal And 2 ^ i%) = DMAIO) Then
            If (OptionVal And BLOCKIO) = BLOCKIO Then
               If (2 ^ i%) > SINGLEIO Then opt = "BLOCKIO"
            ElseIf (OptionVal And DMAIO) = DMAIO Then
               opt = "DMAIO"
            Else
               opt = "SINGLEIO"
            End If
         Else
            opt = Choose(i% + 1, "BACKGROUND", "CONTINUOUS", "EXTCLOCK", "CONVERTDATA", _
            "SCALEDATA", "SINGLEIO", "DMAIO", "BLOCKIO", "WORDXFER", "SIMULTANEOUS")
         End If
         If Not IsNull(opt) Then
            If FormType = COUNTER Then
               If opt = "WORDXFER" Then opt = "CTR32BIT"
               If opt = "NOFILTER" Then opt = "CTR64BIT"
               If opt = "SIMULTANEOUS" Then opt = "CTR48BIT"
            End If
            If (FormType = DIGITAL_IN) Or (FormType = DIGITAL_OUT) _
               Then If opt = "SIMULTANEOUS" Then opt = "DWORDXFER"
            'If (FormType = ANALOG_OUT) Then
            '   If FormFunction Then If opt = "SIMULTANEOUS" Then opt = "DWORDXFER"
            'End If
            If Len(opt) > 0 Then Options$ = Options$ & opt & " "
         End If
         opt = ""
      End If
   Next i%
   For i% = 10 To 22  'this number changes if new options are added
      If (OptionVal And 2 ^ i%) = 2 ^ i% Then
         opt = Choose(i% - 9, "NOFILTER", "EXTMEMORY", "BURSTMODE", _
         "WAITFORNEWDATA", "EXTTRIGGER", "NOCALIBRATEDATA", "BURSTIO", "RETRIGMODE", _
         "NONSTREAMEDIO", "ADCCLOCKTRIG", "ADCCLOCK", "HIGHRESRATE", "SHUNTCAL")
         If Not IsNull(opt) Then
            If FormType = COUNTER Then
               If opt = "NOFILTER" Then opt = "CTR64BIT"
               If opt = "EXTMEMORY" Then opt = "NOCLEAR"
            Else 'If Not (FormType = ANALOG_IO) Then
               If opt = "EXTMEMORY" Then opt = "NOCLEAR"
            End If
            Options$ = Options$ & opt & " "
         End If
      End If
   Next i%
   If (OptionVal And 2 ^ 30) = 2 ^ 30 Then
      opt = "BLOCKIO"
      Options$ = Options$ & opt
   End If
   If Len(Options$) = 0 Then Options$ = "DEFAULTIO"
   GetOptionsString = Options$

End Function

Function GetCtrEdgeString(CtrEdge As Long) As String

   'Select Case CtrEdge
   '   Case CTR_RISING_EDGE
   '      Reply$ = "CTR_RISING_EDGE"
   '   Case CTR_FALLING_EDGE
   '      Reply$ = "CTR_FALLING_EDGE"
   'End Select
   If Not CtrEdge = 0 Then
      For CodeVal% = 0 To 2
         If (CtrEdge And 2 ^ CodeVal%) = 2 ^ CodeVal% Then
            EdgeState$ = Choose(CodeVal% + 1, "-CtrEdge, ", "-BEdge, ", "-ZEdge, ")
            EdgeString$ = EdgeString$ & EdgeState$
         End If
      Next
      EdgeString$ = Left(EdgeString$, Len(EdgeString$) - 2)
   Else
      EdgeString$ = "Positive Edge"
   End If
   GetCtrEdgeString = EdgeString$
   
End Function

Function GetFlagPinString(FlagPin As Long) As String

   Select Case FlagPin
      Case CARRY_BORROW
         Reply$ = "CARRYBORROW"
      Case COMPARE_BORROW
         Reply$ = "COMPAREBORROW"
      Case CARRYBORROW_UPDOWN
         Reply$ = "CARRYBORROWUPDOWN"
      Case INDEX_ERROR
         Reply$ = "INDEXERROR"
      Case Else
         Reply$ = "Invalid Flag Pin value"
   End Select
   GetFlagPinString = Reply$

End Function

Function GetQuadString(QuadMode As Long) As String

   Select Case QuadMode
      Case NO_QUAD
         Reply$ = "NOQUAD"
      Case X1_QUAD
         Reply$ = "X1QUAD"
      Case X2_QUAD
         Reply$ = "X2QUAD"
      Case X4_QUAD
         Reply$ = "X4QUAD"
      Case Else
         Reply$ = "Invalid Quad Mode"
   End Select
   GetQuadString = Reply$

End Function

Function GetCountingModeString(CountMode As Long) As String

   Select Case CountMode
      Case NORMAL_MODE
         Reply$ = "NORMALMODE"
      Case RANGE_LIMIT
         Reply$ = "RANGELIMIT"
      Case NO_RECYCLE
         Reply$ = "NORECYCLE"
      Case MODULO_N
         Reply$ = "MODULON"
      Case Else
         Reply$ = "Invalid Quad Mode"
   End Select
   GetCountingModeString = Reply$

End Function

Function GetIndexModeString(IndexMode As Long) As String

   Select Case IndexMode
      Case INDEX_DISABLED
         Reply$ = "INDEXDISABLED"
      Case LOAD_CTR
         Reply$ = "LOADCTR"
      Case LOAD_OUT_LATCH
         Reply$ = "LOADOUTLATCH"
      Case RESET_CTR
         Reply$ = "RESETCTR"
      Case Else
         Reply$ = "Invalid Mode"
   End Select
   GetIndexModeString = Reply$

End Function

Function GetCtrModeString(CtrMode As Long) As String

   TotalizeMode% = True
   For i% = 1 To 4
      OptVal& = Choose(i%, PERIOD, PULSEWIDTH, TIMING, ENCODER)
      If (CtrMode And &HF00) = OptVal& Then
         PrimModeString$ = Choose(i%, "PERIOD", "PULSEWIDTH", "TIMING", "ENCODER") & ", "
         TotalizeMode% = False
         Exit For
      End If
   Next
   If TotalizeMode% Then OptVal& = TOTALIZE
   Select Case OptVal&
      Case TOTALIZE
         For i% = 1 To 11
            SubMode& = Choose(i%, CLEAR_ON_READ, STOP_AT_MAX, DECREMENT_ON, _
            BIT_32, GATING_ON, LATCH_ON_MAP, UPDOWN_ON, RANGE_LIMIT_ON, _
            NO_RECYCLE_ON, MODULO_N_ON, BIT_48)
            If (SubMode& And CtrMode) = SubMode& Then
               SubModeString$ = SubModeString$ & Choose(i%, "CLEAR_ON_READ", _
               "STOP_AT_MAX", "DECREMENT_ON", "BIT_32", "GATING_ON", _
               "LATCH_ON_MAP", "UPDOWN_ON", "RANGE_LIMIT_ON", _
               "NO_RECYCLE_ON", "MODULO_N_ON", "BIT_48") & ", "
            End If
         Next
      Case PERIOD
         For i% = 1 To 3
            SubMode& = Choose(i%, PERIOD_MODE_X10, _
            PERIOD_MODE_X100, PERIOD_MODE_X1000)
            If (SubMode& And CtrMode) = SubMode& Then
               SubModeString$ = Choose(i%, "PERIOD_MODE_X10", _
               "PERIOD_MODE_X100", PERIOD_MODE_X1000) & ", "
            End If
         Next
      Case PULSEWIDTH
         For i% = 1 To 2
            SubMode& = Choose(i%, PULSEWIDTH_MODE_BIT_32, PULSEWIDTH_MODE_GATING_ON)
            If (SubMode& And CtrMode) = SubMode& Then
               SubModeString$ = SubModeString$ & Choose(i%, _
               "PULSEWIDTH_MODE_BIT_32", "PULSEWIDTH_MODE_GATING_ON") & ", "
            End If
         Next
      Case TIMING
         SubMode& = TIMING_MODE_BIT_32
         If (SubMode& And CtrMode) = SubMode& Then
            SubModeString$ = "TIMING_MODE_BIT_32" & ", "
         End If
      Case ENCODER
         For i% = 1 To 8
            SubMode& = Choose(i%, ENCODER_MODE_X2, ENCODER_MODE_X4, _
            BIT_32, ENCODER_MODE_LATCH_ON_Z, ENCODER_MODE_CLEAR_ON_Z_ON, _
            ENCODER_MODE_RANGE_LIMIT_ON, ENCODER_MODE_NO_RECYCLE_ON, _
            ENCODER_MODE_MODULO_N_ON, BIT_48)
            If (SubMode& And CtrMode) = SubMode& Then
               SubModeString$ = SubModeString$ & Choose(i%, "ENCODER_MODE_X2", _
               "ENCODER_MODE_X4", "BIT_32", "ENCODER_MODE_LATCH_ON_Z", _
               "ENCODER_MODE_CLEAR_ON_Z_ON", "ENCODER_MODE_RANGE_LIMIT_ON", _
               "ENCODER_MODE_NO_RECYCLE_ON", "ENCODER_MODE_MODULO_N_ON", "BIT_48") & ", "
            End If
         Next
   End Select
   
   ModeString$ = PrimModeString$ & SubModeString$
   If Len(ModeString$) = 0 Then Exit Function
   ModeString$ = Left(ModeString$, Len(ModeString$) - 2)
   GetCtrModeString = ModeString$

End Function

Function GetTrigPolarity(TrigTypeVal As Long) As String

   Select Case TrigTypeVal
      Case TRIGABOVE
         TrigTypeString$ = "TRIGABOVE"
      Case TRIGBELOW
         TrigTypeString$ = "TRIGBELOW"
      Case GATENEGHYS
         TrigTypeString$ = "GATENEGHYS"
      Case GATEPOSHYS
         TrigTypeString$ = "GATEPOSHYS"
      Case GATEABOVE
         TrigTypeString$ = "GATEABOVE"
      Case GATEBELOW
         TrigTypeString$ = "GATEBELOW"
      Case GATEINWINDOW
         TrigTypeString$ = "GATEINWINDOW"
      Case GATEOUTWINDOW
         TrigTypeString$ = "GATEOUTWINDOW"
      Case GATEHIGH
         TrigTypeString$ = "GATEHIGH"
      Case GATELOW
         TrigTypeString$ = "GATELOW"
      Case TRIGHIGH
         TrigTypeString$ = "TRIGHIGH"
      Case TRIGLOW
         TrigTypeString$ = "TRIGLOW"
      Case TRIGPOSEDGE
         TrigTypeString$ = "TRIGPOSEDGE"
      Case TRIGNEGEDGE
         TrigTypeString$ = "TRIGNEGEDGE"
   End Select
   GetTrigPolarity = TrigTypeString$

End Function

Function GetTrigTypeString(TrigTypeVal As Long) As String

   Select Case TrigTypeVal
      Case TRIGABOVE
         TrigTypeString$ = "TRIGABOVE"
      Case TRIGBELOW
         TrigTypeString$ = "TRIGBELOW"
      Case GATENEGHYS
         TrigTypeString$ = "GATENEGHYS"
      Case GATEPOSHYS
         TrigTypeString$ = "GATEPOSHYS"
      Case GATEABOVE
         TrigTypeString$ = "GATEABOVE"
      Case GATEBELOW
         TrigTypeString$ = "GATEBELOW"
      Case GATEINWINDOW
         TrigTypeString$ = "GATEINWINDOW"
      Case GATEOUTWINDOW
         TrigTypeString$ = "GATEOUTWINDOW"
      Case GATEHIGH
         TrigTypeString$ = "GATEHIGH"
      Case GATELOW
         TrigTypeString$ = "GATELOW"
      Case TRIGHIGH
         TrigTypeString$ = "TRIGHIGH"
      Case TRIGLOW
         TrigTypeString$ = "TRIGLOW"
      Case TRIGPOSEDGE
         TrigTypeString$ = "TRIGPOSEDGE"
      Case TRIGNEGEDGE
         TrigTypeString$ = "TRIGNEGEDGE"
   End Select
   GetTrigTypeString = TrigTypeString$
   
End Function

Function GetFailIfDeltaString(FailIfDeltaVal As Long) As String

   Select Case FailIfDeltaVal
      Case 0
         Reply$ = "Exceeds ValueOfChange"
      Case 1
         Reply$ = "DoesNotExceed ValueOfChange"
      Case 2
         Reply$ = "IsNotEqualTo ValueOfChange"
      Case 3
         Reply$ = "IsPositive Excluding ValueOfChange (rollover)"
      Case 4
         Reply$ = "IsNegative Excluding ValueOfChange (rollover)"
      Case 5
         Reply$ = "IntervalOfNoChange < ValueOfChange Samples"
      Case 6
         Reply$ = "At least one change = ValueOfChange (± Tolerance optional DeltaOption)"
   End Select
   GetFailIfDeltaString = Reply$
   
End Function

Function GetEventTypeString(EventTypeVal As Long) As String

   If EventTypeVal = ALL_EVENT_TYPES Then
      Reply$ = "ALL_EVENT_TYPES"
   Else
      For eType% = 0 To 6
         If (2 ^ eType% And EventTypeVal) = 2 ^ eType% Then
            EventString$ = Choose(eType% + 1, "ON_SCAN_ERROR", _
            "ON_EXTERNAL_INTERRUPT", "ON_PRETRIGGER", "ON_DATA_AVAILABLE", _
            "ON_END_OF_AI_SCAN", "ON_END_OF_AO_SCAN", "ON_CHANGE_DI")
            Reply$ = Reply$ & EventString$ & ", "
         End If
      Next
      If Len(Reply$) > 2 Then Reply$ = Left(Reply$, Len(Reply$) - 2)
   End If
   GetEventTypeString = Reply$
   
End Function

Function GetChannelTypeString(ChannelType As Long) As String

   TypeOfChan% = ChannelType 'And &HFF
   'SetPoint% = ChannelType And &H100
   SetPoint% = (ChannelType = 11)
   Select Case TypeOfChan%
      Case -2
         Reply$ = "Set Value Only"
      Case -1
         Reply$ = "Queue Only - Load All Values"
      Case ANALOG
         Reply$ = "ANALOG"
      Case DIGITAL8
         Reply$ = "DIGITAL8"
      Case DIGITAL16
         Reply$ = "DIGITAL16"
      Case CTR16
         Reply$ = "CTR16"
      Case CTR32LOW
         Reply$ = "CTR32LOW"
      Case CTR32HIGH
         Reply$ = "CTR32HIGH"
      Case CJC
         Reply$ = "CJC"
      Case TC
         Reply$ = "TC"
      Case ANALOG_SE
         Reply$ = "ANALOG_SE"
      Case ANALOG_DIFF
         Reply$ = "ANALOG_DIFF"
      Case SETPOINTSTATUS
         Reply$ = "SETPOINTSTATUS"
      Case CTRBANK0
         Reply$ = "CTRBANK0"
      Case CTRBANK1
         Reply$ = "CTRBANK1"
      Case CTRBANK2
         Reply$ = "CTRBANK2"
      Case CTRBANK3
         Reply$ = "CTRBANK3"
      Case PADZERO
         Reply$ = "PADZERO"
   End Select
   If Not (SetPoint% = 0) Then SPVal$ = "SP "
   GetChannelTypeString = SPVal$ & Reply$
   
End Function

Function GetPortString(ByVal PortNum As Long) As String

   Select Case PortNum
      Case AUXPORT
         Reply$ = "AUXPORT"
      Case AUXPORT1
         Reply$ = "AUXPORT1"
      Case AUXPORT2
         Reply$ = "AUXPORT2"
      Case FIRSTPORTA
         Reply$ = "FIRSTPORTA"
      Case FIRSTPORTB
         Reply$ = "FIRSTPORTB"
      Case FIRSTPORTCL
         Reply$ = "FIRSTPORTCL"
      Case FIRSTPORTC
         Reply$ = "FIRSTPORTC"
      Case FIRSTPORTCH
         Reply$ = "FIRSTPORTCH"
      Case SECONDPORTA
         Reply$ = "SECONDPORTA"
      Case SECONDPORTB
         Reply$ = "SECONDPORTB"
      Case SECONDPORTCL
         Reply$ = "SECONDPORTCL"
      Case SECONDPORTCH
         Reply$ = "SECONDPORTCH"
      Case THIRDPORTA
         Reply$ = "THIRDPORTA"
      Case THIRDPORTB
         Reply$ = "THIRDPORTB"
      Case THIRDPORTCL
         Reply$ = "THIRDPORTCL"
      Case THIRDPORTCH
         Reply$ = "THIRDPORTCH"
      Case FOURTHPORTA
         Reply$ = "FOURTHPORTA"
      Case FOURTHPORTB
         Reply$ = "FOURTHPORTB"
      Case FOURTHPORTCL
         Reply$ = "FOURTHPORTCL"
      Case FOURTHPORTCH
         Reply$ = "FOURTHPORTCH"
      Case FIFTHPORTA
         Reply$ = "FIFTHPORTA"
      Case FIFTHPORTB
         Reply$ = "FIFTHPORTB"
      Case FIFTHPORTCL
         Reply$ = "FIFTHPORTCL"
      Case FIFTHPORTCH
         Reply$ = "FIFTHPORTCH"
      Case SIXTHPORTA
         Reply$ = "SIXTHPORTA"
      Case SIXTHPORTB
         Reply$ = "SIXTHPORTB"
      Case SIXTHPORTCL
         Reply$ = "SIXTHPORTCL"
      Case SIXTHPORTCH
         Reply$ = "SIXTHPORTCH"
      Case SEVENTHPORTA
         Reply$ = "SEVENTHPORTA"
      Case SEVENTHPORTB
         Reply$ = "SEVENTHPORTB"
      Case SEVENTHPORTCL
         Reply$ = "SEVENTHPORTCL"
      Case SEVENTHPORTCH
         Reply$ = "SEVENTHPORTCH"
      Case EIGHTHPORTA
         Reply$ = "EIGHTHPORTA"
      Case EIGHTHPORTB
         Reply$ = "EIGHTHPORTB"
      Case EIGHTHPORTCL
         Reply$ = "EIGHTHPORTCL"
      Case EIGHTHPORTCH
         Reply$ = "EIGHTHPORTCH"
      Case Else
         Reply$ = "INVALIDPORT"
   End Select
   GetPortString = Reply$
   
End Function

Function GetErrorUnits(ArgVal As Long) As String

   Select Case ArgVal
      Case 0
         Reply$ = "Volts"
      Case 1
         Reply$ = "LSBs"
   End Select
   GetErrorUnits = Reply$
   
End Function

Function GetDirectionString(Direction As Long) As String

   Select Case Direction
      Case DIGITALOUT
         Reply$ = "DIGITALOUT"
      Case DIGITALIN
         Reply$ = "DIGITALIN"
   End Select
   GetDirectionString = Reply$

End Function

Function GetScaleString(ScaleVal As Long) As String

   Select Case ScaleVal
      Case CELSIUS
         Reply$ = "CELSIUS"
      Case FAHRENHEIT
         Reply$ = "FAHRENHEIT"
      Case KELVIN
         Reply$ = "KELVIN"
      Case VOLTS
         Reply$ = "VOLTS"
      Case NOSCALE
         Reply$ = "NOSCALE"
   End Select
   GetScaleString = Reply$
   
End Function

Function GetDataTypeString(DataType As Long) As String

   Select Case DataType
      Case 1
         Reply$ = "Integer"
      Case 2
         Reply$ = "Long"
      Case Else
         Reply$ = "Not Assigned"
   End Select
   GetDataTypeString = Reply$

End Function

Function GetSigTypeString(SigType As Long) As String

   Select Case SigType
      Case 0
         Reply$ = "DC Level"
      Case 1
         Reply$ = "Square Wave"
      Case 2
         Reply$ = "Sine Wave"
      Case 3
         Reply$ = "Ramp"
      Case 4
         Reply$ = "Triangle Wave"
      Case 5
         Reply$ = "Random Signal"
   End Select
   GetSigTypeString = Reply$
   
End Function

Function GetSigDirString(SignalDirection As Long) As String

   Select Case SignalDirection
      Case CBDISABLED
         Reply$ = "DISABLED"
      Case SIGNAL_IN
         Reply$ = "SIGNAL_IN"
      Case SIGNAL_OUT
         Reply$ = "SIGNAL_OUT"
   End Select
   GetSigDirString = Reply$

End Function

Function GetSignalString(Signal As Long) As String

   Select Case Signal
      Case ADC_CONVERT
         Reply$ = "ADC_CONVERT"
      Case ADC_GATE
         Reply$ = "ADC_GATE"
      Case ADC_START_TRIG
         Reply$ = "ADC_START_TRIG"
      Case ADC_STOP_TRIG
         Reply$ = "ADC_STOP_TRIG"
      Case ADC_TB_SRC
         Reply$ = "ADC_TB_SRC"
      Case ADC_SCANCLK
         Reply$ = "ADC_SCANCLK"
      Case ADC_SSH
         Reply$ = "ADC_SSH"
      Case ADC_STARTSCAN
         Reply$ = "ADC_STARTSCAN"
      Case ADC_SCAN_STOP
         Reply$ = "ADC_SCAN_STOP"
      Case DAC_UPDATE
         Reply$ = "DAC_UPDATE"
      Case DAC_TB_SRC
         Reply$ = "DAC_TB_SRC"
      Case DAC_START_TRIG
         Reply$ = "DAC_START_TRIG"
      Case SYNC_CLK
         Reply$ = "SYNC_CLK"
      Case CTR1_CLK
         Reply$ = "CTR1_CLK"
      Case CTR2_CLK
         Reply$ = "CTR2_CLK"
      Case DGND
         Reply$ = "DGND"
   End Select
   GetSignalString = Reply$
   
End Function


Function GetConnectionString(Connection As Long) As String

   Select Case Connection
      Case AUXIN0
         Reply$ = "AUXIN0"
      Case AUXIN1
         Reply$ = "AUXIN1"
      Case AUXIN2
         Reply$ = "AUXIN2"
      Case AUXIN3
         Reply$ = "AUXIN3"
      Case AUXIN4
         Reply$ = "AUXIN4"
      Case AUXIN5
         Reply$ = "AUXIN5"
      Case AUXOUT0
         Reply$ = "AUXOUT0"
      Case AUXOUT1
         Reply$ = "AUXOUT1"
      Case AUXOUT2
         Reply$ = "AUXOUT2"
      Case DS_CONNECTOR
         Reply$ = "ADC_CONVERT"
   End Select
   GetConnectionString = Reply$
   
End Function

Function GetPolarityString(Polarity As Long) As String

   Select Case Polarity
      Case INVERTED
         Reply$ = "INVERTED"
      Case NONINVERTED
         Reply$ = "NONINVERTED"
   End Select
   GetPolarityString = Reply$

End Function

Function Get8254ConfigString(ConfigVal As Long) As String

   Select Case ConfigVal
      Case HIGHONLASTCOUNT
         Reply$ = "HIGHONLASTCOUNT"
      Case ONESHOT
         Reply$ = "ONESHOT"
      Case RATEGENERATOR
         Reply$ = "RATEGENERATOR"
      Case SQUAREWAVE
         Reply$ = "SQUAREWAVE"
      Case SOFTWARESTROBE
         Reply$ = "SOFTWARESTROBE"
      Case HARDWARESTROBE
         Reply$ = "HARDWARESTROBE"
   End Select
   Get8254ConfigString = Reply$

End Function

Function GetLibraryErrReportingString(LibraryErrReporting As Long) As String

   Select Case LibraryErrReporting
      Case -1
         Reply$ = "Reset to Original"
      Case 0
         Reply$ = "DONTPRINT"
      Case 1
         Reply$ = "PRINTWARNINGS"
      Case 2
         Reply$ = "PRINTFATAL"
      Case 3
         Reply$ = "PRINTALL"
      Case Else
         Reply$ = "0(DONTPRINT), 1(PRINTWARNINGS), 2(PRINTFATAL), or 3(PRINTALL)"
   End Select
   GetLibraryErrReportingString = Reply$

End Function
Function GetLibraryErrHandlingString(LibraryErrHandling As Long) As String

   Select Case LibraryErrHandling
      Case -1
         Reply$ = "Reset to Original"
      Case 0
         Reply$ = "DONTSTOP"
      Case 1
         Reply$ = "STOPFATAL"
      Case 2
         Reply$ = "STOPALL"
      Case Else
         Reply$ = "0(DONTSTOP), 1(STOPFATAL), or 2(STOPALL)"
   End Select
   GetLibraryErrHandlingString = Reply$
   
End Function
Function GetLocalErrHandlingString(LocalErrHandling As Long) As String

   Select Case LocalErrHandling
      Case -1
         Reply$ = "Reset to Original"
      Case 0
         Reply$ = "None"
      Case 1
         Reply$ = "AbortRoutine"
      Case 2
         Reply$ = "StopOnError"
      Case 3
         Reply$ = "ContinueAfterError"
   End Select
   GetLocalErrHandlingString = Reply$
   
End Function

Function GetPlotTypeString(ByVal PlotType As Long) As String

   Select Case PlotType
      Case 0
         Reply$ = "Volt vs Time"
      Case 1
         Reply$ = "Histogram"
      Case 2
         Reply$ = "GINUMEXPBOARDS"
      Case 6
         Reply$ = "Derivative"
   End Select
   GetPlotTypeString = Reply$

End Function

Function GetConfigGlobalString(ByVal GlobalVal As Long) As String

   Select Case GlobalVal
      Case 0
         Reply$ = "VoltsVsTime"
      Case 1
         Reply$ = "Histogram"
   End Select
   GetConfigGlobalString = Reply$

End Function

Function GetCfgInfoTypeString(ByVal InfoVal As Long) As String

   Select Case InfoVal
      Case GLOBALINFO
         Reply$ = "GLOBALINFO"
      Case BOARDINFO
         Reply$ = "BOARDINFO"
      Case DIGITALINFO
         Reply$ = "DIGITALINFO"
      Case COUNTERINFO
         Reply$ = "COUNTERINFO"
      Case EXPANSIONINFO
         Reply$ = "EXPANSIONINFO"
      Case MISCINFO
         Reply$ = "MISCINFO"
      Case EXPINFOARRAY
         Reply$ = "EXPINFOARRAY"
      Case MEMINFO
         Reply$ = "MEMINFO"
   End Select
   GetCfgInfoTypeString = Reply$

End Function

Function GetCfgItemString(ByVal CfgItemVal As Long) As String

   Select Case CfgItemVal
      Case BICLOCK             '5       /* Clock freq (1, 10 or bus) */
         Reply$ = "BICLOCK"
      Case BIRANGE             '6       /* Switch selectable range */
         Reply$ = "BIRANGE"
      Case BINUMADCHANS        '7       /* Number of A/D channels */
         Reply$ = "BINUMADCHANS"
      Case BICTR0SRC            '104     /* CTR 0 source */
         Reply$ = "BICTR0SRC"
      Case BICTR1SRC            '105     /* CTR 1 source */
         Reply$ = "BICTR1SRC"
      Case BICTR2SRC            '106     /* CTR 2 source */
         Reply$ = "BICTR2SRC"
      Case BIPACERCTR0SRC       '107     /* Pacer CTR 0 source */
         Reply$ = "BIPACERCTR0SRC"
      Case BITRIGEDGE           '113     AD pacer edge
         Reply$ = "BITRIGEDGE"
      Case BIADCFG              '117     AD Config (SE/DIFF) (DevNo)
         Reply$ = "BIADCFG"
      Case BITEMPREJFREQ         '121     'rejection frequency
         Reply$ = "BITEMPREJFREQ"
      Case BICTR3SRC            '130     /* CTR 3 source */
         Reply$ = "BICTR3SRC"
      Case BICTR4SRC            '131     /* CTR 4 source */
         Reply$ = "BICTR4SRC"
      Case BICTR5SRC            '132     /* CTR 5 source */
         Reply$ = "BICTR5SRC"
      Case BIDACTRIG            '148     DAC pacer edge
         Reply$ = "BIDACTRIG"
      Case BITCCHANTYPE         '169
         Reply$ = "BITCCHANTYPE"
      Case BIFWVERSION          '170
         Reply$ = "BIFWVERSION"
      Case BIAIWAVETYPE         '202     /* analog input wave type (for demo board) */
         Reply$ = "BIAIWAVETYPE"
      Case BIADTRIGSRC          '209     /* Analog trigger source */
         Reply$ = "BIADTRIGSRC"
      Case BIBNCSRC             '210     /* BNC source */
         Reply$ = "BIBNCSRC"
      Case BIBNCTHRESHOLD       '211     /* BNC Threshold 2.5V or 0.0V */
         Reply$ = "BIBNCTHRESHOLD"
      Case BISERIALNUM          '214    /* Serial Number for USB boards */
         Reply$ = "BISERIALNUM"
      Case BIDACUPDATEMODE      '215    /* Update immediately or upon AOUPDATE command */
         Reply$ = "BIDACUPDATEMODE"
      Case BIDACUPDATECMD       '216    /* Issue D/A UPDATE command */
         Reply$ = "BIDACUPDATECMD"
      Case BIDACSTARTUP         '217    /* Store last value written for startup */
         Reply$ = "BIDACSTARTUP"
      Case BIADTRIGCOUNT   '219
         Reply$ = "BIADTRIGCOUNT"
      Case BIADFIFOSIZE   '220
         Reply$ = "BIADFIFOSIZE"
      Case BITEMPSENSORTYPE   '235
         Reply$ = "BITEMPSENSORTYPE"
      Case BIADAIMODE      '249
         Reply$ = "BIADAIMODE"
      Case BIADCSETTLETIME       '270
         Reply$ = "BIADCSETTLETIME"
      Case BIDACTRIGCOUNT        '284
         Reply$ = "BIDACTRIGCOUNT"
      Case BIADRES      '291
         Reply$ = "BIADRES"
      Case BIDACRES     '292
         Reply$ = "BIDACRES"
      Case BIDISCONNECT '340
         Reply$ = "BIDISCONNECT"
   End Select
   GetCfgItemString = Reply$

End Function

Function GetStatCondString(Condition As Long) As String

   Select Case Condition
      Case 0
         Reply$ = "StatusIdle"
      Case 1
         Reply$ = "StatusActive"
      Case 2
         Reply$ = "Active Until TotalCount"
      Case -100
         Reply$ = "StatusFixed"
   End Select
   GetStatCondString = Reply$
   
End Function

Function GetGPIBCmdText(ByVal GPIBCmd As String) As String

   Select Case GPIBCmd
      Case "FU0"
         Reply$ = "DC Only"
      Case "FU1"
         Reply$ = "Sine"
      Case "FU2"
         Reply$ = "Square"
      Case "FU3"
         Reply$ = "Triangle"
      Case "FU4"
         Reply$ = "+Ramp"
      Case "FU5"
         Reply$ = "-Ramp"
   End Select
   GetGPIBCmdText = Reply$
   
End Function

Function GetStatLowString(LowVal As Long) As String

   Select Case LowVal
      Case -100
         Reply$ = "StatusFixed"
      Case Else
         Reply$ = "CurCountLow"
   End Select
   GetStatLowString = Reply$
   
End Function

Function GetStatHighString(HighVal As Long) As String

   GetStatHighString = "CurCountHigh or ms between reading CurCount (for StatusFixed)"
   
End Function

Function GetPortFromIndex(PortIndex As Integer) As Long

   PortType = Choose(PortIndex + 1, AUXPORT, FIRSTPORTA, FIRSTPORTB, _
   FIRSTPORTCL, FIRSTPORTCH, SECONDPORTA, SECONDPORTB, SECONDPORTCL, _
   SECONDPORTCH, THIRDPORTA, THIRDPORTB, THIRDPORTCL, THIRDPORTCH, _
   FOURTHPORTA, FOURTHPORTB, FOURTHPORTCL, FOURTHPORTCH, FIFTHPORTA, _
   FIFTHPORTB, FIFTHPORTCL, FIFTHPORTCH, SIXTHPORTA, SIXTHPORTB, _
   SIXTHPORTCL, SIXTHPORTCH, SEVENTHPORTA, SEVENTHPORTB, SEVENTHPORTCL, _
   SEVENTHPORTCH, EIGHTHPORTA, EIGHTHPORTB, EIGHTHPORTCL, EIGHTHPORTCH)
   
   If Not IsNull(PortType) Then
      TypeOfPort& = Val(PortType)
      GetPortFromIndex = TypeOfPort&
   Else
      GetPortFromIndex = 0
   End If

End Function

Function GetErrorText(ErrorValue As Long) As String

   On Error GoTo AltErrPath
   
   Open TryPath$ & "cbercode.txt" For Input As #4

   Do While Not EOF(4)
      Line Input #4, A1
      Loca = InStr(1, A1, " ")
      If Loca > 0 Then
         ErrCode$ = Left(A1, Loca - 1)
         If Val(ErrCode$) = ErrorValue Then
            Loca = InStr(1, A1, "{")
            If Loca > 0 Then
               ErrString$ = Mid(A1, Loca + 1)
               ErrString$ = Left(ErrString$, Len(ErrString$) - 1)
               Exit Do
            End If
         End If
      End If
   Loop
   Close #4
   GetErrorText = ErrString$
   Exit Function
   
AltErrPath:
   If Err = 53 Then
      Select Case Attempts%
         Case 0
            TryPath$ = App.Path & "\"
         Case 1
            'check registry for UL path
            CurGroupKey$ = "SOFTWARE\Universal Library"
            KeyName$ = "RootDir"
            ProgExists% = GetRegGroup(HKEY_LOCAL_MACHINE, CurGroupKey$, hProgResult&)
            YN% = GetKeyValue(hProgResult&, KeyName$, KeyVal$)
            If YN% Then
               TryPath$ = KeyVal$
            End If
         Case Else
            If mnNoErrorFile Then Exit Function
            MsgBox "Could not find the error file 'cbercode.txt'. " & _
            "Install the Universal Library or copy this file to '" & App.Path & " '.", _
            vbOKOnly, "Universal Test File Error"
            mnNoErrorFile = True
            Exit Function
      End Select
      Attempts% = Attempts% + 1
      Resume 0
   Else
      MsgBox Error(Err), vbOKOnly, "Universal Test File Error"
      Exit Function
   End If

End Function

Public Function GetBoardFile() As String

   On Error GoTo AltBoardListPath
   
   Open TryPath$ & "BoardList.txt" For Input As #4
   Close #4
   
   GetBoardFile = TryPath$ & "BoardList.txt"
   Exit Function
   
AltBoardListPath:
   If Err = 53 Then
      Select Case Attempts%
         Case 0
            TryPath$ = App.Path & "\"
         Case Else
            If mnNoBoardFile Then Exit Function
            MsgBox "Could not find the procuct list file 'BoardList.txt'. " & _
            "Copy this file to '" & App.Path & " '.", _
            vbOKOnly, "Universal Test File Error"
            mnNoBoardFile = True
            Exit Function
      End Select
      Attempts% = Attempts% + 1
      Resume 0
   Else
      MsgBox Error(Err), vbOKOnly, "Universal Test File Error"
      Exit Function
   End If

End Function

Function GetShortPort(ByVal PortNum As Long) As String

   Select Case PortNum
      Case AUXPORT
         Reply$ = "AUX"
      Case FIRSTPORTA
         Reply$ = "1A"
      Case FIRSTPORTB
         Reply$ = "1B"
      Case FIRSTPORTCL
         Reply$ = "1CL"
      Case FIRSTPORTC
         Reply$ = "1C"
      Case FIRSTPORTCH
         Reply$ = "1CH"
      Case SECONDPORTA
         Reply$ = "2A"
      Case SECONDPORTB
         Reply$ = "2B"
      Case SECONDPORTCL
         Reply$ = "2CL"
      Case SECONDPORTCH
         Reply$ = "2CH"
      Case THIRDPORTA
         Reply$ = "3A"
      Case THIRDPORTB
         Reply$ = "3B"
      Case THIRDPORTCL
         Reply$ = "3CL"
      Case THIRDPORTCH
         Reply$ = "3CH"
      Case FOURTHPORTA
         Reply$ = "4A"
      Case FOURTHPORTB
         Reply$ = "4B"
      Case FOURTHPORTCL
         Reply$ = "4CL"
      Case FOURTHPORTCH
         Reply$ = "4CH"
      Case FIFTHPORTA
         Reply$ = "5A"
      Case FIFTHPORTB
         Reply$ = "5B"
      Case FIFTHPORTCL
         Reply$ = "5CL"
      Case FIFTHPORTCH
         Reply$ = "5CH"
      Case SIXTHPORTA
         Reply$ = "6A"
      Case SIXTHPORTB
         Reply$ = "6B"
      Case SIXTHPORTCL
         Reply$ = "6CL"
      Case SIXTHPORTCH
         Reply$ = "6CH"
      Case SEVENTHPORTA
         Reply$ = "7A"
      Case SEVENTHPORTB
         Reply$ = "7B"
      Case SEVENTHPORTCL
         Reply$ = "7CL"
      Case SEVENTHPORTCH
         Reply$ = "7CH"
      Case EIGHTHPORTA
         Reply$ = "8A"
      Case EIGHTHPORTB
         Reply$ = "8B"
      Case EIGHTHPORTCL
         Reply$ = "8CL"
      Case EIGHTHPORTCH
         Reply$ = "8CH"
      Case Else
         Reply$ = "BADPORT"
   End Select
   GetShortPort = Reply$
   
End Function

Function GetParamString(FunctionID As Long, ParamNum As Long, ParamVal As Long) As String

   Select Case FunctionID
      Case AInScan   '2
         Select Case ParamNum
            Case 5
               Reply$ = GetRangeString(ParamVal)
            Case 7
               Reply$ = GetOptionsString(ParamVal, ANALOG_IN)
         End Select
      Case DConfigPort To DOutScan    '22 - 28
         Reply$ = GetPortString(ParamVal)
      Case SetTrigger   '49
         Reply$ = GetTrigTypeString(ParamVal)
      Case EnableEvent  '71
         Reply$ = GetEventTypeString(ParamVal)
      Case Else
         Reply$ = "UNDEFINED (Ask Carl)"
   End Select
   GetParamString = Reply$

End Function

Public Function GetCtrTypeString(ByVal CtrType As Integer) As String

   Select Case CtrType
      Case 1
         CtrName$ = "82C54"
      Case 2
         CtrName$ = "9513"
      Case 3
         CtrName$ = "8536"
      Case 4
         CtrName$ = "7266"
      Case 5
         CtrName$ = "Event Counter"
      Case 6
         CtrName$ = "Scan Counter"
      Case 7
         CtrName$ = "Timer Out"
      Case 8
         CtrName$ = "Quad Counter"
      Case 9
         CtrName$ = "Pulse Out"
   End Select
   GetCtrTypeString = CtrName$
   
End Function

Public Function GetTcTypeString(ByVal TcType As Integer) As String

   Select Case TcType
      Case 0
         CurType$ = "None"
      Case TC_TYPE_J
         CurType$ = "J"
      Case TC_TYPE_K
         CurType$ = "K"
      Case TC_TYPE_T
         CurType$ = "T"
      Case TC_TYPE_E
         CurType$ = "E"
      Case TC_TYPE_R
         CurType$ = "R"
      Case TC_TYPE_S
         CurType$ = "S"
      Case TC_TYPE_B
         CurType$ = "B"
      Case TC_TYPE_N
         CurType$ = "N"
      Case Else
         CurType$ = "Undefined"
   End Select
   GetTcTypeString = CurType$
   
End Function

Public Function GetAiChanTypeString(ByVal ChanType As Integer) As String

   Select Case ChanType
      Case -1
         CurType$ = "Not Configurable"
      Case AI_CHAN_TYPE_VOLTAGE
         CurType$ = "Voltage"
      Case AI_CHAN_TYPE_CURRENT
         CurType$ = "Current"
      Case AI_CHAN_TYPE_RESISTANCE_10K4W
         CurType$ = "R 10k 4-wire"
      Case AI_CHAN_TYPE_RESISTANCE_1K4W
         CurType$ = "R 1k 4-wire"
      Case AI_CHAN_TYPE_RESISTANCE_10K2W
         CurType$ = "R 10k 2-wire"
      Case AI_CHAN_TYPE_RESISTANCE_1K2W
         CurType$ = "R 1k 2-wire"
      Case AI_CHAN_TYPE_TC
         CurType$ = "Thermocouple"
      Case AI_CHAN_TYPE_RTD_1000OHM_4W
         CurType$ = "RTD 1k 4-wire"
      Case AI_CHAN_TYPE_RTD_100OHM_4W
         CurType$ = "RTD 100 4-wire"
      Case AI_CHAN_TYPE_RTD_1000OHM_3W
         CurType$ = "RTD 1k 3-wire"
      Case AI_CHAN_TYPE_RTD_100OHM_3W
         CurType$ = "RTD 100 3-wire"
      Case AI_CHAN_TYPE_QUART_BRIDGE_350OHM
         CurType$ = "QBridge 350ohm"
      Case AI_CHAN_TYPE_QUART_BRIDGE_120OHM
         CurType$ = "QBridge 120ohm"
      Case AI_CHAN_TYPE_HALF_BRIDGE
         CurType$ = "HBridge"
      Case AI_CHAN_TYPE_FULL_BRIDGE_62PT5mVV
         CurType$ = "FBridge 62.5mV/V"
      Case AI_CHAN_TYPE_FULL_BRIDGE_7PT8mVV
         CurType$ = "FBridge 7.8mV/V"
      Case Else
         CurType$ = "Undefined"
   End Select
   GetAiChanTypeString = CurType$

End Function

Public Function GetAiChanModeString(ByVal ChanMode As Integer) As String

   Select Case ChanMode
      Case -1
         CurMode$ = "not configurable"
      Case DIFFERENTIAL
         CurMode$ = "Differential"
      Case SINGLE_ENDED
         CurMode$ = "Single-ended"
      Case GROUNDED
         CurMode$ = "Grounded"
      Case Else
         CurMode$ = "Undefined"
   End Select
   GetAiChanModeString = CurMode$

End Function

