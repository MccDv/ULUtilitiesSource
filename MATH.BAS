Attribute VB_Name = "MATH"
'Description
'  math functions and data conversion functions commonly used
'  in data acquisition

'Public interfaces:
'  BinIntData (Chan As Integer, DatArray() As Integer, FirstBin As Integer,
'                 LargestBin As Integer, HistBins() As Integer) As Integer
'     pass a two dimensional integer array, the first dimension
'     containing the number of channels (zero based),
'     the second dimension containing data. Also pass the
'     channel number to bin (zero based) and a pointer to
'     the HistBins array (which will be redimensioned by
'     this function as required)
'     returns the number of bins (range is between 0 and
'     4095 for 12 bit data) containing data for
'     Chan (which will be the new size of HistBins())
'     sets FirstBin to the code of the first bin (between
'     0 and 4095 for 12 bit res), sets LargestBin to the
'     value of the bin containing the most data (maximum
'     value is the size of the second dimension of DatArray)
'  BinLongData (Chan As Integer, DatArray() As Long, FirstBin As Long,
'                 LargestBin As Integer, HistBins() As Integer) As Long
'     same as BinIntData() but for long values...returns the
'     number of bins (range is between 0 and 65535 for 16 bit
'     data) but will only redimension and fill DatArray() if
'     the number of bins is less than an arbitrary value set
'     internally to this function (1000 originally)
'     Use the IntArrayToULong() function to convert 16 bit integer
'     data values to longs before passing data to this function
'  IntArrayToULong (IntArray() As Integer, RealArray() As Long) As Integer
'     converts integer array of 16 bit data (-32768 to 32767) to
'     a normalized array of type long (0 to 65535) for analysis
'  IntValToULong (IntVal As Integer) As Long
'     same as IntArrayToULong but for single value
'  ULongValToInt (LongVal As Long) As Integer
'     same as IntArrayToULong but for single value
'  StdDev (x() As Integer) As Single
'     calculates Standard Deviation of data passed in x()
'     used mainly to analyze histogram data...pass HistBins()
'  GetCounts (Resolution As Integer, Range As Integer,
'                                 Voltage As Single) As Long
'     pass in Resolution in number-of-bits (i.e., 12 or 16)
'     returns counts in unsigned integer format
'  GetVolts (Resolution As Integer, Range As Integer,
'                                 Counts As Integer) As Single
'     pass in Resolution in number-of-bits (i.e., 12 or 16)
'     pass in counts as integer
'     returns volts based on resolution and range

'Dependencies
'     locinfo.bas

'____________________________________________________

Global Const V_EMPTY = 0
Global Const V_NULL = 1
Global Const V_INTEGER = 2
Global Const V_LONG = 3
Global Const V_SINGLE = 4
Global Const V_DOUBLE = 5
Global Const V_CURRENCY = 6
Global Const V_DATE = 7
Global Const V_STRING = 8

Dim msBoardName As String

Function BinData(Chan As Integer, DatArray As Variant, FirstBin As Variant, LargestBin As Variant, LargestBinCode As Variant, HistBins() As Long, Divisor As Integer, Average As Variant, Optional NumberOfPoints As Variant) As Long

   Dim CurValue As Variant
   Multiplier& = 1
   'If (VarType(DatArray(0, 0)) > 3) Then
   '   Multiplier& = 1000
   'End If
   FirstPoint& = FirstBin * Multiplier&
   MinBin = 2 ^ 48 '65535.5
   MaxBin = -65535.5
   LargestBin = 0

   If IsMissing(NumberOfPoints) Then
      ArraySize = UBound(DatArray, 2) ' - 1
   Else
      ArraySize = NumberOfPoints '- 1
      If ArraySize > UBound(DatArray, 2) Then ArraySize = UBound(DatArray, 2)
   End If
   ArrayElements = -1
   For i& = FirstPoint& To (FirstPoint& + ArraySize)
      CurValue = DatArray(Chan, i&) * Multiplier&
      If Not (IsEmpty(CurValue)) Then
         If CurValue < MinBin Then
            MinBin = CurValue
         End If
         If CurValue > MaxBin Then
            MaxBin = CurValue
         End If
         ArrayElements = ArrayElements + 1
         TotalCounts = TotalCounts + CurValue
      End If
   Next i&
   Average = TotalCounts / (ArraySize + 1)
   If ArrayElements < 0 Then
      BinData = 0
      Exit Function
   End If

   
   FirstBin = MinBin
   BinCount = (MaxBin - MinBin) + 1
   If Not (BinCount < 2147483648#) Then
      NumBins& = 2147483647
      'If mnResolution = 32 Then mlDivisor = &H10000
      'If mnResolution = 24 Then mlDivisor = &H100
   Else
      NumBins& = (MaxBin - MinBin) + 1
   End If
   If (VarType(DatArray(0, 0)) > 3) And NumBins& > 32000 Then
      'removed this condition due to issue with 2616 [If 0 Then ']
      'put it back because it causes scaling issues with cal
      LargestBin = MaxBin
      BinData = NumBins&
      Divisor = 1
   Else
      TotalBins& = NumBins&
      Divisor = 1
      If (Not (NumBins& < 32767)) And (Not WIN32APP) Then
         'vb3 allows no more than 32767 elements in an array
         '16 bit boards have a potential of 65536 bins
         TotalBins& = 32766
         Divisor = 4
      End If
      If TotalBins& < 32768 Then
         ReDim HistBins(TotalBins&)
         CalculateBins% = True
      Else
         TotalBins& = 32766
         'ReDim HistBins(TotalBins&)
         Divisor = 4
      End If
      If CalculateBins% Then
         For i& = FirstPoint& To FirstPoint& + ArrayElements
            CurValue = DatArray(Chan, i&) * Multiplier&
            Element& = (CurValue - MinBin) / Divisor
            HistBins(Element&) = HistBins(Element&) + 1
         Next i&
         For i& = 0 To TotalBins&
            If HistBins(i&) > LargestBin Then
               LargestBin = HistBins(i&)
               LargestBinCode = MinBin + i&
            End If
         Next i&
      End If
      BinData = NumBins&
   End If
   If (VarType(DatArray(0, 0)) > 3) Then LargestBin = MaxBin

End Function

Sub InitMathModule(BoardName As String)

   mnBoardName = BoardName

End Sub

Function IntArrayToULong(IntArray() As Integer, RealArray() As Variant, NumPoints&, StartPoint&) As Long

   IntArrayToULong = 0
   Chans% = UBound(IntArray)

   For Ch% = 0 To Chans%
      For l& = StartPoint& To StartPoint& + NumPoints&
         RealArray(Ch%, l& - StartPoint&) = (IntArray(Ch%, l&) Xor &H8000) + 32768
      Next l&
   Next Ch%
   IntArrayToULong = NumPoints& * (Chans% + 1)

End Function

Function IntArrayToLong(IntArray() As Integer, LongArray() As Long, ByVal NumPoints&, ByVal StartPoint&) As Long

   IntArrayToLong = 0
   Chans% = UBound(IntArray)

   For Ch% = 0 To Chans%
      For l& = StartPoint& To StartPoint& + NumPoints&
         LongArray(Ch%, l& - StartPoint&) = (IntArray(Ch%, l&) Xor &H8000) + 32768
      Next l&
   Next Ch%
   IntArrayToLong = NumPoints& * (Chans% + 1)

End Function

Function NetIntArrayToULong(NetArray() As Integer, RealArray() As Variant, NumPoints&, StartPoint&, Chans%) As Long

   NetIntArrayToULong = 0

   For l& = StartPoint& To StartPoint& + NumPoints& 'Step Chans% + 1
      For Ch% = 0 To Chans%
         RealArray(Ch%, l& - StartPoint&) = (NetArray(Element&) Xor &H8000) + 32768
         Element& = Element& + 1
      Next Ch%
   Next l&
   NetIntArrayToULong = NumPoints& * (Chans% + 1)

End Function

Function IntArrayToInt(IntArray() As Integer, RealArray() As Variant, NumPoints&, StartPoint&) As Long

   IntArrayToInt = 0
   Chans% = UBound(IntArray)

   For Ch% = 0 To Chans%
      For l& = StartPoint& To StartPoint& + NumPoints&
         RealArray(Ch%, l& - StartPoint&) = IntArray(Ch%, l&)
      Next l&
   Next Ch%
   IntArrayToInt = NumPoints& * (Chans% + 1)

End Function

Function IntValToULong(IntVal As Integer) As Long

   IntValToULong = (IntVal Xor &H8000) + 32768

End Function

Function StdDev(x() As Integer) As Single

   Total% = UBound(x)
   ReDim y(9) As Integer

   For i% = 0 To Total%
      tot! = tot! + x(i%)
   Next i%
   mean! = tot! / (Total% + 1)

   For i% = 0 To Total%
      y(i%) = (x(i%) - mean!) ^ 2
   Next i%
   
   tot! = 0
   For i% = 0 To Total%
      tot! = tot! + y(i%)
   Next i%
   mean! = tot! / (Total% + 1)
   StdDev! = Sqr(mean!)

End Function

Function ULongValToInt(LongVal As Long) As Integer

   Select Case LongVal
      Case Is > 65535
         ULongValToInt = -1
      Case Is < 0
         ULongValToInt = 0
      Case Else
         ULongValToInt = (LongVal - 32768) Xor &H8000
   End Select

End Function

Function UDWordValToLong(DWordVal As Double) As Long

   Select Case DWordVal
      Case Is > 4294967295#
         UDWordValToLong = -1
      Case Is > 2147483647
         UDWordValToLong = DWordVal - 4294967296#
      Case Is < 0
         UDWordValToLong = 0
      Case Else
         UDWordValToLong = DWordVal
   End Select

End Function

Function LongToUDWordVal(LongVal As Long) As Double

   Select Case LongVal
      Case Is < 0
         LongToUDWordVal = 4294967296# + LongVal
      Case Else
         LongToUDWordVal = LongVal
   End Select

End Function

Function GetBitsFromArray(FullDataSet As Variant, _
DataType As Long, NumPoints As Long, NumBits As Integer, _
Bits As Variant, Optional Chan As Integer) As Integer

   'returns array of binary values, one per bit, for each value in array
   NumSamples& = NumPoints * NumBits
   If NumSamples& < 1 Then Exit Function
   
   ReDim LngBits(NumSamples& - 1) As Long
   'LngBits should probably be integer but defined as long in header

   If DataType = vbLong Then
      BitRes% = 32
   Else
      BitRes% = 16
   End If
   
   If NumBits > BitRes% Then
      Residual% = NumBits Mod (BitRes%)
      BitsInRes% = ((NumBits + Residual%) - NumBits) \ BitRes%
      BitsToStuff% = BitRes%
   Else
      BitsToStuff% = NumBits
   End If
   
   For Samp& = 0 To NumPoints - 1
      CurVal& = FullDataSet(Chan, Samp&)
      For ResBit% = 0 To Residual% - 1
         'repeat read of sample to fill all bits
         BitValue& = Abs((CurVal& And (2 ^ ResBit%)) = (2 ^ ResBit%))
         LngBits(BitNum&) = BitValue&
         BitNum& = BitNum& + 1
      Next ResBit%
      For ResDupe% = 0 To BitsInRes% '- 1
         For BitLoc% = 0 To BitsToStuff% - 1
            vCurVal = CurVal&: vLoc = BitLoc%
            vResult = Abs((vCurVal And (2 ^ vLoc)) = (2 ^ vLoc))
            BitValue& = vResult
            LngBits(BitNum&) = BitValue&
            BitNum& = BitNum& + 1
         Next BitLoc%
      Next ResDupe%
   Next Samp&
   GetBitsFromArray = True
   Bits = LngBits()
   
End Function

