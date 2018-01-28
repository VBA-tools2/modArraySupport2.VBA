Attribute VB_Name = "mod_Test_ArraySupport"

Option Explicit
Option Compare Text
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.PermissiveAssertClass
Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Public Sub ModuleInitialize()
   'this method runs once per module.
   Set Assert = New Rubberduck.PermissiveAssertClass
   Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
   'this method runs once per module.
   Set Assert = Nothing
   Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
   'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
   'this method runs after every test in the module.
End Sub


'==============================================================================
'unit tests for 'AreDataTypesCompatible'
'==============================================================================

'@TestMethod
Public Sub AreDataTypesCompatible_ScalarSourceArrayDest_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Source As LongPtr
   Dim Dest() As LongPtr
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AreDataTypesCompatible_BothStringScalars_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Source As String
   Dim Dest As String
   
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AreDataTypesCompatible_BothStringArrays_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Source() As String
   Dim Dest() As String
   
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AreDataTypesCompatible_LongSourceIntegerDest_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Source As LongPtr
   Dim Dest As Integer
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AreDataTypesCompatible_IntegerSourceLongDest_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Source As Integer
   Dim Dest As LongPtr
   
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AreDataTypesCompatible_DoubleSourceLongDest_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Source As Double
   Dim Dest As LongPtr
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AreDataTypesCompatible_BothObjects_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Source As Object
   Dim Dest As Object
   
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AreDataTypesCompatible_SingleSourceDateDest_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Source As Single
   Dim Dest As Date
   
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.AreDataTypesCompatible(Source, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''2do: How to do this test?
''     --> in 'ChangeBoundsOfArray_VariantArr_ReturnsTrueAndChangedArr' are
''         'Empty' entries added at the end of the array
''@TestMethod
'Public Sub AreDataTypesCompatible_VariantSourceEmptyDest_ReturnsTrue()
'   On Error GoTo TestFail
'
'   'Arrange:
'   Dim Source(0) As Variant
'   Dim Dest(0) As Variant
'   Dim vDummy As Variant
'
'
'   'Act:
'   vDummy = 4534
'   Source(0) = CVar(vDummy)
'   Dest(0) = Empty
'
'   'Assert:
'   Assert.IsTrue modArraySupport.AreDataTypesCompatible(Source(0), Dest(0))
'
'TestExit:
'   Exit Sub
'TestFail:
'   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


'==============================================================================
'unit tests for 'CompareArrays'
'==============================================================================

'@TestMethod
Public Sub CompareArrays_UnallocatedArrays_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1() As String
   Dim Arr2() As String
   Dim ResArr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CompareArrays(Arr1, Arr2, ResArr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CompareArrays_LegalAndTextCompare_ReturnsTrueAndResArr()
   On Error GoTo TestFail

   Dim Arr1(1 To 5) As String
   Dim Arr2(1 To 5) As String
   Dim ResArr() As LongPtr

   '===========================================================================
   Dim aExpected(1 To 5) As LongPtr
      aExpected(1) = -1
      aExpected(2) = 1
      aExpected(3) = -1
      aExpected(4) = 0
      aExpected(5) = 0
   '===========================================================================


   'Arrange:
   Arr1(1) = "2"
   Arr1(2) = "c"
   Arr1(3) = vbNullString
   Arr1(4) = "."
   Arr1(5) = "B"

   Arr2(1) = "4"
   Arr2(2) = "a"
   Arr2(3) = "x"
   Arr2(4) = "."
   Arr2(5) = "b"

   'Act:
   If Not modArraySupport.CompareArrays(Arr1, Arr2, ResArr, vbTextCompare) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CompareArrays_LegalAndBinaryCompare_ReturnsTrueAndResArr()
   On Error GoTo TestFail

   Dim Arr1(1 To 5) As String
   Dim Arr2(1 To 5) As String
   Dim ResArr() As LongPtr

   '===========================================================================
   Dim aExpected(1 To 5) As LongPtr
      aExpected(1) = -1
      aExpected(2) = 1
      aExpected(3) = -1
      aExpected(4) = 0
      aExpected(5) = -1
   '===========================================================================


   'Arrange:
   Arr1(1) = "2"
   Arr1(2) = "c"
   Arr1(3) = vbNullString
   Arr1(4) = "."
   Arr1(5) = "B"

   Arr2(1) = "4"
   Arr2(2) = "a"
   Arr2(3) = "x"
   Arr2(4) = "."
   Arr2(5) = "b"

   'Act:
   If Not modArraySupport.CompareArrays(Arr1, Arr2, ResArr, vbBinaryCompare) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'ConcatenateArrays'
'==============================================================================

'@TestMethod
Public Sub ConcatenateArrays_StaticResultArray_ResultsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArray(1) As LongPtr       'MUST be dynamic
   Dim ArrayToAppend(1) As LongPtr


   ResultArray(1) = 8
   ArrayToAppend(1) = 111

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_BothArraysUnallocated_ResultsUnallocatedArray()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArray() As LongPtr        'MUST be dynamic
   Dim ArrayToAppend() As LongPtr


   'Act:
   If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
         GoTo TestFail

   'Assert:
   Assert.IsFalse IsArrayAllocated(ResultArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_ArrayToAppendUnallocated_ResultsUnchangedResultArray()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArray() As LongPtr        'MUST be dynamic
   Dim ArrayToAppend() As LongPtr

   '===========================================================================
   Dim aExpected(1 To 2) As LongPtr
      aExpected(1) = 8
      aExpected(2) = 9
   '===========================================================================


   ReDim ResultArray(1 To 2)
   ResultArray(1) = 8
   ResultArray(2) = 9

   'Act:
   If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_LegalLong_ResultsTrueAndResultArray()
   On Error GoTo TestFail

   Dim ResultArray() As LongPtr        'MUST be dynamic
   Dim ArrayToAppend(1 To 3) As Integer

   '===========================================================================
   Dim aExpected(1 To 6) As LongPtr
      aExpected(1) = 8
      aExpected(2) = 9
      aExpected(3) = 10
      aExpected(4) = 111
      aExpected(5) = 112
      aExpected(6) = 113
   '===========================================================================


   'Arrange:
   ReDim ResultArray(1 To 3)
   ResultArray(1) = 8
   ResultArray(2) = 9
   ResultArray(3) = 10

   ArrayToAppend(1) = 111
   ArrayToAppend(2) = 112
   ArrayToAppend(3) = 113

   'Act:
   If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''2do: add a test that involves objects
''     (have a look at <https://stackoverflow.com/a/11254505>
''@TestMethod
'Public Sub ConcatenateArrays_LegalVariant_ResultsTrueAndResultArray()
'   On Error GoTo TestFail
'
'   Dim ResultArray() As Range          'MUST be dynamic
'   Dim ArrayToAppend(0 To 0) As Range
'   Dim i As LongPtr
'
'   '===========================================================================
'   Dim wks As Worksheet
'   Set wks = tblFunctions
'   Dim aExpected(1 To 2) As Range
'   With wks
'      Set aExpected(1) = .Cells(1, 1)
'      Set aExpected(2) = .Cells(1, 2)
'   End With
'   '===========================================================================
'
'
'   'Arrange:
'   With wks
'      ReDim ResultArray(1 To 1)
'      Set ResultArray(1) = .Cells(1, 1)
'      Set ArrayToAppend(0) = .Cells(1, 2)
'   End With
'
'   'Act:
'   If Not modArraySupport.ConcatenateArrays(ResultArray, ArrayToAppend) Then _
'         GoTo TestFail
'
'   'Assert:
'   For i = LBound(ResultArray) To UBound(ResultArray)
'Debug.Print aExpected(i) Is ResultArray(i)
'      Assert.AreSame aExpected(i), ResultArray(i)
'   Next
'
''   If B = True Then
''      If modArraySupport.IsArrayAllocated(ResultArray) = True Then
''         For i = LBound(ResultArray) To UBound(ResultArray)
''            If IsObject(ResultArray(i)) = True Then
''Debug.Print CStr(i), "is object", TypeName(ResultArray(i))
''            Else
''Debug.Print CStr(i), ResultArray(i)
''            End If
''         Next
''      Else
''Debug.Print "Result Array Is Not Allocated."
''      End If
''   Else
''Debug.Print "ConcatenateArrays returned False"
''   End If
'
'TestExit:
'   Exit Sub
'TestFail:
'   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


'==============================================================================
'unit tests for 'CopyArray'
'==============================================================================

'@TestMethod
Public Sub CopyArray_UnallocatedSrc_ResultsTrueAndUnchangedDest()
   On Error GoTo TestFail

   Dim Src() As LongPtr
   Dim Dest(0) As Integer

   '===========================================================================
   Dim aExpected(0) As Integer
      aExpected(0) = 50
   '===========================================================================


   'Arrange:
   Dest(0) = 50

   'Act:
   If Not modArraySupport.CopyArray(Src, Dest) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Dest

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_IncompatibleDest_ResultsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Src(1 To 2) As LongPtr
   Dim Dest(1 To 2) As Integer


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyArray(Src, Dest)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_AllocatedDestLessElementsThenSrc_ResultsTrueAndDestArray()
   On Error GoTo TestFail

   Dim Src(1 To 3) As LongPtr
   Dim Dest(10 To 11) As LongPtr

   '===========================================================================
   Dim aExpected(10 To 11) As LongPtr
      aExpected(10) = 1
      aExpected(11) = 2
   '===========================================================================


   'Arrange:
   Src(1) = 1
   Src(2) = 2
   Src(3) = 3

   'Act:
   If Not modArraySupport.CopyArray(Src, Dest) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Dest

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_AllocatedDestMoreElementsThenSrc_ResultsTrueAndDestArray()
   On Error GoTo TestFail

   Dim Src(1 To 3) As LongPtr
   Dim Dest(10 To 13) As LongPtr

   '===========================================================================
   Dim aExpected(10 To 13) As LongPtr
      aExpected(10) = 1
      aExpected(11) = 2
      aExpected(12) = 3
      aExpected(13) = 0
   '===========================================================================


   'Arrange:
   Src(1) = 1
   Src(2) = 2
   Src(3) = 3

   'Act:
   If Not modArraySupport.CopyArray(Src, Dest) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Dest

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyArray_NoCompatibilityCheck_ResultsTrueAndDestArrayWithOverflow()
   On Error GoTo TestFail

   Dim Src(1 To 2) As LongPtr
   Dim Dest(1 To 2) As Integer

   '===========================================================================
   Dim aExpected(1 To 2) As Integer
      aExpected(1) = 1234
      aExpected(2) = 0
   '===========================================================================


   'Arrange:
   Src(1) = 1234
   Src(2) = Rows.Count * 10

   'Act:
   If Not modArraySupport.CopyArray(Src, Dest, True) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Dest

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'2do: Add tests with Objects


Public Sub DemoCopyArraySubSetToArray()

   Dim InputArray(1 To 10) As LongPtr
   Dim ResultArray() As LongPtr

   Dim StartNdx As LongPtr
   Dim EndNdx As LongPtr
   Dim DestNdx As LongPtr
   Dim B As Boolean
   Dim N As LongPtr


   For N = LBound(InputArray) To UBound(InputArray)
      InputArray(N) = N * 10
   Next N

   ReDim ResultArray(1 To 10)
   For N = LBound(ResultArray) To UBound(ResultArray)
      ResultArray(N) = -N
   Next N

   StartNdx = 1
   EndNdx = 5
   DestNdx = 3

   B = modArraySupport.CopyArraySubSetToArray( _
         InputArray, ResultArray, StartNdx, EndNdx, DestNdx)

   If B = True Then
      If modArraySupport.IsArrayAllocated(ResultArray) = True Then
         For N = LBound(ResultArray) To UBound(ResultArray)
            If IsObject(ResultArray(N)) = True Then
Debug.Print CStr(N), "is object"
            Else
Debug.Print CStr(N), ResultArray(N)
            End If
         Next N
      Else
Debug.Print "ResultArray is not allocated"
      End If
   Else
Debug.Print "CopyArraySubSetToArray returned False"
   End If

End Sub


Public Sub DemoCopyNonNothingObjectsToArray()

   Dim SourceArray(1 To 5) As Object
   Dim ResultArray() As Object
   Dim B As Boolean
   Dim N As LongPtr


   Set SourceArray(1) = Range("A1")
   Set SourceArray(2) = Range("A2")
   Set SourceArray(3) = Nothing
   Set SourceArray(4) = Nothing
   Set SourceArray(5) = Range("A5")

   B = modArraySupport.CopyNonNothingObjectsToArray(SourceArray, ResultArray, False)

   If B = True Then
      For N = LBound(ResultArray) To UBound(ResultArray)
Debug.Print CStr(N), ResultArray(N).Address
      Next N
   Else
Debug.Print "CopyNonNothingObjectsToArray returned False"
   End If

End Sub


'==============================================================================
'unit tests for 'DataTypeOfArray'
'==============================================================================

'@TestMethod
Public Sub DataTypeOfArray_NoArray_ReturnsMinusOne()
   On Error GoTo TestFail

   'Arrange:
   Dim sTest As String
   Dim aActual As VbVarType

   '===========================================================================
   Const aExpected As LongPtr = -1
   '===========================================================================


   'Act:
   aActual = modArraySupport.DataTypeOfArray(sTest)

   'Assert:
   Assert.AreEqual aExpected, CLngPtr(aActual)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DataTypeOfArray_UnallocatedArray_ReturnsVbDouble()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Double
   Dim aActual As VbVarType

   '===========================================================================
   Const aExpected As LongPtr = vbDouble
   '===========================================================================


   'Act:
   aActual = modArraySupport.DataTypeOfArray(Arr)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DataTypeOfArray_Test1DStringArray_ReturnsVbString()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 4) As String
   Dim aActual As VbVarType

   '===========================================================================
   Const aExpected As LongPtr = vbString
   '===========================================================================


   'Act:
   aActual = modArraySupport.DataTypeOfArray(Arr)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DataTypeOfArray_Test2DStringArray_ReturnsVbString()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 4, 5 To 6) As String
   Dim aActual As VbVarType

   '===========================================================================
   Const aExpected As LongPtr = vbString
   '===========================================================================


   'Act:
   aActual = modArraySupport.DataTypeOfArray(Arr)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DataTypeOfArray_Test3DStringArray_ReturnsMinusThree()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 4, 5 To 6, 8 To 8) As String
   Dim aActual As VbVarType

   '===========================================================================
   Const aExpected As LongPtr = vbString
   '===========================================================================


   'Act:
   aActual = modArraySupport.DataTypeOfArray(Arr)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'2do: Add tests with Objects


Public Sub DemoDeleteArrayElement()

   Dim Stat(1 To 3) As LongPtr
   Dim Dyn() As Variant
   Dim N As LongPtr
   Dim B As Boolean


   ReDim Dyn(1 To 3)
   Stat(1) = 1
   Stat(2) = 2
   Stat(3) = 3
   Dyn(1) = "abc"
   Dyn(2) = 1234
   Dyn(3) = "ABC"

   B = modArraySupport.DeleteArrayElement(Stat, 1, False)

   If B = True Then
      For N = LBound(Stat) To UBound(Stat)
Debug.Print CStr(N), Stat(N)
      Next N
   Else
Debug.Print "DeleteArrayElement returned false"
   End If


   B = modArraySupport.DeleteArrayElement(Dyn, 2, False)

   If B = True Then
      For N = LBound(Dyn) To UBound(Dyn)
Debug.Print CStr(N), Dyn(N)
      Next N
   Else
Debug.Print "DeleteArrayElement returned false"
   End If

End Sub


Public Sub DemoFirstNonEmptyStringIndexInArray()

   Dim A(1 To 4) As String
   Dim R As LongPtr


   A(1) = vbNullString
   A(2) = vbNullString
   A(3) = "A"
   A(4) = "B"

   R = modArraySupport.FirstNonEmptyStringIndexInArray(A)
Debug.Print "FirstNonEmptyStringIndexInArray", CStr(R)

End Sub


Public Sub DemoInsertElementIntoArray()

   Dim Arr() As LongPtr
   Dim N As LongPtr
   Dim B As Boolean


   ReDim Arr(1 To 10)
   For N = LBound(Arr) To UBound(Arr)
      Arr(N) = N * 10
   Next N

   B = modArraySupport.InsertElementIntoArray(Arr, 5, 12345)

   If B = True Then
      For N = LBound(Arr) To UBound(Arr)
Debug.Print CStr(N), Arr(N)
      Next N
   Else
Debug.Print "InsertElementIntoArray returned false."
   End If

End Sub


Public Sub DemoIsArrayAllDefault()

   Dim L(1 To 4) As LongPtr
   Dim Obj(1 To 4) As Object
   Dim B As Boolean


   B = modArraySupport.IsArrayAllDefault(L)
Debug.Print "IsArrayAllDefault L", B

   B = modArraySupport.IsArrayAllDefault(Obj)
Debug.Print "IsArrayAllDefault Obj", B

   Set Obj(1) = Range("A1")
   B = modArraySupport.IsArrayAllDefault(Obj)
Debug.Print "IsArrayAllDefault Obj", B

End Sub


'==============================================================================
'unit tests for 'IsArrayAllNumeric'
'==============================================================================

'@TestMethod
Public Sub IsArrayAllNumeric_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim V As Variant


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim V() As Variant


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_IncludingNumericStringAllowNumericStringsFalse_ReturnsTrue()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = "100"
   V(2) = 2
   V(3) = Empty

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V, False)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_IncludingNumericStringAllowNumericStringsTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = "100"
   V(2) = 2
   V(3) = Empty

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllNumeric(V, True)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_IncludingNonNumericString_ReturnsFalse()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = "abc"
   V(2) = 2
   V(3) = Empty

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V, True)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_Numeric1DVariantArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = 123
   V(2) = 456
   V(3) = 789

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_1DVariantArrayWithObject_ReturnsFalse()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = 123
   Set V(2) = ThisWorkbook.Worksheets(1).Range("A1")
   V(3) = 789

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_1DVariantArrayWithUnallocatedEntry_ReturnsTrue()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = 123
   V(3) = 789

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_Numeric2DVariantArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim V(1 To 3, 4 To 5) As Variant


   'Arrange:
   V(1, 4) = 123
   V(2, 4) = 456
   V(3, 4) = 789

   V(1, 5) = -5
   V(3, 5) = -10

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_2DVariantArrayWithObject_ReturnsFalse()
   On Error GoTo TestFail

   Dim V(1 To 3, 4 To 5) As Variant


   'Arrange:
   V(1, 4) = 123
   Set V(2, 4) = ThisWorkbook.Worksheets(1).Range("A1")
   V(3, 4) = 789

   V(1, 5) = -5
   V(3, 5) = -10

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowArrayElementsFalse_ReturnsFalse()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = 123
   V(2) = Array(-5)
   V(3) = 789

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowArrayElementsTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = 123
   V(2) = Array(-5)
   V(3) = 789

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllNumeric(V, , True)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowArrayElementsTrue_ReturnsFalse()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = 123
   V(2) = Array(-5, "-5")
   V(3) = 789

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllNumeric(V, , True)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowNumericStringsTrueAllowArrayElementsTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim V(1 To 3) As Variant


   'Arrange:
   V(1) = 123
   V(2) = Array(-5, "-5")
   V(3) = 789

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllNumeric(V, True, True)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'IsArrayAllocated'
'==============================================================================

'@TestMethod
Public Sub IsArrayAllocated_AllocatedArray_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim AllocArray(1 To 3) As Variant


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllocated(AllocArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllocated_UnAllocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim UnAllocArray() As Variant


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllocated(UnAllocArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


Public Sub DemoIsArrayDynamic()

   Dim B As Boolean
   Dim StaticArray(1 To 3) As LongPtr
   Dim DynArray() As LongPtr


   ReDim DynArray(1 To 3)

   B = modArraySupport.IsArrayDynamic(StaticArray)
Debug.Print "IsArrayDynamic StaticArray:", B

   B = modArraySupport.IsArrayDynamic(DynArray)
Debug.Print "IsArrayDynamic DynArray:", B

End Sub


Public Sub DemoIsArrayObjects()

   Dim V(1 To 3) As Variant
   Dim B As Boolean


   Set V(1) = Nothing
   Set V(2) = Range("A1")
   V(3) = Range("A1")

   B = modArraySupport.IsArrayObjects(V, True)
Debug.Print "IsArrayObjects With AllowNothing = True:", B

   B = modArraySupport.IsArrayObjects(V, False)
Debug.Print "IsArrayObjects With AllowNothing = False:", B

End Sub


Public Sub DemoIsNumericDataType()

   Dim V As Variant
   Dim VEmpty As Variant
   Dim S As String
   Dim B As Boolean


   V = 123
   S = "123"

   B = modArraySupport.IsNumericDataType(V)
Debug.Print "IsNumericDataType:", B

   B = modArraySupport.IsNumericDataType(S)
Debug.Print "IsNumericDataType:", B

   B = modArraySupport.IsNumericDataType(VEmpty)
Debug.Print "IsNumericDataType:", B

   V = Array(1, 2, 3)
   B = modArraySupport.IsNumericDataType(V)
Debug.Print "IsNumericDataType:", B

   V = Array("a", "b", "c")
   B = modArraySupport.IsNumericDataType(V)
Debug.Print "IsNumericDataType:", B

End Sub


Public Sub DemoIsVariantArrayConsistent()

   Dim B As Boolean
   Dim V(1 To 3) As Variant


   Set V(1) = Range("A1")
   Set V(2) = Nothing
   Set V(3) = Range("A3")

   B = modArraySupport.IsVariantArrayConsistent(V)
Debug.Print "IsVariantArrayConsistent:", B

End Sub


Public Sub DemoMoveEmptyStringsToEndOfArray()

   Dim B As Boolean
   Dim N As LongPtr
   Dim S(1 To 5) As String


   S(1) = vbNullString
   S(2) = vbNullString
   S(3) = "C"
   S(4) = "D"
   S(5) = "E"

   B = modArraySupport.MoveEmptyStringsToEndOfArray(S)

   If B = True Then
      For N = LBound(S) To UBound(S)
         If S(N) = vbNullString Then
Debug.Print CStr(N), "is vbNullString"
         Else
Debug.Print CStr(N), S(N)
         End If
      Next N
   Else
Debug.Print "MoveEmptyStringsToEndOfArray returned False"
   End If

End Sub


'==============================================================================
'unit tests for 'NumberOfArrayDimensions'
'==============================================================================

'@TestMethod
Public Sub NumberOfArrayDimensions_UnallocatedArray_ReturnsZero()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr
   Dim iNoOfArrDimensions As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = 0
   '===========================================================================


   'Act:
   iNoOfArrDimensions = modArraySupport.NumberOfArrayDimensions(Arr)

   'Assert:
   Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumberOfArrayDimensions_1DArray_ReturnsOne()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 3) As LongPtr
   Dim iNoOfArrDimensions As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = 1
   '===========================================================================


   'Act:
   iNoOfArrDimensions = modArraySupport.NumberOfArrayDimensions(Arr)

   'Assert:
   Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumberOfArrayDimensions_3DArray_ReturnsThree()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 3, 1 To 2, 1 To 1)
   Dim iNoOfArrDimensions As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = 3
   '===========================================================================


   'Act:
   iNoOfArrDimensions = modArraySupport.NumberOfArrayDimensions(Arr)

   'Assert:
   Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


Public Sub DemoNumElements()

   Dim N As LongPtr
   Dim EmptyArray() As LongPtr
   Dim OneArray(1 To 3) As LongPtr
   Dim ThreeArray(1 To 3, 1 To 2, 1 To 1)


   N = modArraySupport.NumElements(EmptyArray, 1)
Debug.Print "NumElements EmptyArray", N

   N = modArraySupport.NumElements(OneArray, 1)
Debug.Print "NumElements OneArray", N

   N = modArraySupport.NumElements(ThreeArray, 3)
Debug.Print "NumElements ThreeArray", N

End Sub


Public Sub DemoResetVariantArrayToDefaults()

   Dim V(1 To 5) As Variant
   Dim B As Boolean
   Dim N As LongPtr


   V(1) = CInt(123)
   V(2) = "abcd"
   Set V(3) = Range("A1")
   V(4) = CDec(123)
   V(5) = Null

   B = modArraySupport.ResetVariantArrayToDefaults(V)

   If B = True Then
      For N = LBound(V) To UBound(V)
         If IsObject(V(N)) = True Then
            If V(N) Is Nothing Then
Debug.Print CStr(N), "Is Nothing"
            Else
Debug.Print CStr(N), "Is Object"
            End If
         Else
Debug.Print CStr(N), TypeName(V(N)), V(N)
         End If
      Next N
   Else
Debug.Print "ResetVariantArrayToDefaults  returned false"
   End If

End Sub


Public Sub DemoReverseArrayInPlace()

   Dim V(1 To 5) As LongPtr
   Dim N As LongPtr
   Dim B As Boolean


   V(1) = 1
   V(2) = 2
   V(3) = 3
   V(4) = 4
   V(5) = 5

   B = modArraySupport.ReverseArrayInPlace(V)

   If B = True Then
Debug.Print "REVERSED ARRAY --------------------------------------"
      For N = LBound(V) To UBound(V)
Debug.Print V(N)
      Next N
   End If

End Sub


Public Sub DemoReverseArrayOfObjectsInPlace()

   Dim B As Boolean
   Dim N As LongPtr
   Dim V(1 To 5) As Object


   Set V(1) = Range("A1")
   Set V(2) = Nothing
   Set V(3) = Range("A3")
   Set V(4) = Range("A4")
   Set V(5) = Range("A5")

   B = modArraySupport.ReverseArrayOfObjectsInPlace(V)

   If B = True Then
Debug.Print "REVERSED ARRAY --------------------------------------"
      For N = LBound(V) To UBound(V)
         If V(N) Is Nothing Then
Debug.Print CStr(N), "Is Nothing"
         Else
Debug.Print CStr(N), V(N).Address
         End If
      Next N
   End If
End Sub


Public Sub DemoSetObjectArrayToNothing()

   Dim StaticArray(1 To 2) As Range
   Dim DynamicArray(1 To 2) As Range
   Dim B As Boolean
   Dim N As LongPtr


   Set StaticArray(1) = Range("A1")
   Set StaticArray(2) = Nothing
   Set DynamicArray(1) = Range("A1")
   Set DynamicArray(2) = Range("A2")

   B = modArraySupport.SetObjectArrayToNothing(StaticArray)

   If B = True Then
      For N = LBound(StaticArray) To UBound(StaticArray)
         If StaticArray(N) Is Nothing Then
Debug.Print CStr(N), "is nothing "
         End If
      Next N
   End If


   B = modArraySupport.SetObjectArrayToNothing(DynamicArray)

   If B = True Then
      For N = LBound(DynamicArray) To UBound(DynamicArray)
         If DynamicArray(N) Is Nothing Then
Debug.Print CStr(N), "is nothing "
         End If
      Next N
   End If

End Sub


Public Sub DemoVectorsToArray()

   Dim A() As Variant
   Dim B As Boolean
   Dim R As LongPtr
   Dim C As LongPtr
   Dim S As String

   Dim AA()
   Dim BB()
   Dim CC() As String


   ReDim AA(0 To 2)
   ReDim BB(1 To 5)
   ReDim CC(2 To 5)


   AA(0) = 16
   AA(1) = 2
   AA(2) = 3
   'AA(3) = 3
   BB(1) = 11
   BB(2) = 22
   BB(3) = 33
   BB(4) = 44
   BB(5) = 55
   CC(2) = "A"
   CC(3) = "B"
   CC(4) = "C"
   CC(5) = "D"

   B = modArraySupport.VectorsToArray(A, AA, BB, CC)

   If B = True Then
      For R = LBound(A, 1) To UBound(A, 1)
         S = vbNullString
         For C = LBound(A, 2) To UBound(A, 2)
            S = S & A(R, C) & " "
         Next C
Debug.Print S
      Next R
   Else
Debug.Print "VectorsToArray Failed"
   End If

End Sub


'==============================================================================
'unit tests for 'TransposeArray'
'==============================================================================

'@TestMethod
Public Sub TransposeArray_ScalarInput_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Const Scalar As LongPtr = 5
   Dim TransposedArr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.TransposeArray(Scalar, TransposedArr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TransposeArray_1DInputArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(2) As LongPtr
   Dim TransposedArr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.TransposeArray(Arr, TransposedArr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TransposeArray_ScalarOutput_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 3, 2 To 5) As LongPtr
   Dim Scalar As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.TransposeArray(Arr, Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TransposeArray_StaticOutputArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 3, 2 To 5) As LongPtr
   Dim TransposedArr(2 To 5, 1 To 3) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.TransposeArray(Arr, TransposedArr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TransposeArray_Valid2DArr_ReturnsTrueAndTransposedArr()
   On Error GoTo TestFail

   Dim Arr() As LongPtr
   Dim TransposedArr() As LongPtr
   Dim i As LongPtr
   Dim j As LongPtr


   'Arrange:
   ReDim Arr(1 To 3, 2 To 5)
   Arr(1, 2) = 1
   Arr(1, 3) = 2
   Arr(1, 4) = 3
   Arr(1, 5) = 33
   Arr(2, 2) = 4
   Arr(2, 3) = 5
   Arr(2, 4) = 6
   Arr(2, 5) = 66
   Arr(3, 2) = 7
   Arr(3, 3) = 8
   Arr(3, 4) = 9
   Arr(3, 5) = 100

   'Act:
   If Not modArraySupport.TransposeArray(Arr, TransposedArr) _
         Then GoTo TestFail

   'Assert:
   For i = LBound(TransposedArr) To UBound(TransposedArr)
      For j = LBound(TransposedArr, 2) To UBound(TransposedArr, 2)
         Assert.AreEqual Arr(j, i), TransposedArr(i, j)
      Next
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'ChangeBoundsOfArray'
'==============================================================================

'@TestMethod
Public Sub ChangeBoundsOfArray_LBGreaterUB_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(2 To 4) As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 5
   Const NewUB As LongPtr = 3
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_ScalarInput_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Const Scalar As LongPtr = 1
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfArray(Scalar, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_StaticArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(2 To 4) As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_2DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(2 To 5, 1 To 1) As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_LongInputArr_ReturnsTrueAndChangedArr()
   On Error GoTo TestFail

   Dim Arr() As LongPtr

   '===========================================================================
   Const NewLB As LongPtr = 20
   Const NewUB As LongPtr = 25
   '===========================================================================
   Dim aExpected(NewLB To NewUB) As LongPtr
      aExpected(20) = 11
      aExpected(21) = 22
      aExpected(22) = 33
      aExpected(23) = 0
      aExpected(24) = 0
      aExpected(25) = 0
   '===========================================================================

   'Arrange:
   ReDim Arr(5 To 7)
   Arr(5) = 11
   Arr(6) = 22
   Arr(7) = 33


   'Act:
   If Not modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_SmallerUBDiffThanSource_ReturnsTrueAndChangedArr()
   On Error GoTo TestFail

   Dim Arr() As LongPtr

   '===========================================================================
   Const NewLB As LongPtr = 20
   Const NewUB As LongPtr = 21
   '===========================================================================
   Dim aExpected(NewLB To NewUB) As LongPtr
      aExpected(20) = 11
      aExpected(21) = 22
   '===========================================================================

   'Arrange:
   ReDim Arr(5 To 7)
   Arr(5) = 11
   Arr(6) = 22
   Arr(7) = 33


   'Act:
   If Not modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_VariantArr_ReturnsTrueAndChangedArr()
   On Error GoTo TestFail

   Dim Arr() As Variant
   Dim i As LongPtr

   '===========================================================================
   Const NewLB As LongPtr = 20
   Const NewUB As LongPtr = 25
   '===========================================================================
   Dim aExpected(NewLB To NewUB) As Variant
      aExpected(20) = Array(1, 2, 3)
      aExpected(21) = Array(4, 5, 6)
      aExpected(22) = Array(7, 8, 9)
      aExpected(23) = Empty
      aExpected(24) = Empty
      aExpected(25) = Empty
   '===========================================================================

   'Arrange:
   ReDim Arr(5 To 7)
   Arr(5) = Array(1, 2, 3)
   Arr(6) = Array(4, 5, 6)
   Arr(7) = Array(7, 8, 9)


   'Act:
   If Not modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB) _
         Then GoTo TestFail

   'Assert:
   For i = NewLB To NewUB
      If IsArray(Arr(i)) Then
         Assert.SequenceEquals aExpected(i), Arr(i)
      Else
         Assert.AreEqual aExpected(i), Arr(i)
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_LongInputArrNoUpperBound_ReturnsTrueAndChangedArr()
   On Error GoTo TestFail

   Dim Arr() As LongPtr
'   Dim i As LongPtr

   '===========================================================================
   Const NewLB As LongPtr = 20
   Const NewUB As LongPtr = 22
   '===========================================================================
   Dim aExpected(NewLB To NewUB) As LongPtr
      aExpected(20) = 11
      aExpected(21) = 22
      aExpected(22) = 33
   '===========================================================================

   'Arrange:
   ReDim Arr(5 To 7)
   Arr(5) = 11
   Arr(6) = 22
   Arr(7) = 33


   'Act:
   If Not modArraySupport.ChangeBoundsOfArray(Arr, NewLB) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'2do: not sure if the test is done right
'     --> is testing for 'Is(Not)Nothing sufficient?
'@TestMethod
Public Sub ChangeBoundsOfArray_RangeArr_ReturnsTrueAndChangedArr()
   On Error GoTo TestFail

   Dim Arr() As Range
   Dim i As LongPtr

   '===========================================================================
   Const NewLB As LongPtr = 20
   Const NewUB As LongPtr = 25
   '===========================================================================
   Dim aExpected(NewLB To NewUB) As Range
   With ThisWorkbook.Worksheets(1)
      Set aExpected(20) = .Range("A1")
      Set aExpected(21) = .Range("A2")
      Set aExpected(22) = .Range("A3")
      Set aExpected(23) = Nothing
      Set aExpected(24) = Nothing
      Set aExpected(25) = Nothing
   End With
   '===========================================================================

   'Arrange:
   ReDim Arr(5 To 7)
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = .Range("A1")
      Set Arr(6) = .Range("A2")
      Set Arr(7) = .Range("A3")
   End With

   'Act:
   If Not modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB) _
         Then GoTo TestFail

   'Assert:
   For i = NewLB To NewUB
      If aExpected(i) Is Nothing Then
         Assert.IsNothing Arr(i)
      Else
         Assert.IsNotNothing Arr(i)
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfArray_CustomClass_ReturnsTrueAndChangedArr()
   On Error GoTo TestFail

   Dim Arr() As cls_4Test_modArraySupport
   Dim i As LongPtr

   '===========================================================================
   Const NewLB As LongPtr = 20
   Const NewUB As LongPtr = 25
   '===========================================================================
   Dim aExpected(NewLB To NewUB) As cls_4Test_modArraySupport
   Set aExpected(20) = New cls_4Test_modArraySupport
   Set aExpected(21) = New cls_4Test_modArraySupport
   Set aExpected(22) = New cls_4Test_modArraySupport
   aExpected(20).Name = "Name 1"
   aExpected(20).Value = 1
   aExpected(21).Name = "Name 2"
   aExpected(21).Value = 3
   aExpected(22).Name = "Name 3"
   aExpected(22).Value = 3
   Set aExpected(23) = Nothing
   Set aExpected(24) = Nothing
   Set aExpected(25) = Nothing
   '===========================================================================

   'Arrange:
   ReDim Arr(5 To 7)
   Set Arr(5) = New cls_4Test_modArraySupport
   Set Arr(6) = New cls_4Test_modArraySupport
   Set Arr(7) = New cls_4Test_modArraySupport
   Arr(5).Name = "Name 1"
   Arr(5).Value = 1
   Arr(6).Name = "Name 2"
   Arr(6).Value = 3
   Arr(7).Name = "Name 3"
   Arr(7).Value = 3

   'Act:
   If Not modArraySupport.ChangeBoundsOfArray(Arr, NewLB, NewUB) _
         Then GoTo TestFail

   'Assert:
   For i = NewLB To NewUB
      If aExpected(i) Is Nothing Then
         Assert.IsNothing Arr(i)
      Else
         Assert.IsNotNothing Arr(i)
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


Public Sub DemoIsArraySorted()

   Dim S(1 To 3) As String
   Dim L(1 To 3) As LongPtr
   Dim R As Variant
   Dim Desc As Boolean


   Desc = True
   S(1) = "B"
   S(2) = "B"
   S(3) = "A"

   L(1) = 1
   L(2) = 2
   L(3) = 3

   R = modArraySupport.IsArraySorted(S, Desc)

   If IsNull(R) = True Then
Debug.Print "Error From IsArraySorted"
   Else
      If R = True Then
Debug.Print "Array Is Sorted"
      Else
Debug.Print "Array is Unsorted"
      End If
   End If

End Sub


'==============================================================================
'unit tests for 'CombineTwoDArrays'
'==============================================================================

'@TestMethod
Public Sub CombineTwoDArrays_ScalarArr1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar1 As LongPtr
   Dim Arr2(1 To 2, 2 To 3) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================


   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Scalar1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_ScalarArr2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 1 To 2) As LongPtr
   Dim Scalar2 As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Scalar2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_1DArr1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3) As LongPtr
   Dim Arr2(1 To 3, 1 To 2) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_3DArr1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 1 To 2, 1 To 4) As LongPtr
   Dim Arr2(1 To 3, 1 To 2) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_1DArr2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 1 To 2) As LongPtr
   Dim Arr2(1 To 3) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_3DArr2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 1 To 2) As LongPtr
   Dim Arr2(1 To 3, 1 To 2, 1 To 4) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_DifferentColNumbers_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 1 To 2) As LongPtr
   Dim Arr2(1 To 3, 1 To 3) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_DifferentLBoundRows_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 1 To 2) As LongPtr
   Dim Arr2(2 To 3, 1 To 2) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_DifferentLBoundCol1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 2 To 3) As LongPtr
   Dim Arr2(1 To 3, 1 To 2) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_DifferentLBoundCol2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1(1 To 3, 1 To 2) As LongPtr
   Dim Arr2(1 To 3, 2 To 3) As LongPtr
   Dim ResArr As Variant
   
   '===========================================================================
   Const aExpected As Variant = Null
   '===========================================================================

   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.AreEqual aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_1BasedStringArrays_ReturnsResultArr()
   On Error GoTo TestFail

   Dim Arr1(1 To 2, 1 To 2) As String
   Dim Arr2(1 To 2, 1 To 2) As String
   Dim ResArr As Variant

   '===========================================================================
   Dim aExpected(1 To 4, 1 To 2) As Variant
      aExpected(1, 1) = "a"
      aExpected(1, 2) = "b"
      aExpected(2, 1) = "c"
      aExpected(2, 2) = "d"
      
      aExpected(3, 1) = "e"
      aExpected(3, 2) = "f"
      aExpected(4, 1) = "g"
      aExpected(4, 2) = "h"
   '===========================================================================


   'Arrange:
   Arr1(1, 1) = "a"
   Arr1(1, 2) = "b"
   Arr1(2, 1) = "c"
   Arr1(2, 2) = "d"
   
   Arr2(1, 1) = "e"
   Arr2(1, 2) = "f"
   Arr2(2, 1) = "g"
   Arr2(2, 2) = "h"

   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_0BasedStringArrays_ReturnsResultArr()
   On Error GoTo TestFail

   Dim Arr1(0 To 1, 0 To 1) As String
   Dim Arr2(0 To 1, 0 To 1) As String
   Dim ResArr As Variant

   '===========================================================================
   Dim aExpected(0 To 3, 0 To 1) As Variant
      aExpected(0, 0) = "a"
      aExpected(0, 1) = "b"
      aExpected(1, 0) = "c"
      aExpected(1, 1) = "d"
      
      aExpected(2, 0) = "e"
      aExpected(2, 1) = "f"
      aExpected(3, 0) = "g"
      aExpected(3, 1) = "h"
   '===========================================================================


   'Arrange:
   Arr1(0, 0) = "a"
   Arr1(0, 1) = "b"
   Arr1(1, 0) = "c"
   Arr1(1, 1) = "d"

   Arr2(0, 0) = "e"
   Arr2(0, 1) = "f"
   Arr2(1, 0) = "g"
   Arr2(1, 1) = "h"

   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_PositiveBasedStringArrays_ReturnsResultArr()
   On Error GoTo TestFail

   Dim Arr1(5 To 6, 5 To 6) As String
   Dim Arr2(5 To 6, 5 To 6) As String
   Dim ResArr As Variant

   '===========================================================================
   Dim aExpected(5 To 8, 5 To 6) As Variant
      aExpected(5, 5) = "a"
      aExpected(5, 6) = "b"
      aExpected(6, 5) = "c"
      aExpected(6, 6) = "d"
      
      aExpected(7, 5) = "e"
      aExpected(7, 6) = "f"
      aExpected(8, 5) = "g"
      aExpected(8, 6) = "h"
   '===========================================================================


   'Arrange:
   Arr1(5, 5) = "a"
   Arr1(5, 6) = "b"
   Arr1(6, 5) = "c"
   Arr1(6, 6) = "d"

   Arr2(5, 5) = "e"
   Arr2(5, 6) = "f"
   Arr2(6, 5) = "g"
   Arr2(6, 6) = "h"

   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)
   
   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_NetativeBasedStringArrays_ReturnsResultArr()
   On Error GoTo TestFail

   Dim Arr1(-6 To -5, -6 To -5) As String
   Dim Arr2(-6 To -5, -6 To -5) As String
   Dim ResArr As Variant

   '===========================================================================
   Dim aExpected(-6 To -3, -6 To -5) As Variant
      aExpected(-6, -6) = "a"
      aExpected(-6, -5) = "b"
      aExpected(-5, -6) = "c"
      aExpected(-5, -5) = "d"

      aExpected(-4, -6) = "e"
      aExpected(-4, -5) = "f"
      aExpected(-3, -6) = "g"
      aExpected(-3, -5) = "h"
   '===========================================================================


   'Arrange:
   Arr1(-6, -6) = "a"
   Arr1(-6, -5) = "b"
   Arr1(-5, -6) = "c"
   Arr1(-5, -5) = "d"

   Arr2(-6, -6) = "e"
   Arr2(-6, -5) = "f"
   Arr2(-5, -6) = "g"
   Arr2(-5, -5) = "h"

   'Act:
   ResArr = modArraySupport.CombineTwoDArrays(Arr1, Arr2)

   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CombineTwoDArrays_NestedStringArrays_ReturnsResultArr()
   On Error GoTo TestFail

   Dim Arr1(1 To 2, 1 To 2) As String
   Dim Arr2(1 To 2, 1 To 2) As String
   Dim Arr3(1 To 2, 1 To 2) As String
   Dim Arr4(1 To 2, 1 To 2) As String
   Dim ResArr As Variant

   '===========================================================================
   Dim aExpected(1 To 8, 1 To 2) As Variant
      aExpected(1, 1) = "a"
      aExpected(1, 2) = "b"
      aExpected(2, 1) = "c"
      aExpected(2, 2) = "d"
      
      aExpected(3, 1) = "e"
      aExpected(3, 2) = "f"
      aExpected(4, 1) = "g"
      aExpected(4, 2) = "h"

      aExpected(5, 1) = "i"
      aExpected(5, 2) = "j"
      aExpected(6, 1) = "k"
      aExpected(6, 2) = "l"
      
      aExpected(7, 1) = "m"
      aExpected(7, 2) = "n"
      aExpected(8, 1) = "o"
      aExpected(8, 2) = "p"
   '===========================================================================


   'Arrange:
   Arr1(1, 1) = "a"
   Arr1(1, 2) = "b"
   Arr1(2, 1) = "c"
   Arr1(2, 2) = "d"
   
   Arr2(1, 1) = "e"
   Arr2(1, 2) = "f"
   Arr2(2, 1) = "g"
   Arr2(2, 2) = "h"

   Arr3(1, 1) = "i"
   Arr3(1, 2) = "j"
   Arr3(2, 1) = "k"
   Arr3(2, 2) = "l"

   Arr4(1, 1) = "m"
   Arr4(1, 2) = "n"
   Arr4(2, 1) = "o"
   Arr4(2, 2) = "p"
   
   'Act:
   ResArr = modArraySupport.CombineTwoDArrays( _
      modArraySupport.CombineTwoDArrays( _
         modArraySupport.CombineTwoDArrays(Arr1, Arr2), _
         Arr3), _
      Arr4 _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


Public Sub DemoExpandArray()

   Dim A As Variant
   Dim B As Variant
   Dim RowNdx As LongPtr
   Dim ColNdx As LongPtr
   Dim S As String


   'ReDim A(-5 To -3, 0 To 3)
   'A(-5, 0) = "a"
   'A(-5, 1) = "b"
   'A(-5, 2) = "c"
   'A(-5, 3) = "d"
   'A(-4, 0) = "e"
   'A(-4, 1) = "f"
   'A(-4, 2) = "g"
   'A(-4, 3) = "h"
   'A(-3, 0) = "i"
   'A(-3, 1) = "j"
   'A(-3, 2) = "k"
   'A(-3, 3) = "l"
   '

   ReDim A(1 To 2, 1 To 4)
   A(1, 1) = "a"
   A(1, 2) = "b"
   A(1, 3) = "c"
   A(1, 4) = "d"
   A(2, 1) = "e"
   A(2, 2) = "f"
   A(2, 3) = "g"
   A(2, 4) = "h"

   Dim C As Variant

Debug.Print "BEFORE:================================="
   For RowNdx = LBound(A, 1) To UBound(A, 1)
      S = vbNullString
      For ColNdx = LBound(A, 2) To UBound(A, 2)
         S = S & A(RowNdx, ColNdx) & " "
      Next ColNdx
Debug.Print S
   Next RowNdx

   S = vbNullString
   B = modArraySupport.ExpandArray(A, 1, 3, "x")

   'C = modArraySupport.ExpandArray( _
         ExpandArray(A, 1, 3, "F"), _
               2, 4, "S")

Debug.Print "AFTER:================================="
   For RowNdx = LBound(B, 1) To UBound(B, 1)
      S = vbNullString
      For ColNdx = LBound(B, 2) To UBound(B, 2)
         S = S & B(RowNdx, ColNdx) & " "
      Next ColNdx
Debug.Print S
   Next RowNdx

'Debug.Print "AFTER:================================="
'   For RowNdx = LBound(C, 1) To UBound(C, 1)
'       S = vbNullString
'       For ColNdx = LBound(C, 2) To UBound(C, 2)
'           S = S & C(RowNdx, ColNdx) & " "
'       Next ColNdx
'       Debug.Print S
'   Next RowNdx

End Sub


Public Sub DemoSwapArrayRows()

   Dim R As LongPtr
   Dim C As LongPtr
   Dim S As String
   Dim A(1 To 3, 1 To 2)
   Dim B()


   A(1, 1) = "a"
   A(1, 2) = "b"
   A(2, 1) = "c"
   A(2, 2) = "d"
   A(3, 1) = "e"
   A(3, 2) = "f"

Debug.Print "BEFORE============================"
   For R = LBound(A, 1) To UBound(A, 1)
      S = vbNullString
      For C = LBound(A, 2) To UBound(A, 2)
         S = S & A(R, C) & " "
      Next C
Debug.Print S
   Next R

   B = modArraySupport.SwapArrayRows(A, 2, 3)

Debug.Print "AFTER============================"
   For R = LBound(B, 1) To UBound(B, 1)
      S = vbNullString
      For C = LBound(B, 2) To UBound(B, 2)
         S = S & B(R, C) & " "
      Next C
Debug.Print S
   Next R

End Sub


Public Sub DemoSwapArrayColumns()

   Dim R As LongPtr
   Dim C As LongPtr
   Dim S As String
   Dim A(1 To 3, 1 To 2)
   Dim B()


   A(1, 1) = "a"
   A(1, 2) = "b"
   A(2, 1) = "c"
   A(2, 2) = "d"
   A(3, 1) = "e"
   A(3, 2) = "f"

Debug.Print "BEFORE============================"
   For R = LBound(A, 1) To UBound(A, 1)
      S = vbNullString
      For C = LBound(A, 2) To UBound(A, 2)
         S = S & A(R, C) & " "
      Next C
Debug.Print S
   Next R

   B = modArraySupport.SwapArrayColumns(A, 1, 2)

Debug.Print "AFTER============================"
   For R = LBound(B, 1) To UBound(B, 1)
      S = vbNullString
      For C = LBound(B, 2) To UBound(B, 2)
         S = S & B(R, C) & " "
      Next C
Debug.Print S
   Next R

End Sub


Public Sub DemoGetColumn()

   Dim InputArr(1 To 2, 1 To 3)
   Dim Result() As LongPtr
   Dim B As Boolean
   Dim N As LongPtr


   InputArr(1, 1) = 1
   InputArr(1, 2) = 2
   InputArr(1, 3) = 3
   InputArr(2, 1) = 4
   InputArr(2, 2) = 5
   InputArr(2, 3) = 6

   B = modArraySupport.GetColumn(InputArr, Result, 3)

   If B = True Then
      For N = LBound(Result) To UBound(Result)
Debug.Print Result(N)
      Next N
   Else
Debug.Print "Error from GetColumn"
   End If

End Sub


Public Sub DemoGetRow()

   Dim InputArr(1 To 2, 1 To 3)
   Dim Result() As LongPtr
   Dim B As Boolean
   Dim N As LongPtr


   InputArr(1, 1) = 1
   InputArr(1, 2) = 2
   InputArr(1, 3) = 3
   InputArr(2, 1) = 4
   InputArr(2, 2) = 5
   InputArr(2, 3) = 6

   B = modArraySupport.GetRow(InputArr, Result, 2)

   If B = True Then
      For N = LBound(Result) To UBound(Result)
Debug.Print Result(N)
      Next N
   Else
Debug.Print "Error from GetRow"
   End If

End Sub
