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
''     --> in 'ChangeBoundsOfVector_VariantArr_ReturnsTrueAndChangedArr' are
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
'unit tests for 'ChangeBoundsOfVector'
'==============================================================================

'@TestMethod
Public Sub ChangeBoundsOfVector_LBGreaterUB_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(2 To 4) As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 5
   Const NewUB As LongPtr = 3
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfVector_ScalarInput_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Const Scalar As LongPtr = 1
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfVector(Scalar, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfVector_StaticArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(2 To 4) As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfVector_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfVector_2DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(2 To 5, 1 To 1) As LongPtr
   
   '===========================================================================
   Const NewLB As LongPtr = 3
   Const NewUB As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfVector_LongInputArr_ReturnsTrueAndChangedArr()
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
   If Not modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfVector_SmallerUBDiffThanSource_ReturnsTrueAndChangedArr()
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
   If Not modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ChangeBoundsOfVector_VariantArr_ReturnsTrueAndChangedArr()
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
   If Not modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
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
Public Sub ChangeBoundsOfVector_LongInputArrNoUpperBound_ReturnsTrueAndChangedArr()
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
   If Not modArraySupport.ChangeBoundsOfVector(Arr, NewLB) _
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
Public Sub ChangeBoundsOfVector_RangeArr_ReturnsTrueAndChangedArr()
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
   If Not modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
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
Public Sub ChangeBoundsOfVector_CustomClass_ReturnsTrueAndChangedArr()
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
   If Not modArraySupport.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
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
Public Sub CombineTwoDArrays_1BasedStringArrays_ReturnsCombinedResultArr()
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
Public Sub CombineTwoDArrays_0BasedStringArrays_ReturnsCombinedResultArr()
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
Public Sub CombineTwoDArrays_PositiveBasedStringArrays_ReturnsCombinedResultArr()
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
Public Sub CombineTwoDArrays_NetativeBasedStringArrays_ReturnsCombinedResultArr()
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
Public Sub CombineTwoDArrays_NestedStringArrays_ReturnsCombinedResultArr()
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


'==============================================================================
'unit tests for 'CompareVectors'
'==============================================================================

'@TestMethod
Public Sub CompareVectors_UnallocatedArrays_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr1() As String
   Dim Arr2() As String
   Dim ResArr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CompareVectors(Arr1, Arr2, ResArr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CompareVectors_LegalAndTextCompare_ReturnsTrueAndResArr()
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
   If Not modArraySupport.CompareVectors(Arr1, Arr2, ResArr, vbTextCompare) _
         Then GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CompareVectors_LegalAndBinaryCompare_ReturnsTrueAndResArr()
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
   If Not modArraySupport.CompareVectors(Arr1, Arr2, ResArr, vbBinaryCompare) _
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

   Dim ResultArray(1) As LongPtr
   Dim ArrayToAppend(1) As LongPtr

   '===========================================================================
   Const CompatabilityCheck As Boolean = True
   '===========================================================================


   'Arrange:
   ResultArray(1) = 8
   ArrayToAppend(1) = 111

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ConcatenateArrays( _
         ResultArray, _
         ArrayToAppend, _
         CompatabilityCheck _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_BothArraysUnallocated_ResultsTrueAndUnallocatedArray()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArray() As LongPtr
   Dim ArrayToAppend() As LongPtr

   '===========================================================================
   Const CompatabilityCheck As Boolean = True
   '===========================================================================


   'Act:
   If Not modArraySupport.ConcatenateArrays( _
         ResultArray, _
         ArrayToAppend, _
         CompatabilityCheck _
   ) Then _
         GoTo TestFail

   'Assert:
   Assert.IsFalse IsArrayAllocated(ResultArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_UnallocatedArrayToAppend_ResultsTrueAndUnchangedResultArray()
   On Error GoTo TestFail

   Dim ResultArray() As LongPtr
   Dim ArrayToAppend() As LongPtr

   '===========================================================================
   Const CompatabilityCheck As Boolean = True
   
   Dim aExpected(1 To 2) As LongPtr
      aExpected(1) = 8
      aExpected(2) = 9
   '===========================================================================


   'Arrange:
   ReDim ResultArray(1 To 2)
   ResultArray(1) = 8
   ResultArray(2) = 9

   'Act:
   If Not modArraySupport.ConcatenateArrays( _
         ResultArray, _
         ArrayToAppend, _
         CompatabilityCheck _
   ) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_IntegerArrayToAppendLongResultArray_ResultsTrueAndResultArray()
   On Error GoTo TestFail

   Dim ResultArray() As LongPtr
   Dim ArrayToAppend(1 To 3) As Integer

   '===========================================================================
   Const CompatabilityCheck As Boolean = True
   
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
   If Not modArraySupport.ConcatenateArrays( _
         ResultArray, _
         ArrayToAppend, _
         CompatabilityCheck _
   ) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_LongArrayToAppendIntegerResultArray_ResultsFalse()
   On Error GoTo TestFail

   Dim ResultArray() As Integer
   Dim ArrayToAppend(1 To 3) As LongPtr

   '===========================================================================
   Const CompatabilityCheck As Boolean = True
   '===========================================================================


   'Arrange:
   ReDim ResultArray(1 To 3)
   ResultArray(1) = 8
   ResultArray(2) = 9
   ResultArray(3) = 10

   ArrayToAppend(1) = 111
   ArrayToAppend(2) = 112
   ArrayToAppend(3) = 113

   'Assert:
   'Act:
   Assert.IsFalse modArraySupport.ConcatenateArrays( _
         ResultArray, _
         ArrayToAppend, _
         CompatabilityCheck _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_LongArrayToAppendIntegerResultArrayFalseCompatabilityCheck_ResultsTrueAndResultArray()
   On Error GoTo TestFail

   Dim ResultArray() As Integer
   Dim ArrayToAppend(1 To 3) As LongPtr

   '===========================================================================
   Const CompatabilityCheck As Boolean = False
   
   Dim aExpected(1 To 6) As Integer
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
   If Not modArraySupport.ConcatenateArrays( _
         ResultArray, _
         ArrayToAppend, _
         CompatabilityCheck _
   ) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ConcatenateArrays_LongArrayToAppendWithLongNumberIntegerResultArrayFalseCompatabilityCheck_ResultsFalse()
   On Error GoTo TestFail
    
   Dim ResultArray() As Integer
   Dim ArrayToAppend(1 To 3) As LongPtr
   Dim Success As Boolean

   '===========================================================================
   Const CompatabilityCheck As Boolean = False
   
   Const ExpectedError As LongPtr = 6
   '===========================================================================

    
   'Arrange:
   ReDim ResultArray(1 To 3)
   ResultArray(1) = 8
   ResultArray(2) = 9
   ResultArray(3) = 10

   ArrayToAppend(1) = 111
   ArrayToAppend(2) = 32768   'no valid Integer
   ArrayToAppend(3) = 113

   'Act:
   Success = modArraySupport.ConcatenateArrays( _
         ResultArray, _
         ArrayToAppend, _
         CompatabilityCheck _
   )
   
   'Assert:
Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
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
'   Const CompatabilityCheck As Boolean = True
'
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
'   If Not modArraySupport.ConcatenateArrays( _
'         ResultArray, _
'         ArrayToAppend, _
'         CompatibilityCheck _
'   ) Then _
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
   Const CompatabilityCheck As Boolean = True
   
   Dim aExpected(0) As Integer
      aExpected(0) = 50
   '===========================================================================


   'Arrange:
   Dest(0) = 50

   'Act:
   If Not modArraySupport.CopyArray( _
         Src, _
         Dest, _
         CompatabilityCheck _
   ) Then _
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

   '===========================================================================
   Const CompatabilityCheck As Boolean = True
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyArray( _
         Src, _
         Dest, _
         CompatabilityCheck _
   )

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
   Const CompatabilityCheck As Boolean = True
   
   Dim aExpected(10 To 11) As LongPtr
      aExpected(10) = 1
      aExpected(11) = 2
   '===========================================================================


   'Arrange:
   Src(1) = 1
   Src(2) = 2
   Src(3) = 3

   'Act:
   If Not modArraySupport.CopyArray( _
         Src, _
         Dest, _
         CompatabilityCheck _
   ) Then _
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
   Const CompatabilityCheck As Boolean = True
   
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
   If Not modArraySupport.CopyArray( _
         Src, _
         Dest, _
         CompatabilityCheck _
   ) Then _
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
   Const CompatabilityCheck As Boolean = False
   
   Dim aExpected(1 To 2) As Integer
      aExpected(1) = 1234
      aExpected(2) = 0
   '===========================================================================


   'Arrange:
   Src(1) = 1234
   Src(2) = 32768       'no valid Integer

   'Act:
   If Not modArraySupport.CopyArray( _
         Src, _
         Dest, _
         CompatabilityCheck _
   ) Then _
         GoTo TestFail

   'Assert:
   Assert.SequenceEquals aExpected, Dest

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'2do: Add tests with Objects


'==============================================================================
'unit tests for 'CopyNonNothingObjectsToVector'
'==============================================================================

'@TestMethod
Public Sub CopyNonNothingObjectsToVector_ScalarResultArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim SourceArray() As Object
   Dim ScalarResult As Object
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyNonNothingObjectsToVector( _
         SourceArray, _
         ScalarResult _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyNonNothingObjectsToVector_StaticResultArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim SourceArray() As Object
   Dim ResultArray(1 To 2) As Object
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyNonNothingObjectsToVector( _
         SourceArray, _
         ResultArray _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyNonNothingObjectsToVector_2DResultArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim SourceArray() As Object
   Dim ResultArray() As Object
   
   
   'Arrange:
   ReDim ResultArray(1 To 2, 1 To 1)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyNonNothingObjectsToVector( _
         SourceArray, _
         ResultArray _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyNonNothingObjectsToVector_NonObjectOnlySourceArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim SourceArray(5 To 6) As Variant
   Dim ResultArray() As Object
   
   
   'Arrange:
   Set SourceArray(5) = Nothing
   SourceArray(6) = 1
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyNonNothingObjectsToVector( _
         SourceArray, _
         ResultArray _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyNonNothingObjectsToVector_ValidNonNothingOnlySourceArray_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim SourceArray(5 To 6) As Variant
   Dim ResultArray() As Object
   Dim i As LongPtr
   
   
   'Arrange:
   Set SourceArray(5) = Nothing
   Set SourceArray(6) = ThisWorkbook.Worksheets(1).Range("B2")
   
   'Act:
   If Not modArraySupport.CopyNonNothingObjectsToVector( _
         SourceArray, _
         ResultArray _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(ResultArray) To UBound(ResultArray)
      Assert.IsNotNothing ResultArray(i)
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyNonNothingObjectsToVector_NothingOnlySourceArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim SourceArray(5 To 6) As Variant
   Dim ResultArray() As Object
   Dim i As LongPtr
   
   
   'Arrange:
   Set SourceArray(5) = Nothing
   Set SourceArray(6) = Nothing
   
   'Act:
   If Not modArraySupport.CopyNonNothingObjectsToVector( _
         SourceArray, _
         ResultArray _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllocated(ResultArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'CopyVectorSubSetToVector'
'==============================================================================

'@TestMethod
Public Sub CopyVectorSubSetToVector_ScalarInput_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 1
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         Scalar, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_ScalarResult_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray() As LongPtr
   Dim ScalarResult As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 1
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ScalarResult, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_UnallocatedInputArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray() As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 1
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_2DInputArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 1
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(0 To 1, 0 To 1)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_2DResultArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 1
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(0 To 1)
   ReDim ResultArray(0 To 1, 0 To 1)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_TooSmallFirstElementToCopy_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = -1
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(0 To 1)
   ReDim ResultArray(0 To 1, 0 To 1)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_TooLargeLastElementToCopy_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 1
   Const LastElementToCopy As LongPtr = 2
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(0 To 1)
   ReDim ResultArray(0 To 1, 0 To 1)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_FirstElementLargerLastElement_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 1
   Const LastElementToCopy As LongPtr = 0
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(0 To 1)
   ReDim ResultArray(0 To 1, 0 To 1)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_NotEnoughRoomInStaticResultArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(0 To 1) As LongPtr
   Dim ResultArray(0 To 1) As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 0
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_TooSmallDestinationElementInStaticResultArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(0 To 1) As LongPtr
   Dim ResultArray(5 To 7) As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 0
   Const LastElementToCopy As LongPtr = 1
   Const DestinationElement As LongPtr = 1
   '===========================================================================
   
   
   'Arrange:
   InputArray(0) = 0
   InputArray(1) = 1
   
   ResultArray(5) = 10
   ResultArray(6) = 20
   ResultArray(7) = 30
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_UnallocatedResultArrayDestinationElementLargerBase_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 11) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 10
   Const DestinationElement As LongPtr = 5
   
   Dim aExpected(1 To 5) As LongPtr
      aExpected(1) = 0
      aExpected(2) = 0
      aExpected(3) = 0
      aExpected(4) = 0
      aExpected(5) = 10
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 10
   InputArray(11) = 20
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_UnallocatedResultArrayLastDestinationElementSmallerBase_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 11) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 10
   Const DestinationElement As LongPtr = -5
   
   Dim aExpected(-5 To 1) As LongPtr
      aExpected(-5) = 10
      aExpected(-4) = 0
      aExpected(-3) = 0
      aExpected(-2) = 0
      aExpected(-1) = 0
      aExpected(0) = 0
      aExpected(1) = 0
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 10
   InputArray(11) = 20
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_UnallocatedResultArrayFromNegToPos_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 13) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 13
   Const DestinationElement As LongPtr = -1
   
   Dim aExpected(-1 To 2) As LongPtr
      aExpected(-1) = 10
      aExpected(0) = 20
      aExpected(1) = 30
      aExpected(2) = 40
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 10
   InputArray(11) = 20
   InputArray(12) = 30
   InputArray(13) = 40
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_UnallocatedResultArray_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 11) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 10
   Const DestinationElement As LongPtr = 1
   
   Dim aExpected(1 To 1) As LongPtr
      aExpected(1) = 0
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 0
   InputArray(11) = 1
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_SubArrayLargerThanAllocatedResultArray1_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 13) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 13
   Const DestinationElement As LongPtr = -1
   
   Dim aExpected(-1 To 2) As LongPtr
      aExpected(-1) = 0
      aExpected(0) = 1
      aExpected(1) = 2
      aExpected(2) = 3
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 0
   InputArray(11) = 1
   InputArray(12) = 2
   InputArray(13) = 3
   
   ReDim ResultArray(0 To 1)
   ResultArray(0) = 10
   ResultArray(1) = 20
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_SubArrayLargerThanAllocatedResultArray2_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 12) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 12
   Const DestinationElement As LongPtr = -1
   
   Dim aExpected(-1 To 1) As LongPtr
      aExpected(-1) = 0
      aExpected(0) = 1
      aExpected(1) = 2
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 0
   InputArray(11) = 1
   InputArray(12) = 2
   
   ReDim ResultArray(0 To 1)
   ResultArray(0) = 10
   ResultArray(1) = 20
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_SubArrayLargerThanAllocatedResultArray3_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 12) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 12
   Const DestinationElement As LongPtr = 1
   
   Dim aExpected(1 To 3) As LongPtr
      aExpected(1) = 0
      aExpected(2) = 1
      aExpected(3) = 2
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 0
   InputArray(11) = 1
   InputArray(12) = 2
   
   ReDim ResultArray(1 To 2)
   ResultArray(1) = 10
   ResultArray(2) = 20
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_TooSmallFirstDestinationElementInDynamicAllocatedResultArray_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 11) As LongPtr
   Dim ResultArray() As LongPtr

   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 11
   Const DestinationElement As LongPtr = -1
   
   Dim aExpected(-1 To 1) As LongPtr
      aExpected(-1) = 0
      aExpected(0) = 1
      aExpected(1) = 20
   '===========================================================================


   'Arrange:
   InputArray(10) = 0
   InputArray(11) = 1

   ReDim ResultArray(0 To 1)
   ResultArray(0) = 10
   ResultArray(1) = 20

   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_TooLargeLastDestinationElementInDynamicAllocatedResultArray_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 11) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 11
   Const DestinationElement As LongPtr = 1
   
   Dim aExpected(0 To 2) As LongPtr
      aExpected(0) = 10
      aExpected(1) = 0
      aExpected(2) = 1
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 0
   InputArray(11) = 1
   
   ReDim ResultArray(0 To 1)
   ResultArray(0) = 10
   ResultArray(1) = 20
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_DestinationElementEvenLargerThanUboundInDynamicAllocatedResultArray_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 11) As LongPtr
   Dim ResultArray() As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 11
   Const DestinationElement As LongPtr = 5
   
   Dim aExpected(0 To 6) As LongPtr
      aExpected(0) = 10
      aExpected(1) = 20
      aExpected(2) = 0
      aExpected(3) = 0
      aExpected(4) = 0
      aExpected(5) = 11
      aExpected(6) = 12
   '===========================================================================
   
   
   'Arrange:
   InputArray(10) = 11
   InputArray(11) = 12
   
   ReDim ResultArray(0 To 1)
   ResultArray(0) = 10
   ResultArray(1) = 20
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub CopyVectorSubSetToVector_TestWithObjects_ReturnsTrueAndResultArray()
   On Error GoTo TestFail

   Dim InputArray(10 To 11) As Object
   Dim ResultArray() As Object
   Dim i As LongPtr
   
   '===========================================================================
   Const FirstElementToCopy As LongPtr = 10
   Const LastElementToCopy As LongPtr = 11
   Const DestinationElement As LongPtr = 6
   
   Dim aExpected(5 To 7) As Object
   With ThisWorkbook.Worksheets(1)
      Set aExpected(5) = .Range("B5")
      Set aExpected(6) = .Range("B10")
      Set aExpected(7) = .Range("B11")
   End With
   '===========================================================================
   
   
   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(10) = .Range("B10")
      Set InputArray(11) = .Range("B11")
   
      ReDim ResultArray(5 To 6)
      Set ResultArray(5) = .Range("B5")
      Set ResultArray(6) = .Range("B6")
   End With
   
   'Act:
   If Not modArraySupport.CopyVectorSubSetToVector( _
         InputArray, _
         ResultArray, _
         FirstElementToCopy, _
         LastElementToCopy, _
         DestinationElement _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(ResultArray) To UBound(ResultArray)
      If ResultArray(i) Is Nothing Then
         Assert.IsNothing aExpected(i)
      Else
         Assert.AreEqual aExpected(i).Address, ResultArray(i).Address
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
   Assert.AreEqual aExpected, aActual

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
Public Sub DataTypeOfArray_Test3DStringArray_ReturnsVbString()
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


'==============================================================================
'unit tests for 'DeleteVectorElement'
'==============================================================================

'@TestMethod
Public Sub DeleteVectorElement_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = False
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.DeleteVectorElement( _
         Scalar, _
         ElementNumer, _
         ResizeDynamic _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray() As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = False
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_2DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 7, 1 To 1) As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = False
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_TooLowElementNumber_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 7) As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 3
   Const ResizeDynamic As Boolean = False
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_TooHighElementNumber_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 7) As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 9
   Const ResizeDynamic As Boolean = False
   '===========================================================================
   
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_RemoveElementOfStaticArray_ReturnsTrueAndModifiedInputArray()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = False
   
   Dim aExpected(5 To 7) As LongPtr
      aExpected(5) = 10
      aExpected(6) = 30
      aExpected(7) = 0
   '===========================================================================
   
   
   'Arrange:
   InputArray(5) = 10
   InputArray(6) = 20
   InputArray(7) = 30
   
   'Act:
   If Not modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_RemoveElementOfStaticObjectArray_ReturnsTrueAndModifiedInputArray()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As Object
   Dim i As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = False
   
   Dim aExpected(5 To 7) As Object
      With ThisWorkbook.Worksheets(1)
         Set aExpected(5) = .Range("B5")
         Set aExpected(6) = .Range("B7")
         Set aExpected(7) = Nothing
      End With
   '===========================================================================
   
   
   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = .Range("B6")
      Set InputArray(7) = .Range("B7")
   End With
   
   'Act:
   If Not modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(InputArray) To UBound(InputArray)
      If InputArray(i) Is Nothing Then
         Assert.IsNothing aExpected(i)
      Else
         Assert.AreEqual aExpected(i).Address, InputArray(i).Address
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_RemoveElementOfDynamicArrayDontResize_ReturnsTrueAndModifiedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = False
   
   Dim aExpected(5 To 7) As LongPtr
      aExpected(5) = 10
      aExpected(6) = 30
      aExpected(7) = 0
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(5 To 7)
   InputArray(5) = 10
   InputArray(6) = 20
   InputArray(7) = 30
   
   'Act:
   If Not modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'2do: why does this test fail?
''@TestMethod
'Public Sub DeleteVectorElement_RemoveElementOfDynamicArrayDontResize2_ReturnsTrueAndModifiedInputArray()
'   On Error GoTo TestFail
'
'   Dim InputArray() As Variant
'
'   '===========================================================================
'   Const ElementNumer As LongPtr = 6
'   Const ResizeDynamic As Boolean = False
'
'   Dim aExpected(5 To 7) As Variant
'      aExpected(5) = "abc"
'      aExpected(6) = "ABC"
'      aExpected(7) = vbNullString
'   '===========================================================================
'
'
'   'Arrange:
'   ReDim InputArray(5 To 7)
'   InputArray(5) = "abc"
'   InputArray(6) = 1234
'   InputArray(7) = "ABC"
'
'   'Act:
'   If Not modArraySupport.DeleteVectorElement( _
'         InputArray, _
'         ElementNumer, _
'         ResizeDynamic _
'   ) Then _
'         GoTo TestFail
'
'   'Assert:
'   Assert.SequenceEquals aExpected, InputArray
'
'TestExit:
'   Exit Sub
'TestFail:
'   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


'@TestMethod
Public Sub DeleteVectorElement_RemoveElementOfDynamicObjectArrayDontResize_ReturnsTrueAndModifiedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As Object
   Dim i As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = False
   
   Dim aExpected(5 To 7) As Object
      With ThisWorkbook.Worksheets(1)
         Set aExpected(5) = .Range("B5")
         Set aExpected(6) = .Range("B7")
         Set aExpected(7) = Nothing
      End With
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(5 To 7)
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = .Range("B6")
      Set InputArray(7) = .Range("B7")
   End With
   
   'Act:
   If Not modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(InputArray) To UBound(InputArray)
      If InputArray(i) Is Nothing Then
         Assert.IsNothing aExpected(i)
      Else
         Assert.AreEqual aExpected(i).Address, InputArray(i).Address
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_RemoveElementOfDynamicArrayResize_ReturnsTrueAndModifiedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = True
   
   Dim aExpected(5 To 6) As LongPtr
      aExpected(5) = 10
      aExpected(6) = 30
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(5 To 7)
   InputArray(5) = 10
   InputArray(6) = 20
   InputArray(7) = 30
   
   'Act:
   If Not modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_RemoveElementOfDynamicObjectArrayResize_ReturnsTrueAndModifiedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As Object
   Dim i As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 6
   Const ResizeDynamic As Boolean = True
   
   Dim aExpected(5 To 6) As Object
      With ThisWorkbook.Worksheets(1)
         Set aExpected(5) = .Range("B5")
         Set aExpected(6) = .Range("B7")
      End With
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(5 To 7)
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = .Range("B6")
      Set InputArray(7) = .Range("B7")
   End With
   
   'Act:
   If Not modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(InputArray) To UBound(InputArray)
      If InputArray(i) Is Nothing Then
         Assert.IsNothing aExpected(i)
      Else
         Assert.AreEqual aExpected(i).Address, InputArray(i).Address
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub DeleteVectorElement_RemoveOnlyElementOfDynamicObjectArrayResize_ReturnsTrueAndModifiedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As String
   Dim i As LongPtr
   
   '===========================================================================
   Const ElementNumer As LongPtr = 5
   Const ResizeDynamic As Boolean = True
   
   Dim aExpected() As String
   '===========================================================================
   
   
   'Arrange:
   ReDim InputArray(5 To 5)
   InputArray(5) = "abc"
   
   'Act:
   If Not modArraySupport.DeleteVectorElement( _
         InputArray, _
         ElementNumer, _
         ResizeDynamic _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.AreEqual aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'ExpandArray'
'==============================================================================

'@TestMethod
Public Sub ExpandArray_NoArray_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 1
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 11
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_UnallocatedArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 1
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 11
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_1DArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 1
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 11
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_3DArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4, 2 To 3) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 1
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 11
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_WhichDimSmallerOne_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 0
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 11
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_WhichDimLargerTwo_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 3
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 11
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_AdditionalElementsSmallerZero_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 1
   Const AdditionalElements As LongPtr = -1
   Const FillValue As LongPtr = 11
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_AdditionalElementsEqualsZero_ReturnsExpandedArray()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 1
   Const AdditionalElements As LongPtr = 0
   Const FillValue As LongPtr = 33
   
   Dim aExpected(5 To 6, 3 To 4) As LongPtr
      aExpected(5, 3) = 10
      aExpected(6, 3) = 11
      aExpected(5, 4) = 20
      aExpected(6, 4) = 21
   '===========================================================================
   
   
   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_AddTwoAdditionalRows_ReturnsExpandedArray()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 1
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 33
   
   Dim aExpected(5 To 8, 3 To 4) As LongPtr
      aExpected(5, 3) = 10
      aExpected(6, 3) = 11
      aExpected(5, 4) = 20
      aExpected(6, 4) = 21
      aExpected(7, 3) = 33
      aExpected(8, 3) = 33
      aExpected(7, 4) = 33
      aExpected(8, 4) = 33
   '===========================================================================
   
   
   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ExpandArray_AddTwoAdditionalCols_ReturnsExpandedArray()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const WhichDim As LongPtr = 2
   Const AdditionalElements As LongPtr = 2
   Const FillValue As LongPtr = 33
   
   Dim aExpected(5 To 6, 3 To 6) As LongPtr
      aExpected(5, 3) = 10
      aExpected(6, 3) = 11
      aExpected(5, 4) = 20
      aExpected(6, 4) = 21
      aExpected(5, 5) = 33
      aExpected(6, 5) = 33
      aExpected(5, 6) = 33
      aExpected(6, 6) = 33
   '===========================================================================
   
   
   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   ResultArr = modArraySupport.ExpandArray( _
         Arr, _
         WhichDim, _
         AdditionalElements, _
         FillValue _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'FirstNonEmptyStringIndexInVector'
'==============================================================================

'@TestMethod
Public Sub FirstNonEmptyStringIndexInVector_NoArray_ReturnsMinusOne()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As String
   Dim aActual As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = -1
   '===========================================================================


   'Act:
   aActual = modArraySupport.FirstNonEmptyStringIndexInVector(Scalar)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub FirstNonEmptyStringIndexInVector_UnallocatedArray_ReturnsMinusOne()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray() As String
   Dim aActual As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = -1
   '===========================================================================


   'Act:
   aActual = modArraySupport.FirstNonEmptyStringIndexInVector(InputArray)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub FirstNonEmptyStringIndexInVector_2DArray_ReturnsMinusOne()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6, 3 To 4) As String
   Dim aActual As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = -1
   '===========================================================================


   'Act:
   aActual = modArraySupport.FirstNonEmptyStringIndexInVector(InputArray)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub FirstNonEmptyStringIndexInVector_NoNonEmptyString_ReturnsMinusOne()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As String
   Dim aActual As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = -1
   '===========================================================================


   'Arrange:
   InputArray(5) = vbNullString
   InputArray(6) = vbNullString
   InputArray(7) = vbNullString
   
   'Act:
   aActual = modArraySupport.FirstNonEmptyStringIndexInVector(InputArray)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub FirstNonEmptyStringIndexInVector_WithNonEmptyStringEntry_ReturnsSeven()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As String
   Dim aActual As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = 7
   '===========================================================================


   'Arrange:
   InputArray(5) = vbNullString
   InputArray(6) = ""
   InputArray(7) = "ghi"
   
   'Act:
   aActual = modArraySupport.FirstNonEmptyStringIndexInVector(InputArray)

   'Assert:
   Assert.AreEqual aExpected, aActual

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'GetColumn'
'==============================================================================

'@TestMethod
Public Sub GetColumn_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 4
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetColumn( _
         Scalar, _
         ResultArr, _
         ColumnNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetColumn_1DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 4
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetColumn( _
         Arr, _
         ResultArr, _
         ColumnNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetColumn_3DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4, -1 To 0) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 4
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetColumn( _
         Arr, _
         ResultArr, _
         ColumnNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetColumn_StaticResultArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr(-5 To -4) As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 4
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetColumn( _
         Arr, _
         ResultArr, _
         ColumnNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetColumn_TooSmallColumnNumber_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 2
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetColumn( _
         Arr, _
         ResultArr, _
         ColumnNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetColumn_TooLargeColumnNumber_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 5
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetColumn( _
         Arr, _
         ResultArr, _
         ColumnNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetColumn_LegalEntries_ReturnsTrueAndResultArr()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 4
   
   Dim aExpected(5 To 6) As LongPtr
      aExpected(5) = 20
      aExpected(6) = 21
   '===========================================================================


   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   If Not modArraySupport.GetColumn( _
         Arr, _
         ResultArr, _
         ColumnNumber _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetColumn_LegalEntriesWithObjects_ReturnsTrueAndResultArr()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As Variant
   Dim ResultArr() As Variant
   Dim i As LongPtr

   '===========================================================================
   Const ColumnNumber As LongPtr = 4
   
   Dim aExpected(5 To 6) As Variant
   With ThisWorkbook.Worksheets(1)
      aExpected(5) = vbNullString
      Set aExpected(6) = .Range("B5")
   End With
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Arr(5, 3) = 10
      Arr(6, 3) = 11
      Arr(5, 4) = vbNullString
      Set Arr(6, 4) = .Range("B5")
   End With
   
   'Act:
   If Not modArraySupport.GetColumn( _
         Arr, _
         ResultArr, _
         ColumnNumber _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(ResultArr) To UBound(ResultArr)
      If IsObject(ResultArr(i)) Then
         If ResultArr(i) Is Nothing Then
            Assert.IsNothing aExpected(i)
         Else
            Assert.AreEqual aExpected(i).Address, ResultArr(i).Address
         End If
      Else
         Assert.AreEqual aExpected(i), ResultArr(i)
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'GetRow'
'==============================================================================

'@TestMethod
Public Sub GetRow_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 6
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetRow( _
         Scalar, _
         ResultArr, _
         RowNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetRow_1DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 6
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetRow( _
         Arr, _
         ResultArr, _
         RowNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetRow_3DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4, -1 To 0) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 6
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetRow( _
         Arr, _
         ResultArr, _
         RowNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetRow_StaticResultArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr(-5 To -4) As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 6
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetRow( _
         Arr, _
         ResultArr, _
         RowNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetRow_TooSmallRowNumber_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 4
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetRow( _
         Arr, _
         ResultArr, _
         RowNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetRow_TooLargeRowNumber_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 7
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.GetRow( _
         Arr, _
         ResultArr, _
         RowNumber _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetRow_LegalEntries_ReturnsTrueAndResultArr()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr() As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 6
   
   Dim aExpected(3 To 4) As LongPtr
      aExpected(3) = 11
      aExpected(4) = 21
   '===========================================================================


   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   If Not modArraySupport.GetRow( _
         Arr, _
         ResultArr, _
         RowNumber _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub GetRow_LegalEntriesWithObjects_ReturnsTrueAndResultArr()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As Variant
   Dim ResultArr() As Variant
   Dim i As LongPtr

   '===========================================================================
   Const RowNumber As LongPtr = 6
   
   Dim aExpected(3 To 4) As Variant
   With ThisWorkbook.Worksheets(1)
      aExpected(3) = vbNullString
      Set aExpected(4) = .Range("B5")
   End With
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Arr(5, 3) = 10
      Arr(6, 3) = vbNullString
      Arr(5, 4) = 20
      Set Arr(6, 4) = .Range("B5")
   End With
   
   'Act:
   If Not modArraySupport.GetRow( _
         Arr, _
         ResultArr, _
         RowNumber _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(ResultArr) To UBound(ResultArr)
      If IsObject(ResultArr(i)) Then
         If ResultArr(i) Is Nothing Then
            Assert.IsNothing aExpected(i)
         Else
            Assert.AreEqual aExpected(i).Address, ResultArr(i).Address
         End If
      Else
         Assert.AreEqual aExpected(i), ResultArr(i)
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'InsertElementIntoVector'
'==============================================================================

'@TestMethod
Public Sub InsertElementIntoVector_StaticInputArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6) As LongPtr

   '===========================================================================
   Const Index As LongPtr = 6
   Const Value As LongPtr = 33
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.InsertElementIntoVector( _
         InputArray, _
         Index, _
         Value _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub InsertElementIntoVector_2DInputArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr

   '===========================================================================
   Const Index As LongPtr = 6
   Const Value As LongPtr = 33
   '===========================================================================


   'Arrange:
   ReDim InputArray(5 To 6, 3 To 4)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.InsertElementIntoVector( _
         InputArray, _
         Index, _
         Value _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub InsertElementIntoVector_TooSmallIndex_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr

   '===========================================================================
   Const Index As LongPtr = 4
   Const Value As LongPtr = 33
   '===========================================================================


   'Arrange:
   ReDim InputArray(5 To 6)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.InsertElementIntoVector( _
         InputArray, _
         Index, _
         Value _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub InsertElementIntoVector_TooLargeIndex_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr

   '===========================================================================
   Const Index As LongPtr = 8
   Const Value As LongPtr = 33
   '===========================================================================


   'Arrange:
   ReDim InputArray(5 To 6)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.InsertElementIntoVector( _
         InputArray, _
         Index, _
         Value _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub InsertElementIntoVector_WrongValueType_ReturnsFalseAndUnchangedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr

   '===========================================================================
   Const Index As LongPtr = 6
   Const Value As String = "abc"
   
   Dim aExpected(5 To 6) As LongPtr
      aExpected(5) = 10
      aExpected(6) = 11
   '===========================================================================


   'Arrange:
   ReDim InputArray(5 To 6)
   InputArray(5) = 10
   InputArray(6) = 11
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.InsertElementIntoVector( _
         InputArray, _
         Index, _
         Value _
   )
   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub InsertElementIntoVector_ValidTestWithLongs_ReturnsTrueAndChangedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As LongPtr

   '===========================================================================
   Const Index As LongPtr = 6
   Const Value As LongPtr = 33
   
   Dim aExpected(5 To 7) As LongPtr
      aExpected(5) = 10
      aExpected(6) = 33
      aExpected(7) = 11
   '===========================================================================


   'Arrange:
   ReDim InputArray(5 To 6)
   InputArray(5) = 10
   InputArray(6) = 11
   
   'Act:
   If Not modArraySupport.InsertElementIntoVector( _
         InputArray, _
         Index, _
         Value _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub InsertElementIntoVector_ValidTestWithStrings_ReturnsTrueAndChangedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As String
   Dim i As LongPtr

   '===========================================================================
   Const Index As LongPtr = 7
   Const Value As String = "XYZ"
   
   Dim aExpected(5 To 7) As String
      aExpected(5) = "abc"
      aExpected(6) = vbNullString
      aExpected(7) = "XYZ"
   '===========================================================================


   'Arrange:
   ReDim InputArray(5 To 6)
   InputArray(5) = "abc"
   InputArray(6) = vbNullString
   
   'Act:
   If Not modArraySupport.InsertElementIntoVector( _
         InputArray, _
         Index, _
         Value _
   ) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(InputArray) To UBound(InputArray)
      Assert.AreEqual aExpected(i), InputArray(i)
   Next
'2do: why does the following line result in an error?
'   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub InsertElementIntoVector_ValidTestWithObjects_ReturnsTrueAndChangedInputArray()
   On Error GoTo TestFail

   Dim InputArray() As Object
   Dim wks As Worksheet
      Set wks = ThisWorkbook.Worksheets(1)
   Dim i As LongPtr
   
   
   With wks

      '========================================================================
      Const Index As LongPtr = 6
      Dim Value As Object
         Set Value = .Range("B2")
      
      Dim aExpected(5 To 7) As Object
         Set aExpected(5) = .Range("B5")
         Set aExpected(6) = .Range("B2")
         Set aExpected(7) = Nothing
      '========================================================================
   
   
      'Arrange:
      ReDim InputArray(5 To 6)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = Nothing
      
      'Act:
      If Not modArraySupport.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
      ) Then _
            GoTo TestFail
      
      'Assert:
      For i = LBound(InputArray) To UBound(InputArray)
         If InputArray(i) Is Nothing Then
            Assert.IsNothing aExpected(i)
         Else
            Assert.AreEqual aExpected(i).Address, InputArray(i).Address
         End If
      Next
   End With

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'IsArrayAllDefault'
'==============================================================================

'@TestMethod
Public Sub IsArrayAllDefault_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllDefault(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_UnallocatedArray_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray() As LongPtr


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_DefaultVariantArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant


   'Arrange:
   InputArray(5) = Empty
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_NonDefaultVariantArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 5) As Variant


   'Arrange:
   InputArray(5) = 10
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_DefaultStringArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As String


   'Arrange:
   InputArray(5) = vbNullString
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_NonDefaultStringArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 5) As String


   'Arrange:
   InputArray(5) = "abc"
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_DefaultNumericArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As LongPtr


   'Arrange:
   InputArray(5) = 0
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_NonDefaultNumericArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 5) As LongPtr


   'Arrange:
   InputArray(5) = -1
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_Default3DNumericArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6, 3 To 4, -2 To -1) As LongPtr


   'Arrange:
   InputArray(5, 3, -2) = 0
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_NonDefault3DNumericArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 6, 3 To 4, -2 To -1) As LongPtr


   'Arrange:
   InputArray(6, 4, -1) = -1
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_DefaultObjectArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Object


   'Arrange:
   Set InputArray(5) = Nothing
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_NonDefaultObjectArray_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 5) As Object


   'Arrange:
   Set InputArray(5) = ThisWorkbook.Worksheets(1).Range("B5")
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayAllDefault(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
Public Sub IsArrayAllNumeric_IncludingNumericStringAllowNumericStringsFalse_ReturnsFalse()
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


'==============================================================================
'unit tests for 'IsArrayDynamic'
'==============================================================================

'@TestMethod
Public Sub IsArrayDynamic_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayDynamic(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayDynamic_UnallocatedArray_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayDynamic(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayDynamic_1DDynamicArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim Arr() As LongPtr


   'Arrange:
   ReDim Arr(5 To 6)
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayDynamic(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayDynamic_1DStaticArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayDynamic(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayDynamic_2DDynamicArray_ReturnsTrue()
   On Error GoTo TestFail

   Dim Arr() As LongPtr


   'Arrange:
   ReDim Arr(5 To 6, 3 To 4)
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayDynamic(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayDynamic_2DStaticArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayDynamic(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'IsArrayObjects'
'==============================================================================

'@TestMethod
Public Sub IsArrayObjects_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr

   '===========================================================================
   Const AllowNothing As Boolean = True
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayObjects(Scalar, AllowNothing)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_LongPtrInputArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6) As LongPtr

   '===========================================================================
   Const AllowNothing As Boolean = True
   '===========================================================================


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayObjects(InputArray, AllowNothing)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_ObjectInputArrayNothingOnlyAllowNothingTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Object
   Dim Element As Variant

   '===========================================================================
   Const AllowNothing As Boolean = True
   '===========================================================================


   'Arrange:
   Set InputArray(5) = Nothing
   Set InputArray(6) = Nothing

   'Act:
   If Not modArraySupport.IsArrayObjects(InputArray, AllowNothing) Then _
         GoTo TestFail
   
   'Assert:
   For Each Element In InputArray
      Assert.IsNothing Element
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_ObjectInputArrayNothingOnlyAllowNothingFalse_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6) As Object

   '===========================================================================
   Const AllowNothing As Boolean = False
   '===========================================================================


   'Arrange:
   Set InputArray(5) = Nothing
   Set InputArray(6) = Nothing

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayObjects(InputArray, AllowNothing)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_ObjectInputArrayNonNothingOnlyAllowNothingTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Object
   Dim Element As Variant

   '===========================================================================
   Const AllowNothing As Boolean = True
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = .Range("B6")
   End With

   'Act:
   If Not modArraySupport.IsArrayObjects(InputArray, AllowNothing) Then _
         GoTo TestFail
   
   'Assert:
   For Each Element In InputArray
      Assert.IsNotNothing Element
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_ObjectInputArrayNonNothingOnlyAllowNothingFalse_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6) As Object
   Dim Element As Variant

   '===========================================================================
   Const AllowNothing As Boolean = False
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = .Range("B6")
   End With

   'Act:
   If Not modArraySupport.IsArrayObjects(InputArray, AllowNothing) Then _
         GoTo TestFail
   
   'Assert:
   For Each Element In InputArray
      Assert.IsNotNothing Element
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_VariantInputArrayAllowNothingFalse_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6) As Variant
   Dim Element As Variant

   '===========================================================================
   Const AllowNothing As Boolean = False
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = Nothing
   End With

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayObjects(InputArray, AllowNothing)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_VariantInputArrayAllowNothingTrue_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6) As Variant
   Dim Element As Variant

   '===========================================================================
   Const AllowNothing As Boolean = True
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5) = .Range("B5")
      Set InputArray(6) = Nothing
   End With

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayObjects(InputArray, AllowNothing)
   
TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_2DVariantInputArrayAllowNothingFalse_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6, 3 To 4) As Variant
   Dim Element As Variant

   '===========================================================================
   Const AllowNothing As Boolean = False
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5, 3) = .Range("B5")
      Set InputArray(6, 3) = .Range("B6")
      Set InputArray(5, 4) = Nothing
      Set InputArray(6, 4) = Nothing
   End With

   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsArrayObjects(InputArray, AllowNothing)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayObjects_2DVariantInputArrayAllowNothingTrue_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6, 3 To 4) As Variant
   Dim Element As Variant

   '===========================================================================
   Const AllowNothing As Boolean = True
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set InputArray(5, 3) = .Range("B5")
      Set InputArray(6, 3) = .Range("B6")
      Set InputArray(5, 4) = Nothing
      Set InputArray(6, 4) = Nothing
   End With

   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsArrayObjects(InputArray, AllowNothing)
   
TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'IsNumericDataType'
'==============================================================================

'@TestMethod
Public Sub IsNumericDataType_LongPtrScalar_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsNumericDataType(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_CurrencyScalar_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As Currency


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsNumericDataType(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_StringScalar_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As String


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_ObjectScalar_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_VariantScalarUninitialized_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As Variant


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_VariantScalarNumericContent_ReturnsTrue()
   On Error GoTo TestFail

   Dim Scalar As Variant


   'Arrange:
   Scalar = 3
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsNumericDataType(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_VariantScalarNonNumericContent_ReturnsFalse()
   On Error GoTo TestFail

   Dim Scalar As Variant


   'Arrange:
   Scalar = "abc"
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_LongPtrArrayUnallocated_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_LongPtrStaticArray_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_CurrencyArray_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Currency


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_StringArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As String


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_ObjectArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_VariantArrayUnallocated_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Variant


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_StaticVariantArrayNumericContent_ReturnsTrue()
   On Error GoTo TestFail

   Dim Arr(5 To 6) As Variant


   'Arrange:
   Arr(5) = 3
   Arr(6) = 7.8
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_StaticVariantArrayMixedContentNumericFirst_ReturnsFalse()
   On Error GoTo TestFail

   Dim Arr(5 To 6) As Variant


   'Arrange:
   Arr(5) = -2
   Arr(6) = "abc"
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsNumericDataType_StaticVariantArrayMixedContentNonNumericFirst_ReturnsFalse()
   On Error GoTo TestFail

   Dim Arr(5 To 6) As Variant


   'Arrange:
   Arr(5) = "abc"
   Arr(6) = -2
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsNumericDataType(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'IsVariantArrayConsistent'
'==============================================================================

'@TestMethod
Public Sub IsVariantArrayConsistent_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsVariantArrayConsistent(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_AllocatedLongTypeArray_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_AllocatedObjectTypeArray_ReturnsTrue()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As Object


   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_AllocatedVariantTypeArrayConsistentIntegers_ReturnsTrue()
   On Error GoTo TestFail

   Dim Arr(5 To 6) As Variant


   'Arrange:
   Arr(5) = -100
   Arr(6) = 3
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_AllocatedVariantTypeArrayConsistentObjects_ReturnsTrue()
   On Error GoTo TestFail

   Dim Arr(5 To 7) As Variant


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = .Range("B5")
      Set Arr(6) = Nothing
      Set Arr(7) = .Range("B7")
   End With
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_AllocatedVariantTypeArrayInconsistentTypes_ReturnsFalse()
   On Error GoTo TestFail

   Dim Arr(5 To 6) As Variant


   'Arrange:
   Arr(5) = -100
   Arr(6) = "abc"
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_2DAllocatedVariantTypeArrayConsistentIntegers_ReturnsTrue()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As Variant


   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_2DAllocatedVariantTypeArrayConsistentObjects_ReturnsTrue()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As Variant


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5, 3) = .Range("B5")
      Set Arr(6, 3) = Nothing
      Set Arr(5, 4) = .Range("B7")
      Set Arr(6, 4) = Nothing
   End With
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVariantArrayConsistent_2DAllocatedVariantTypeArrayInconsistentTypes_ReturnsFalse()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As Variant


   'Arrange:
   Arr(5, 3) = -100
   Arr(6, 3) = "abc"
   Arr(5, 4) = Empty
   Set Arr(6, 4) = Nothing
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.IsVariantArrayConsistent(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'IsVectorSorted'
'==============================================================================

'@TestMethod
Public Sub IsVectorSorted_NoArray_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         Scalar, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue IsNull(aResult)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_UnallocatedArray_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray() As LongPtr
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue IsNull(aResult)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_2DArray_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6, 3 To 4) As LongPtr
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue IsNull(aResult)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_ObjectArray_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6) As Object
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue IsNull(aResult)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_StringArrayDescendingFalse_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As String
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   InputArray(5) = "ABC"
   InputArray(6) = "abc"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayContainingObjectDescendingFalse_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   Set InputArray(5) = ThisWorkbook.Worksheets(1).Range("B5")
   InputArray(6) = vbNullString
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArraySmallNumericStringPlusLargerNumberDescendingFalse_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   InputArray(5) = "45"
   InputArray(6) = 123
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArraySmallNumberPlusLargerNumericStringDescendingFalse_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   InputArray(5) = 45
   InputArray(6) = "123"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayLargeNumberPlusSmallNumericStringDescendingFalse_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   '(it seems that the numbers are always considered smaller than any string)
   InputArray(5) = 9
   InputArray(6) = ""
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayNumberPlusStringDescendingFalse_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   InputArray(5) = 45
   InputArray(6) = "abc"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayNumberPlusStringsDescendingFalse_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 8) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   '(but then strings seem to be compared as usual)
   InputArray(5) = 5
   InputArray(6) = "1"
   InputArray(7) = "Abc"
   InputArray(8) = "defg"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayNumberPlusStrings2DescendingFalse_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 8) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = False
   '===========================================================================


   'Arrange:
   InputArray(5) = 5
   InputArray(6) = "zbc"
   InputArray(7) = "defg"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_StringArrayDescendingTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As String
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   InputArray(5) = "ABC"
   InputArray(6) = "abc"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayContainingObjectDescendingTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   Set InputArray(5) = ThisWorkbook.Worksheets(1).Range("B5")
   InputArray(6) = vbNullString
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArraySmallNumericStringPlusLargerNumberDescendingTrue_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   InputArray(5) = "45"
   InputArray(6) = 123
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsTrue aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArraySmallNumberPlusLargerNumericStringDescendingTrue_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   InputArray(5) = 45
   InputArray(6) = "123"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayLargeNumberPlusSmallNumericStringDescendingTrue_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   '(it seems that the numbers are always considered smaller than any string)
   InputArray(5) = 9
   InputArray(6) = ""
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayNumberPlusStringDescendingTrue_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 6) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   InputArray(5) = 45
   InputArray(6) = "abc"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayNumberPlusStringsDescendingTrue_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 8) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   '(but then strings seem to be compared as usual)
   InputArray(5) = 5
   InputArray(6) = "1"
   InputArray(7) = "Abc"
   InputArray(8) = "defg"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsVectorSorted_VariantArrayNumberPlusStrings2DescendingTrue_ReturnsFalse()
   On Error GoTo TestFail

   Dim InputArray(5 To 8) As Variant
   Dim aResult As Variant

   '===========================================================================
   Const Descending As Boolean = True
   '===========================================================================


   'Arrange:
   InputArray(5) = 5
   InputArray(6) = "zbc"
   InputArray(7) = "defg"
   
   'Act:
   aResult = modArraySupport.IsVectorSorted( _
         InputArray, _
         Descending _
   )
   
   'Assert:
   Assert.IsFalse aResult

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'MoveEmptyStringsToEndOfArray'
'==============================================================================

'@TestMethod
Public Sub MoveEmptyStringsToEndOfVector_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As String


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.MoveEmptyStringsToEndOfVector(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub MoveEmptyStringsToEndOfVector_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray() As String


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub MoveEmptyStringsToEndOfVector_2DArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim InputArray(5 To 6, 3 To 4) As String


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub MoveEmptyStringsToEndOfVector_vbNullStringArrayOnly_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As String

   
   'Arrange:
   InputArray(5) = vbNullString
   InputArray(6) = vbNullString
   InputArray(7) = vbNullString
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub MoveEmptyStringsToEndOfVector_NonvbNullStringArrayOnly_ReturnsTrue()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As String


   'Arrange:
   InputArray(5) = "abc"
   InputArray(6) = "def"
   InputArray(7) = "ghi"
   
   'Act:
   'Assert:
   Assert.IsTrue modArraySupport.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub MoveEmptyStringsToEndOfVector_StringArray_ReturnsTrueAndModifiedArr()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As String
   Dim i As LongPtr

   '===========================================================================
   Dim aExpected(5 To 7) As String
      aExpected(5) = "abc"
      aExpected(6) = vbNullString
      aExpected(7) = vbNullString
   '===========================================================================


   'Arrange:
   InputArray(5) = vbNullString
   InputArray(6) = vbNullString
   InputArray(7) = "abc"
   
   'Act:
   If Not modArraySupport.MoveEmptyStringsToEndOfVector(InputArray) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(InputArray) To UBound(InputArray)
      Assert.AreEqual aExpected(i), InputArray(i)
   Next
'   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub MoveEmptyStringsToEndOfVector_VariantArray_ReturnsTrueAndModifiedArr()
   On Error GoTo TestFail

   Dim InputArray(5 To 7) As Variant
   Dim i As LongPtr

   '===========================================================================
   Dim aExpected(5 To 7) As Variant
      aExpected(5) = "abc"
      aExpected(6) = "def"
      aExpected(7) = vbNullString
   '===========================================================================


   'Arrange:
   InputArray(5) = vbNullString
   InputArray(6) = "abc"
   InputArray(7) = "def"
   
   'Act:
   If Not modArraySupport.MoveEmptyStringsToEndOfVector(InputArray) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(InputArray) To UBound(InputArray)
      Assert.AreEqual aExpected(i), InputArray(i)
   Next
'   Assert.SequenceEquals aExpected, InputArray

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''@TestMethod
'Public Sub MoveEmptyStringsToEndOfVector_StringArray2_ReturnsTrueAndModifiedArr()
'   On Error GoTo TestFail
'
'   Dim Arr As Variant
'   Dim InputArray() As String
'   Dim i As LongPtr
'
'   '===========================================================================
'   Dim aExpected() As String
'   '===========================================================================
'
'
'   'Arrange:
''move entries in the shown range 3 cells down
'   Arr = ThisWorkbook.Worksheets(1).Range("B32:B44")
'
'   'Act:
'    If Not modArraySupport.GetColumn(Arr, InputArray, 1) Then GoTo TestFail
'   If Not modArraySupport.MoveEmptyStringsToEndOfVector(InputArray) Then _
'         GoTo TestFail
'   Arr = ThisWorkbook.Worksheets(1).Range("B35:B47")
'   If Not modArraySupport.GetColumn(Arr, aExpected, 1) Then GoTo TestFail
'
'   'Assert:
'   For i = LBound(InputArray) To UBound(InputArray)
'      Assert.AreEqual aExpected(i), InputArray(i)
'   Next
''   Assert.SequenceEquals aExpected, InputArray
'
'TestExit:
'   Exit Sub
'TestFail:
'   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


'==============================================================================
'unit tests for 'NumberOfArrayDimensions'
'==============================================================================

'@TestMethod
Public Sub NumberOfArrayDimensions_UnallocatedLongArray_ReturnsZero()
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
Public Sub NumberOfArrayDimensions_UnallocatedVariantArray_ReturnsZero()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Variant
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
Public Sub NumberOfArrayDimensions_UnallocatedObjectArray_ReturnsZero()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Object
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


'==============================================================================
'unit tests for 'NumElements'
'==============================================================================

'@TestMethod
Public Sub NumElements_NoArray_ReturnsZero()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const Dimension As LongPtr = 1
   
   Const aExpected As LongPtr = 0
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Scalar, Dimension)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumElements_UnallocatedArray_ReturnsZero()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const Dimension As LongPtr = 1
   
   Const aExpected As LongPtr = 0
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Arr, Dimension)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumElements_DimensionLowerOne_ReturnsZero()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4, 1 To 1) As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const Dimension As LongPtr = 0
   
   Const aExpected As LongPtr = 0
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Arr, Dimension)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumElements_DimensionHigherNoOfArrDimensions_ReturnsZero()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4, 1 To 1) As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const Dimension As LongPtr = 4
   
   Const aExpected As LongPtr = 0
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Arr, Dimension)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumElements_DimensionOne_ReturnsThree()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4, 1 To 1) As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const Dimension As LongPtr = 1
   
   Const aExpected As LongPtr = 3
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Arr, Dimension)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumElements_DimensionTwo_ReturnsTwo()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4, 1 To 1) As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const Dimension As LongPtr = 2
   
   Const aExpected As LongPtr = 2
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Arr, Dimension)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumElements_DimensionThree_ReturnsOne()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4, 1 To 1) As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const Dimension As LongPtr = 3
   
   Const aExpected As LongPtr = 1
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Arr, Dimension)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub NumElements_DefaultDimension_ReturnsThree()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4, 1 To 1) As LongPtr
   Dim iNoOfElements As LongPtr

   '===========================================================================
   Const aExpected As LongPtr = 3
   '===========================================================================


   'Act:
   iNoOfElements = modArraySupport.NumElements(Arr)

   'Assert:
   Assert.AreEqual aExpected, iNoOfElements

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'ResetVariantArrayToDefaults'
'==============================================================================

'@TestMethod
Public Sub ResetVariantArrayToDefaults_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ResetVariantArrayToDefaults(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ResetVariantArrayToDefaults_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ResetVariantArrayToDefaults(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ResetVariantArrayToDefaults_4DArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(1 To 8, 4 To 5, 3 To 3, 2 To 2) As Variant


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ResetVariantArrayToDefaults(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ResetVariantArrayToDefaults_AllSetVariableToDefaultElementsIn1DArr_ReturnsTrueAndResettedArr()
   On Error GoTo TestFail

   Dim Arr(1 To 15) As Variant
   Dim i As LongPtr

   '===========================================================================
   Dim aExpected(1 To 15) As Variant
      Set aExpected(1) = Nothing
      aExpected(2) = Array()
         SetVariableToDefault aExpected(2)
      aExpected(3) = False
      aExpected(4) = CByte(0)
      aExpected(5) = CCur(0)
      aExpected(6) = CDate(0)
      aExpected(7) = CDec(0)
      aExpected(8) = CDbl(0)
      aExpected(9) = Empty
      aExpected(10) = Empty
      aExpected(11) = CInt(0)
      aExpected(12) = CLngPtr(0)
      aExpected(13) = Null
      aExpected(14) = CSng(0)
      aExpected(15) = vbNullString
   '===========================================================================


   'Arrange:
   Set Arr(1) = ThisWorkbook.Worksheets(1).Range("B5")
   Arr(2) = Array(123)
   Arr(3) = True
   Arr(4) = CByte(1)
   Arr(5) = CCur(1)
   Arr(6) = #2/12/1969#
   Arr(7) = CDec(10000000.0587)
   Arr(8) = CDbl(-123.456)
   Arr(9) = Empty
   Arr(10) = CVErr(xlErrNA)
   Arr(11) = CInt(2345.5678)
   Arr(12) = CLngPtr(123456789)
   Arr(13) = Null
   Arr(14) = CSng(654.321)
   Arr(15) = "abc"
   
   'Act:
   If Not modArraySupport.ResetVariantArrayToDefaults(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(Arr) To UBound(Arr)
      If IsObject(Arr(i)) Then
         Assert.IsNothing Arr(i)
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
Public Sub ResetVariantArrayToDefaults_AllSetVariableToDefaultElementsIn2DArr_ReturnsTrueAndResettedArr()
   On Error GoTo TestFail

   Dim Arr(1 To 8, 4 To 5) As Variant
   Dim i As LongPtr
   Dim j As LongPtr

   '===========================================================================
   Dim aExpected(1 To 8, 4 To 5) As Variant
      Set aExpected(1, 4) = Nothing
      aExpected(2, 4) = Array()
         SetVariableToDefault aExpected(2, 4)
      aExpected(3, 4) = False
      aExpected(4, 4) = CByte(0)
      aExpected(5, 4) = CCur(0)
      aExpected(6, 4) = CDate(0)
      aExpected(7, 4) = CDec(0)
      aExpected(8, 4) = CDbl(0)
   
      aExpected(1, 5) = Empty
      aExpected(2, 5) = Empty
      aExpected(3, 5) = CInt(0)
      aExpected(4, 5) = CLngPtr(0)
      aExpected(5, 5) = Null
      aExpected(6, 5) = CSng(0)
      aExpected(7, 5) = vbNullString
      aExpected(8, 5) = Empty     'non-initialized Variant entry
   '===========================================================================


   'Arrange:
   Set Arr(1, 4) = ThisWorkbook.Worksheets(1).Range("B5")
   Arr(2, 4) = Array(123)
   Arr(3, 4) = True
   Arr(4, 4) = CByte(1)
   Arr(5, 4) = CCur(1)
   Arr(6, 4) = #2/12/1969#
   Arr(7, 4) = CDec(10000000.0587)
   Arr(8, 4) = CDbl(-123.456)

   Arr(1, 5) = Empty
   Arr(2, 5) = CVErr(xlErrNA)
   Arr(3, 5) = CInt(2345.5678)
   Arr(4, 5) = CLngPtr(123456789)
   Arr(5, 5) = Null
   Arr(6, 5) = CSng(654.321)
   Arr(7, 5) = "abc"
   
   'Act:
   If Not modArraySupport.ResetVariantArrayToDefaults(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(Arr, 1) To UBound(Arr, 1)
      For j = LBound(Arr, 2) To UBound(Arr, 2)
         If IsObject(Arr(i, j)) Then
            Assert.IsNothing Arr(i, j)
         Else
            Assert.AreEqual aExpected(i, j), Arr(i, j)
         End If
      Next
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ResetVariantArrayToDefaults_AllSetVariableToDefaultElementsIn3DArr_ReturnsTrueAndResettedArr()
   On Error GoTo TestFail

   Dim Arr(1 To 8, 4 To 5, 3 To 3) As Variant
   Dim i As LongPtr
   Dim j As LongPtr
   Dim k As LongPtr

   '===========================================================================
   Dim aExpected(1 To 8, 4 To 5, 3 To 3) As Variant
      Set aExpected(1, 4, 3) = Nothing
      aExpected(2, 4, 3) = Array()
         SetVariableToDefault aExpected(2, 4, 3)
      aExpected(3, 4, 3) = False
      aExpected(4, 4, 3) = CByte(0)
      aExpected(5, 4, 3) = CCur(0)
      aExpected(6, 4, 3) = CDate(0)
      aExpected(7, 4, 3) = CDec(0)
      aExpected(8, 4, 3) = CDbl(0)
   
      aExpected(1, 5, 3) = Empty
      aExpected(2, 5, 3) = Empty
      aExpected(3, 5, 3) = CInt(0)
      aExpected(4, 5, 3) = CLngPtr(0)
      aExpected(5, 5, 3) = Null
      aExpected(6, 5, 3) = CSng(0)
      aExpected(7, 5, 3) = vbNullString
      aExpected(8, 5, 3) = Empty     'non-initialized Variant entry
   '===========================================================================


   'Arrange:
   Set Arr(1, 4, 3) = ThisWorkbook.Worksheets(1).Range("B5")
   Arr(2, 4, 3) = Array(123)
   Arr(3, 4, 3) = True
   Arr(4, 4, 3) = CByte(1)
   Arr(5, 4, 3) = CCur(1)
   Arr(6, 4, 3) = #2/12/1969#
   Arr(7, 4, 3) = CDec(10000000.0587)
   Arr(8, 4, 3) = CDbl(-123.456)

   Arr(1, 5, 3) = Empty
   Arr(2, 5, 3) = CVErr(xlErrNA)
   Arr(3, 5, 3) = CInt(2345.5678)
   Arr(4, 5, 3) = CLngPtr(123456789)
   Arr(5, 5, 3) = Null
   Arr(6, 5, 3) = CSng(654.321)
   Arr(7, 5, 3) = "abc"
   
   'Act:
   If Not modArraySupport.ResetVariantArrayToDefaults(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(Arr, 1) To UBound(Arr, 1)
      For j = LBound(Arr, 2) To UBound(Arr, 2)
         For k = LBound(Arr, 3) To UBound(Arr, 3)
            If IsObject(Arr(i, j, k)) Then
               Assert.IsNothing Arr(i, j, k)
            Else
               Assert.AreEqual aExpected(i, j, k), Arr(i, j, k)
            End If
         Next
      Next
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'ReverseArrayInPlace'
'==============================================================================

'@TestMethod
Public Sub ReverseVectorInPlace_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ReverseVectorInPlace(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorInPlace_UnallocatedArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ReverseVectorInPlace(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorInPlace_2DArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ReverseVectorInPlace(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorInPlace_ValidEven1DArr_ReturnsTrueAndReversedArr()
   On Error GoTo TestFail

   Dim Arr(5 To 8) As LongPtr

   '===========================================================================
   Dim aExpected(5 To 8) As LongPtr
      aExpected(5) = 8
      aExpected(6) = 7
      aExpected(7) = 6
      aExpected(8) = 5
   '===========================================================================


   'Arrange:
   Arr(5) = 5
   Arr(6) = 6
   Arr(7) = 7
   Arr(8) = 8
   
   'Act:
   If Not modArraySupport.ReverseVectorInPlace(Arr) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorInPlace_ValidEven1DVariantArr_ReturnsTrueAndReversedArr()
   On Error GoTo TestFail

   Dim Arr(5 To 8) As Variant

   '===========================================================================
   Dim aExpected(5 To 8) As Variant
      aExpected(5) = 8
      aExpected(6) = "ghi"
      aExpected(7) = 6
      aExpected(8) = "abc"
   '===========================================================================


   'Arrange:
   Arr(5) = "abc"
   Arr(6) = 6
   Arr(7) = "ghi"
   Arr(8) = 8
   
   'Act:
   If Not modArraySupport.ReverseVectorInPlace(Arr) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorInPlace_1DVariantArrWithObject_ReturnsTrueAndReversedArr()
   On Error GoTo TestFail

   Dim Arr(5 To 6) As Variant

   '===========================================================================
   Dim aExpected(5 To 6) As Variant
      aExpected(5) = "AreDataTypesCompatible"   '*content* of the below cell
      aExpected(6) = 5
   '===========================================================================


   'Arrange:
   Arr(5) = 5
   Set Arr(6) = ThisWorkbook.Worksheets(1).Range("B5")
   
   'Act:
   If Not modArraySupport.ReverseVectorInPlace(Arr) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorInPlace_ValidOdd1DArr_ReturnsTrueAndReversedArr()
   On Error GoTo TestFail

   Dim Arr(5 To 9) As LongPtr

   '===========================================================================
   Dim aExpected(5 To 9) As LongPtr
      aExpected(5) = 9
      aExpected(6) = 8
      aExpected(7) = 7
      aExpected(8) = 6
      aExpected(9) = 5
   '===========================================================================


   'Arrange:
   Arr(5) = 5
   Arr(6) = 6
   Arr(7) = 7
   Arr(8) = 8
   Arr(9) = 9
   
   'Act:
   If Not modArraySupport.ReverseVectorInPlace(Arr) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, Arr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'ReverseVectorOfObjectsInPlace'
'==============================================================================

'@TestMethod
Public Sub ReverseVectorOfObjectsInPlace_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ReverseVectorOfObjectsInPlace(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorOfObjectsInPlace_UnallocatedObjectArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ReverseVectorOfObjectsInPlace(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorOfObjectsInPlace_2DObjectArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4) As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ReverseVectorOfObjectsInPlace(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorOfObjectsInPlace_ValidEven1DObjectArr_ReturnsTrueAndReversedArr()
   On Error GoTo TestFail

   Dim Arr(5 To 8) As Object
   Dim i As LongPtr

   '===========================================================================
   Dim aExpected(5 To 8) As Object
   With ThisWorkbook.Worksheets(1)
      Set aExpected(5) = .Range("B8")
      Set aExpected(6) = .Range("B7")
      Set aExpected(7) = .Range("B6")
      Set aExpected(8) = .Range("B5")
   End With
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = .Range("B5")
      Set Arr(6) = .Range("B6")
      Set Arr(7) = .Range("B7")
      Set Arr(8) = .Range("B8")
   End With
   
   'Act:
   If Not modArraySupport.ReverseVectorOfObjectsInPlace(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(Arr) To UBound(Arr)
      If Arr(i) Is Nothing Then
         Assert.IsNothing aExpected(i)
      Else
         Assert.AreEqual Arr(i).Address, Arr(i).Address
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorOfObjectsInPlace_ValidEven1DVariantArr_ReturnsTrueAndReversedArr()
   On Error GoTo TestFail

   Dim Arr(5 To 8) As Variant
   Dim i As LongPtr

   '===========================================================================
   Dim aExpected(5 To 8) As Variant
   With ThisWorkbook.Worksheets(1)
      Set aExpected(5) = .Range("B8")
      Set aExpected(6) = Nothing
      Set aExpected(7) = .Range("B6")
      Set aExpected(8) = Nothing
   End With
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = Nothing
      Set Arr(6) = .Range("B6")
      Set Arr(7) = Nothing
      Set Arr(8) = .Range("B8")
   End With
   
   'Act:
   If Not modArraySupport.ReverseVectorOfObjectsInPlace(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(Arr) To UBound(Arr)
      If Arr(i) Is Nothing Then
         Assert.IsNothing aExpected(i)
      Else
         Assert.AreEqual Arr(i).Address, Arr(i).Address
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorOfObjectsInPlace_1DVariantArrWithNonObject_ReturnsFalse()
   On Error GoTo TestFail

   Dim Arr(5 To 6) As Variant


   'Arrange:
   Set Arr(5) = ThisWorkbook.Worksheets(1).Range("B5")
   Arr(6) = 6
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.ReverseVectorOfObjectsInPlace(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub ReverseVectorOfObjectsInPlace_ValidOdd1DObjectArr_ReturnsTrueAndReversedArr()
   On Error GoTo TestFail

   Dim Arr(5 To 9) As Object
   Dim i As LongPtr

   '===========================================================================
   Dim aExpected(5 To 9) As Object
   With ThisWorkbook.Worksheets(1)
      Set aExpected(5) = .Range("B9")
      Set aExpected(6) = Nothing
      Set aExpected(7) = .Range("B7")
      Set aExpected(8) = .Range("B6")
      Set aExpected(9) = .Range("B5")
   End With
   '===========================================================================


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = .Range("B5")
      Set Arr(6) = .Range("B6")
      Set Arr(7) = .Range("B7")
      Set Arr(8) = Nothing
      Set Arr(9) = .Range("B9")
   End With
   
   'Act:
   If Not modArraySupport.ReverseVectorOfObjectsInPlace(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For i = LBound(Arr) To UBound(Arr)
      If Arr(i) Is Nothing Then
         Assert.IsNothing aExpected(i)
      Else
         Assert.AreEqual Arr(i).Address, Arr(i).Address
      End If
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'SetObjectArrayToNothing'
'==============================================================================

'@TestMethod
Public Sub SetObjectArrayToNothing_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.SetObjectArrayToNothing(Scalar)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_UnallocatedLongArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.SetObjectArrayToNothing(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_UnallocatedObjectArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.SetObjectArrayToNothing(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_UnallocatedVariantArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As Variant


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.SetObjectArrayToNothing(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_1DLongArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.SetObjectArrayToNothing(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_1DObjectArr_ReturnsTrueAndNothingArr()
   On Error GoTo TestFail

   Dim Arr(5 To 7) As Object
   Dim Element As Variant


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = .Range("B5")
      Set Arr(6) = Nothing
      Set Arr(7) = .Range("B7")
   End With
   
   'Act:
   If Not modArraySupport.SetObjectArrayToNothing(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For Each Element In Arr
      Assert.IsNothing Element
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_1DVariantArr_ReturnsTrueAndNothingArr()
   On Error GoTo TestFail

   Dim Arr(5 To 7) As Variant
   Dim Element As Variant


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = .Range("B5")
      Set Arr(6) = Nothing
      Set Arr(7) = .Range("B7")
   End With
   
   'Act:
   If Not modArraySupport.SetObjectArrayToNothing(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For Each Element In Arr
      Assert.IsNothing Element
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_1DVariantArrWithEmptyElement_ReturnsFalse()
   On Error GoTo TestFail

   Dim Arr(5 To 7) As Variant


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5) = .Range("B5")
      Set Arr(6) = Nothing
      Arr(7) = Empty
   End With
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.SetObjectArrayToNothing(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_2DObjectArr_ReturnsTrueAndNothingArr()
   On Error GoTo TestFail

   Dim Arr(5 To 7, 3 To 4) As Object
   Dim Element As Variant


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5, 3) = .Range("B5")
      Set Arr(6, 3) = Nothing
      Set Arr(7, 3) = .Range("B7")
   
      Set Arr(5, 4) = .Range("B9")
      Set Arr(6, 4) = Nothing
      Set Arr(7, 4) = .Range("B11")
   End With
   
   'Act:
   If Not modArraySupport.SetObjectArrayToNothing(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For Each Element In Arr
      Assert.IsNothing Element
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_3DObjectArr_ReturnsTrueAndNothingArr()
   On Error GoTo TestFail

   Dim Arr(5 To 7, 3 To 4, 2 To 2) As Object
   Dim Element As Variant


   'Arrange:
   With ThisWorkbook.Worksheets(1)
      Set Arr(5, 3, 2) = .Range("B5")
      Set Arr(6, 3, 2) = Nothing
      Set Arr(7, 3, 2) = .Range("B7")
   
      Set Arr(5, 4, 2) = .Range("B9")
      Set Arr(6, 4, 2) = Nothing
      Set Arr(7, 4, 2) = .Range("B11")
   End With
   
   'Act:
   If Not modArraySupport.SetObjectArrayToNothing(Arr) Then _
         GoTo TestFail
   
   'Assert:
   For Each Element In Arr
      Assert.IsNothing Element
   Next

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SetObjectArrayToNothing_4DObjectArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 7, 3 To 4, 2 To 2, 1 To 1) As Object


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.SetObjectArrayToNothing(Arr)

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'SetVariableToDefault'
'==============================================================================

'all tests are done in the unit tests for function 'ResetVariantArrayToDefaults'


'==============================================================================
'unit tests for 'SwapArrayColumns'
'==============================================================================

'@TestMethod
Public Sub SwapArrayColumns_NoArray_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 4
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Scalar, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_UnallocatedArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 4
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_1DArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 4
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_3DArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4, 2 To 2) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 4
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_TooSmallCol1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 2
   Const Col2 As LongPtr = 4
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_TooSmallCol2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 2
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_TooLargeCol1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 5
   Const Col2 As LongPtr = 4
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_TooLargeCol2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 5
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_EqualColNumbers_ReturnsResultArrEqualToArr()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 3
   
   Dim aExpected(5 To 6, 3 To 4) As LongPtr
      aExpected(5, 3) = 10
      aExpected(6, 3) = 11
      aExpected(5, 4) = 20
      aExpected(6, 4) = 21
   '===========================================================================
   
   
   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   ResultArr = modArraySupport.SwapArrayColumns( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayColumns_UnequalColNumbers_ReturnsResultArrWithSwappedColumns()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Col1 As LongPtr = 3
   Const Col2 As LongPtr = 4
   
   Dim aExpected(5 To 6, 3 To 4) As LongPtr
      aExpected(5, 3) = 20
      aExpected(6, 3) = 21
      aExpected(5, 4) = 10
      aExpected(6, 4) = 11
   '===========================================================================
   
   
   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   ResultArr = modArraySupport.SwapArrayColumns( _
         Arr, _
         Col1, _
         Col2 _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'==============================================================================
'unit tests for 'SwapArrayRows'
'==============================================================================

'@TestMethod
Public Sub SwapArrayRows_NoArray_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 6
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Scalar, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_UnallocatedArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr() As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 6
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_1DArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 6
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_3DArr_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4, 2 To 2) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 6
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_TooSmallRow1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 4
   Const Row2 As LongPtr = 6
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_TooSmallRow2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 4
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_TooLargeRow1_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 7
   Const Row2 As LongPtr = 6
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_TooLargeRow2_ReturnsNull()
   On Error GoTo TestFail

   'Arrange:
   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 7
   
   Const aExpected As Variant = Null
   '===========================================================================
   
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.AreEqual aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_EqualRowNumbers_ReturnsResultArrEqualToArr()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 5
   
   Dim aExpected(5 To 6, 3 To 4) As LongPtr
      aExpected(5, 3) = 10
      aExpected(6, 3) = 11
      aExpected(5, 4) = 20
      aExpected(6, 4) = 21
   '===========================================================================
   
   
   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub SwapArrayRows_UnequalRowNumbers_ReturnsResultArrWithSwappedRows()
   On Error GoTo TestFail

   Dim Arr(5 To 6, 3 To 4) As LongPtr
   Dim ResultArr As Variant
   
   '===========================================================================
   Const Row1 As LongPtr = 5
   Const Row2 As LongPtr = 6
   
   Dim aExpected(5 To 6, 3 To 4) As LongPtr
      aExpected(5, 3) = 11
      aExpected(6, 3) = 10
      aExpected(5, 4) = 21
      aExpected(6, 4) = 20
   '===========================================================================
   
   
   'Arrange:
   Arr(5, 3) = 10
   Arr(6, 3) = 11
   Arr(5, 4) = 20
   Arr(6, 4) = 21
   
   'Act:
   ResultArr = modArraySupport.SwapArrayRows( _
         Arr, _
         Row1, _
         Row2 _
   )
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
'unit tests for 'VectorsToArray'
'==============================================================================

'@TestMethod
Public Sub VectorsToArray_NoArray_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim Scalar As LongPtr
   Dim VectorA(5 To 7) As LongPtr
   Dim VectorB(4 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         Scalar, _
         VectorA, _
         VectorB _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_StaticArr_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArr(0 To 2) As LongPtr
   Dim VectorA(5 To 7) As LongPtr
   Dim VectorB(4 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         ResultArr, _
         VectorA, _
         VectorB _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_MissingVectors_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArr() As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         ResultArr _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_ScalarVector_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArr() As LongPtr
   Dim ScalarA As LongPtr
   Dim VectorB(4 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         ResultArr, _
         ScalarA, _
         VectorB _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_UninitiallizedVector_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArr() As LongPtr
   Dim ArrayA() As LongPtr
   Dim VectorB(4 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         ResultArr, _
         ArrayA, _
         VectorB _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_2DVector_ReturnsFalse()
   On Error GoTo TestFail

   'Arrange:
   Dim ResultArr() As LongPtr
   Dim ArrayA(5 To 7, 3 To 4) As LongPtr
   Dim VectorB(4 To 6) As LongPtr


   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         ResultArr, _
         ArrayA, _
         VectorB _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_ArrayInVector_ReturnsFalse()
   On Error GoTo TestFail

   Dim ResultArr() As Variant
   Dim VectorA(5 To 7) As Variant
   Dim VectorB(4 To 6) As LongPtr


   'Arrange:
   VectorA(5) = Array(5, 6, 7)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         ResultArr, _
         VectorA, _
         VectorB _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_ObjectInVector_ReturnsFalse()
   On Error GoTo TestFail

   Dim ResultArr() As Variant
   Dim VectorA(5 To 7) As Variant
   Dim VectorB(4 To 6) As LongPtr


   'Arrange:
   Set VectorA(5) = ThisWorkbook.Worksheets(1)
   
   'Act:
   'Assert:
   Assert.IsFalse modArraySupport.VectorsToArray( _
         ResultArr, _
         VectorA, _
         VectorB _
   )

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub VectorsToArray_ValidLongVectors_ReturnsTrueAndResultArr()
   On Error GoTo TestFail

   Dim ResultArr() As LongPtr
   Dim VectorA(5 To 7) As LongPtr
   Dim VectorB(4 To 6) As LongPtr

   '===========================================================================
   Dim aExpected(0 To 2, 0 To 1) As LongPtr
      aExpected(0, 0) = 10
      aExpected(1, 0) = 11
      aExpected(2, 0) = 12
      aExpected(0, 1) = 20
      aExpected(1, 1) = 21
      aExpected(2, 1) = 22
   '===========================================================================

   'Arrange:
   VectorA(5) = 10
   VectorA(6) = 11
   VectorA(7) = 12
   
   VectorB(4) = 20
   VectorB(5) = 21
   VectorB(6) = 22
   
   'Act:
   If Not modArraySupport.VectorsToArray( _
         ResultArr, _
         VectorA, _
         VectorB _
   ) Then _
         GoTo TestFail
   
   'Assert:
   Assert.SequenceEquals aExpected, ResultArr

TestExit:
   Exit Sub
TestFail:
   Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub