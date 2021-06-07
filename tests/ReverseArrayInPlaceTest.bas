Attribute VB_Name = "ReverseArrayInPlaceTest"

'@TestModule
'@Folder("modArraySupport2.Tests")

Option Explicit
Option Compare Text
Option Private Module

'change value from 'LateBindTests' to '1' for late bound tests
'alternatively add
'    LateBindTests = 1
'to Tools > <project name> Properties > General > Conditional Compilation Arguments
'to make it work for *all* test modules in the project
#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
#If LateBind Then
    Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
#Else
    Set Assert = New Rubberduck.PermissiveAssertClass
#End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
End Sub


'==============================================================================
'unit tests for 'ReverseArrayInPlace'
'==============================================================================

'@TestMethod("ReverseArrayInPlace")
Public Sub ReverseVectorInPlace_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ReverseVectorInPlace(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseArrayInPlace")
Public Sub ReverseVectorInPlace_UnallocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ReverseVectorInPlace(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseArrayInPlace")
Public Sub ReverseVectorInPlace_2DArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ReverseVectorInPlace(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseArrayInPlace")
Public Sub ReverseVectorInPlace_ValidEven1DArr_ReturnsTrueAndReversedArr()
    On Error GoTo TestFail

    Dim Arr(5 To 8) As Long
    
    '==========================================================================
    Dim aExpected(5 To 8) As Long
        aExpected(5) = 8
        aExpected(6) = 7
        aExpected(7) = 6
        aExpected(8) = 5
    '==========================================================================
    
    
    'Arrange:
    Arr(5) = 5
    Arr(6) = 6
    Arr(7) = 7
    Arr(8) = 8
    
    'Act:
    If Not modArraySupport2.ReverseVectorInPlace(Arr) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Arr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseArrayInPlace")
Public Sub ReverseVectorInPlace_ValidEven1DVariantArr_ReturnsTrueAndReversedArr()
    On Error GoTo TestFail

    Dim Arr(5 To 8) As Variant
    
    '==========================================================================
    Dim aExpected(5 To 8) As Variant
        aExpected(5) = 8
        aExpected(6) = "ghi"
        aExpected(7) = 6
        aExpected(8) = "abc"
    '==========================================================================
    
    
    'Arrange:
    Arr(5) = "abc"
    Arr(6) = 6
    Arr(7) = "ghi"
    Arr(8) = 8
    
    'Act:
    If Not modArraySupport2.ReverseVectorInPlace(Arr) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Arr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseArrayInPlace")
Public Sub ReverseVectorInPlace_1DVariantArrWithObject_ReturnsTrueAndReversedArr()
    On Error GoTo TestFail

    Dim Arr(5 To 6) As Variant
    
    '==========================================================================
    Dim aExpected(5 To 6) As Variant
        aExpected(5) = "AreDataTypesCompatible"   '*content* of the below cell
        aExpected(6) = 5
    '==========================================================================
    
    
    'Arrange:
    Arr(5) = 5
    Set Arr(6) = ThisWorkbook.Worksheets(1).Range("B5")
    
    'Act:
    If Not modArraySupport2.ReverseVectorInPlace(Arr) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Arr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseArrayInPlace")
Public Sub ReverseVectorInPlace_ValidOdd1DArr_ReturnsTrueAndReversedArr()
    On Error GoTo TestFail

    Dim Arr(5 To 9) As Long
    
    '==========================================================================
    Dim aExpected(5 To 9) As Long
        aExpected(5) = 9
        aExpected(6) = 8
        aExpected(7) = 7
        aExpected(8) = 6
        aExpected(9) = 5
    '==========================================================================
    
    
    'Arrange:
    Arr(5) = 5
    Arr(6) = 6
    Arr(7) = 7
    Arr(8) = 8
    Arr(9) = 9
    
    'Act:
    If Not modArraySupport2.ReverseVectorInPlace(Arr) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Arr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
