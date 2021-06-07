Attribute VB_Name = "TransposeArrayTest"

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
'unit tests for 'TransposeArray'
'==============================================================================

'@TestMethod("TransposeArray")
Public Sub TransposeArray_ScalarInput_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Const Scalar As Long = 5
    Dim TransposedArr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.TransposeArray(Scalar, TransposedArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("TransposeArray")
Public Sub TransposeArray_1DInputArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(2) As Long
    Dim TransposedArr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.TransposeArray(Arr, TransposedArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("TransposeArray")
Public Sub TransposeArray_ScalarOutput_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 3, 2 To 5) As Long
    Dim Scalar As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.TransposeArray(Arr, Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("TransposeArray")
Public Sub TransposeArray_StaticOutputArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 3, 2 To 5) As Long
    Dim TransposedArr(2 To 5, 1 To 3) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.TransposeArray(Arr, TransposedArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("TransposeArray")
Public Sub TransposeArray_Valid2DArr_ReturnsTrueAndTransposedArr()
    On Error GoTo TestFail

    Dim Arr() As Long
    Dim TransposedArr() As Long
    Dim i As Long
    Dim j As Long
    
    
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
    If Not modArraySupport2.TransposeArray(Arr, TransposedArr) _
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
