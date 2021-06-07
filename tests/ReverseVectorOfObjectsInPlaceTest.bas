Attribute VB_Name = "ReverseVectorOfObjectsInPlaceTest"

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
'unit tests for 'ReverseVectorOfObjectsInPlace'
'==============================================================================

'@TestMethod("ReverseVectorOfObjectsInPlace")
Public Sub ReverseVectorOfObjectsInPlace_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ReverseVectorOfObjectsInPlace(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseVectorOfObjectsInPlace")
Public Sub ReverseVectorOfObjectsInPlace_UnallocatedObjectArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ReverseVectorOfObjectsInPlace(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseVectorOfObjectsInPlace")
Public Sub ReverseVectorOfObjectsInPlace_2DObjectArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4) As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ReverseVectorOfObjectsInPlace(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseVectorOfObjectsInPlace")
Public Sub ReverseVectorOfObjectsInPlace_ValidEven1DObjectArr_ReturnsTrueAndReversedArr()
    On Error GoTo TestFail

    Dim Arr(5 To 8) As Object
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(5 To 8) As Object
    With ThisWorkbook.Worksheets(1)
        Set aExpected(5) = .Range("B8")
        Set aExpected(6) = .Range("B7")
        Set aExpected(7) = .Range("B6")
        Set aExpected(8) = .Range("B5")
    End With
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = .Range("B5")
        Set Arr(6) = .Range("B6")
        Set Arr(7) = .Range("B7")
        Set Arr(8) = .Range("B8")
    End With
    
    'Act:
    If Not modArraySupport2.ReverseVectorOfObjectsInPlace(Arr) Then _
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


'@TestMethod("ReverseVectorOfObjectsInPlace")
Public Sub ReverseVectorOfObjectsInPlace_ValidEven1DVariantArr_ReturnsTrueAndReversedArr()
    On Error GoTo TestFail

    Dim Arr(5 To 8) As Variant
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(5 To 8) As Variant
    With ThisWorkbook.Worksheets(1)
        Set aExpected(5) = .Range("B8")
        Set aExpected(6) = Nothing
        Set aExpected(7) = .Range("B6")
        Set aExpected(8) = Nothing
    End With
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = Nothing
        Set Arr(6) = .Range("B6")
        Set Arr(7) = Nothing
        Set Arr(8) = .Range("B8")
    End With
    
    'Act:
    If Not modArraySupport2.ReverseVectorOfObjectsInPlace(Arr) Then _
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


'@TestMethod("ReverseVectorOfObjectsInPlace")
Public Sub ReverseVectorOfObjectsInPlace_1DVariantArrWithNonObject_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(5 To 6) As Variant
    
    
    'Arrange:
    Set Arr(5) = ThisWorkbook.Worksheets(1).Range("B5")
    Arr(6) = 6
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ReverseVectorOfObjectsInPlace(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ReverseVectorOfObjectsInPlace")
Public Sub ReverseVectorOfObjectsInPlace_ValidOdd1DObjectArr_ReturnsTrueAndReversedArr()
    On Error GoTo TestFail

    Dim Arr(5 To 9) As Object
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(5 To 9) As Object
    With ThisWorkbook.Worksheets(1)
        Set aExpected(5) = .Range("B9")
        Set aExpected(6) = Nothing
        Set aExpected(7) = .Range("B7")
        Set aExpected(8) = .Range("B6")
        Set aExpected(9) = .Range("B5")
    End With
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = .Range("B5")
        Set Arr(6) = .Range("B6")
        Set Arr(7) = .Range("B7")
        Set Arr(8) = Nothing
        Set Arr(9) = .Range("B9")
    End With
    
    'Act:
    If Not modArraySupport2.ReverseVectorOfObjectsInPlace(Arr) Then _
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
