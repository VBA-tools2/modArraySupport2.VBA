Attribute VB_Name = "VectorsToArrayTest"

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
'unit tests for 'VectorsToArray'
'==============================================================================

'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    Dim VectorA(5 To 7) As Long
    Dim VectorB(4 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            Scalar, _
            VectorA, _
            VectorB _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_StaticArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim ResultArr(0 To 2) As Long
    Dim VectorA(5 To 7) As Long
    Dim VectorB(4 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            ResultArr, _
            VectorA, _
            VectorB _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_MissingVectors_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim ResultArr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            ResultArr _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_ScalarVector_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim ResultArr() As Long
    Dim ScalarA As Long
    Dim VectorB(4 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            ResultArr, _
            ScalarA, _
            VectorB _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_UninitializedVector_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim ResultArr() As Long
    Dim ArrayA() As Long
    Dim VectorB(4 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            ResultArr, _
            ArrayA, _
            VectorB _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_2DVector_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim ResultArr() As Long
    Dim ArrayA(5 To 7, 3 To 4) As Long
    Dim VectorB(4 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            ResultArr, _
            ArrayA, _
            VectorB _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_ArrayInVector_ReturnsFalse()
    On Error GoTo TestFail

    Dim ResultArr() As Variant
    Dim VectorA(5 To 7) As Variant
    Dim VectorB(4 To 6) As Long
    
    
    'Arrange:
    VectorA(5) = Array(5, 6, 7)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            ResultArr, _
            VectorA, _
            VectorB _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_ObjectInVector_ReturnsFalse()
    On Error GoTo TestFail

    Dim ResultArr() As Variant
    Dim VectorA(5 To 7) As Variant
    Dim VectorB(4 To 6) As Long
    
    
    'Arrange:
    Set VectorA(5) = ThisWorkbook.Worksheets(1)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.VectorsToArray( _
            ResultArr, _
            VectorA, _
            VectorB _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("VectorsToArray")
Public Sub VectorsToArray_ValidLongVectors_ReturnsTrueAndResultArr()
    On Error GoTo TestFail

    Dim ResultArr() As Long
    Dim VectorA(5 To 7) As Long
    Dim VectorB(4 To 6) As Long
    
    '==========================================================================
    Dim aExpected(0 To 2, 0 To 1) As Long
        aExpected(0, 0) = 10
        aExpected(1, 0) = 11
        aExpected(2, 0) = 12
        aExpected(0, 1) = 20
        aExpected(1, 1) = 21
        aExpected(2, 1) = 22
    '==========================================================================
    
    'Arrange:
    VectorA(5) = 10
    VectorA(6) = 11
    VectorA(7) = 12
    
    VectorB(4) = 20
    VectorB(5) = 21
    VectorB(6) = 22
    
    'Act:
    If Not modArraySupport2.VectorsToArray( _
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
