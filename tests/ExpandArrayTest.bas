Attribute VB_Name = "ExpandArrayTest"

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
'unit tests for 'ExpandArray'
'==============================================================================

'@TestMethod("ExpandArray")
Public Sub ExpandArray_NoArray_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 1
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 11
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
            Arr, _
            WhichDim, _
            AdditionalElements, _
            FillValue _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpandArray")
Public Sub ExpandArray_UnallocatedArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 1
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 11
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
            Arr, _
            WhichDim, _
            AdditionalElements, _
            FillValue _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpandArray")
Public Sub ExpandArray_1DArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 1
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 11
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
            Arr, _
            WhichDim, _
            AdditionalElements, _
            FillValue _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpandArray")
Public Sub ExpandArray_3DArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4, 2 To 3) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 1
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 11
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
            Arr, _
            WhichDim, _
            AdditionalElements, _
            FillValue _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpandArray")
Public Sub ExpandArray_WhichDimSmallerOne_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 0
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 11
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
            Arr, _
            WhichDim, _
            AdditionalElements, _
            FillValue _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpandArray")
Public Sub ExpandArray_WhichDimLargerTwo_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 3
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 11
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
            Arr, _
            WhichDim, _
            AdditionalElements, _
            FillValue _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpandArray")
Public Sub ExpandArray_AdditionalElementsSmallerZero_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 1
    Const AdditionalElements As Long = -1
    Const FillValue As Long = 11
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
            Arr, _
            WhichDim, _
            AdditionalElements, _
            FillValue _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpandArray")
Public Sub ExpandArray_AdditionalElementsEqualsZero_ReturnsExpandedArray()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 1
    Const AdditionalElements As Long = 0
    Const FillValue As Long = 33
    
    Dim aExpected(5 To 6, 3 To 4) As Long
        aExpected(5, 3) = 10
        aExpected(6, 3) = 11
        aExpected(5, 4) = 20
        aExpected(6, 4) = 21
    '==========================================================================
    
    
    'Arrange:
    Arr(5, 3) = 10
    Arr(6, 3) = 11
    Arr(5, 4) = 20
    Arr(6, 4) = 21
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
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


'@TestMethod("ExpandArray")
Public Sub ExpandArray_AddTwoAdditionalRows_ReturnsExpandedArray()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 1
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 33
    
    Dim aExpected(5 To 8, 3 To 4) As Long
        aExpected(5, 3) = 10
        aExpected(6, 3) = 11
        aExpected(5, 4) = 20
        aExpected(6, 4) = 21
        aExpected(7, 3) = 33
        aExpected(8, 3) = 33
        aExpected(7, 4) = 33
        aExpected(8, 4) = 33
    '==========================================================================
    
    
    'Arrange:
    Arr(5, 3) = 10
    Arr(6, 3) = 11
    Arr(5, 4) = 20
    Arr(6, 4) = 21
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
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


'@TestMethod("ExpandArray")
Public Sub ExpandArray_AddTwoAdditionalCols_ReturnsExpandedArray()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const WhichDim As Long = 2
    Const AdditionalElements As Long = 2
    Const FillValue As Long = 33
    
    Dim aExpected(5 To 6, 3 To 6) As Long
        aExpected(5, 3) = 10
        aExpected(6, 3) = 11
        aExpected(5, 4) = 20
        aExpected(6, 4) = 21
        aExpected(5, 5) = 33
        aExpected(6, 5) = 33
        aExpected(5, 6) = 33
        aExpected(6, 6) = 33
    '==========================================================================
    
    
    'Arrange:
    Arr(5, 3) = 10
    Arr(6, 3) = 11
    Arr(5, 4) = 20
    Arr(6, 4) = 21
    
    'Act:
    ResultArr = modArraySupport2.ExpandArray( _
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
