Attribute VB_Name = "SwapArrayRowsTest"

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
'unit tests for 'SwapArrayRows'
'==============================================================================

'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_NoArray_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 6
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Scalar, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_UnallocatedArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 6
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_1DArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 6
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_3DArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4, 2 To 2) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 6
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_TooSmallRow1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 4
    Const Row2 As Long = 6
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_TooSmallRow2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 4
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_TooLargeRow1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 7
    Const Row2 As Long = 6
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_TooLargeRow2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 7
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Row1, _
            Row2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_EqualRowNumbers_ReturnsResultArrEqualToArr()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 5
    
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
    ResultArr = modArraySupport2.SwapArrayRows( _
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


'@TestMethod("SwapArrayRows")
Public Sub SwapArrayRows_UnequalRowNumbers_ReturnsResultArrWithSwappedRows()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Row1 As Long = 5
    Const Row2 As Long = 6
    
    Dim aExpected(5 To 6, 3 To 4) As Long
        aExpected(5, 3) = 11
        aExpected(6, 3) = 10
        aExpected(5, 4) = 21
        aExpected(6, 4) = 20
    '==========================================================================
    
    
    'Arrange:
    Arr(5, 3) = 10
    Arr(6, 3) = 11
    Arr(5, 4) = 20
    Arr(6, 4) = 21
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
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
