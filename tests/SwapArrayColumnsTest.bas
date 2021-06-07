Attribute VB_Name = "SwapArrayColumnsTest"

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
'unit tests for 'SwapArrayColumns'
'==============================================================================

'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_NoArray_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 4
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Scalar, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_UnallocatedArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 4
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_1DArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 4
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_3DArr_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4, 2 To 2) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 4
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_TooSmallCol1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 2
    Const Col2 As Long = 4
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_TooSmallCol2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 2
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_TooLargeCol1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 5
    Const Col2 As Long = 4
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_TooLargeCol2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 5
    '==========================================================================
    
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayRows( _
            Arr, _
            Col1, _
            Col2 _
    )
    
    'Assert:
    Assert.IsTrue IsNull(ResultArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_EqualColNumbers_ReturnsResultArrEqualToArr()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 3
    
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
    ResultArr = modArraySupport2.SwapArrayColumns( _
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


'@TestMethod("SwapArrayColumns")
Public Sub SwapArrayColumns_UnequalColNumbers_ReturnsResultArrWithSwappedColumns()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr As Variant
    
    '==========================================================================
    Const Col1 As Long = 3
    Const Col2 As Long = 4
    
    Dim aExpected(5 To 6, 3 To 4) As Long
        aExpected(5, 3) = 20
        aExpected(6, 3) = 21
        aExpected(5, 4) = 10
        aExpected(6, 4) = 11
    '==========================================================================
    
    
    'Arrange:
    Arr(5, 3) = 10
    Arr(6, 3) = 11
    Arr(5, 4) = 20
    Arr(6, 4) = 21
    
    'Act:
    ResultArr = modArraySupport2.SwapArrayColumns( _
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
