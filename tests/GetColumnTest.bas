Attribute VB_Name = "GetColumnTest"

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
'unit tests for 'GetColumn'
'==============================================================================

'@TestMethod("GetColumn")
Public Sub GetColumn_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    Dim ResultArr() As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 4
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.GetColumn( _
            Scalar, _
            ResultArr, _
            ColumnNumber _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("GetColumn")
Public Sub GetColumn_1DArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Long
    Dim ResultArr() As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 4
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.GetColumn( _
            Arr, _
            ResultArr, _
            ColumnNumber _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("GetColumn")
Public Sub GetColumn_3DArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4, -1 To 0) As Long
    Dim ResultArr() As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 4
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.GetColumn( _
            Arr, _
            ResultArr, _
            ColumnNumber _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("GetColumn")
Public Sub GetColumn_StaticResultArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr(-5 To -4) As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 4
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.GetColumn( _
            Arr, _
            ResultArr, _
            ColumnNumber _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("GetColumn")
Public Sub GetColumn_TooSmallColumnNumber_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr() As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 2
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.GetColumn( _
            Arr, _
            ResultArr, _
            ColumnNumber _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("GetColumn")
Public Sub GetColumn_TooLargeColumnNumber_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr() As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 5
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.GetColumn( _
            Arr, _
            ResultArr, _
            ColumnNumber _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("GetColumn")
Public Sub GetColumn_LegalEntries_ReturnsTrueAndResultArr()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Long
    Dim ResultArr() As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 4
    
    Dim aExpected(5 To 6) As Long
        aExpected(5) = 20
        aExpected(6) = 21
    '==========================================================================
    
    
    'Arrange:
    Arr(5, 3) = 10
    Arr(6, 3) = 11
    Arr(5, 4) = 20
    Arr(6, 4) = 21
    
    'Act:
    If Not modArraySupport2.GetColumn( _
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


'@TestMethod("GetColumn")
Public Sub GetColumn_LegalEntriesWithObjects_ReturnsTrueAndResultArr()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Variant
    Dim ResultArr() As Variant
    Dim i As Long
    
    '==========================================================================
    Const ColumnNumber As Long = 4
    
    Dim aExpected(5 To 6) As Variant
    With ThisWorkbook.Worksheets(1)
        aExpected(5) = vbNullString
        Set aExpected(6) = .Range("B5")
    End With
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Arr(5, 3) = 10
        Arr(6, 3) = 11
        Arr(5, 4) = vbNullString
        Set Arr(6, 4) = .Range("B5")
    End With
    
    'Act:
    If Not modArraySupport2.GetColumn( _
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
