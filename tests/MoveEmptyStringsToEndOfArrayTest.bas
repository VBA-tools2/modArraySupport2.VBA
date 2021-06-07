Attribute VB_Name = "MoveEmptyStringsToEndOfArrayTest"

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
'unit tests for 'MoveEmptyStringsToEndOfArray'
'==============================================================================

'@TestMethod("MoveEmptyStringsToEndOfArray")
Public Sub MoveEmptyStringsToEndOfVector_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As String
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.MoveEmptyStringsToEndOfVector(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MoveEmptyStringsToEndOfArray")
Public Sub MoveEmptyStringsToEndOfVector_UnallocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray() As String
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MoveEmptyStringsToEndOfArray")
Public Sub MoveEmptyStringsToEndOfVector_2DArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6, 3 To 4) As String
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MoveEmptyStringsToEndOfArray")
Public Sub MoveEmptyStringsToEndOfVector_vbNullStringArrayOnly_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As String
    
    
    'Arrange:
    InputArray(5) = vbNullString
    InputArray(6) = vbNullString
    InputArray(7) = vbNullString
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MoveEmptyStringsToEndOfArray")
Public Sub MoveEmptyStringsToEndOfVector_NoneVbNullStringArrayOnly_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As String
    
    
    'Arrange:
    InputArray(5) = "abc"
    InputArray(6) = "def"
    InputArray(7) = "ghi"
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.MoveEmptyStringsToEndOfVector(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MoveEmptyStringsToEndOfArray")
Public Sub MoveEmptyStringsToEndOfVector_StringArray_ReturnsTrueAndModifiedArr()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As String
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(5 To 7) As String
        aExpected(5) = "abc"
        aExpected(6) = vbNullString
        aExpected(7) = vbNullString
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = vbNullString
    InputArray(6) = vbNullString
    InputArray(7) = "abc"
    
    'Act:
    If Not modArraySupport2.MoveEmptyStringsToEndOfVector(InputArray) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(InputArray) To UBound(InputArray)
        Assert.AreEqual aExpected(i), InputArray(i)
    Next
'    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("MoveEmptyStringsToEndOfArray")
Public Sub MoveEmptyStringsToEndOfVector_VariantArray_ReturnsTrueAndModifiedArr()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As Variant
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(5 To 7) As Variant
        aExpected(5) = "abc"
        aExpected(6) = "def"
        aExpected(7) = vbNullString
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = vbNullString
    InputArray(6) = "abc"
    InputArray(7) = "def"
    
    'Act:
    If Not modArraySupport2.MoveEmptyStringsToEndOfVector(InputArray) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(InputArray) To UBound(InputArray)
        Assert.AreEqual aExpected(i), InputArray(i)
    Next
'    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''@TestMethod("MoveEmptyStringsToEndOfArray")
'Public Sub MoveEmptyStringsToEndOfVector_StringArray2_ReturnsTrueAndModifiedArr()
'    On Error GoTo TestFail
'
'    Dim Arr As Variant
'    Dim InputArray() As String
'    Dim i As Long
'
'    '==========================================================================
'    Dim aExpected() As String
'    '==========================================================================
'
'
'    'Arrange:
''move entries in the shown range 3 cells down
'    Arr = ThisWorkbook.Worksheets(1).Range("B32:B44")
'
'    'Act:
'     If Not modArraySupport2.GetColumn(Arr, InputArray, 1) Then GoTo TestFail
'    If Not modArraySupport2.MoveEmptyStringsToEndOfVector(InputArray) Then _
'           GoTo TestFail
'    Arr = ThisWorkbook.Worksheets(1).Range("B35:B47")
'    If Not modArraySupport2.GetColumn(Arr, aExpected, 1) Then GoTo TestFail
'
'    'Assert:
'    For i = LBound(InputArray) To UBound(InputArray)
'        Assert.AreEqual aExpected(i), InputArray(i)
'    Next
''    Assert.SequenceEquals aExpected, InputArray
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub
