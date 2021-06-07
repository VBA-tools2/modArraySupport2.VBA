Attribute VB_Name = "ConcatenateVectorsTest"

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
'unit tests for 'ConcatenateVectors'
'==============================================================================

'@TestMethod("ConcatenateVectors")
Public Sub ConcatenateVectors_StaticResultArray_ResultsFalse()
    On Error GoTo TestFail

    Dim ResultArray(1) As Long
    Dim ArrayToAppend(1) As Long
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    ResultArray(1) = 8
    ArrayToAppend(1) = 111
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ConcatenateVectors( _
            ResultArray, _
            ArrayToAppend, _
            CompatibilityCheck _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConcatenateVectors")
Public Sub ConcatenateVectors_BothArraysUnallocated_ResultsTrueAndUnallocatedArray()
    On Error GoTo TestFail

    'Arrange:
    Dim ResultArray() As Long
    Dim ArrayToAppend() As Long
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    '==========================================================================
    
    
    'Act:
    If Not modArraySupport2.ConcatenateVectors( _
            ResultArray, _
            ArrayToAppend, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.IsFalse IsArrayAllocated(ResultArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConcatenateVectors")
Public Sub ConcatenateVectors_UnallocatedArrayToAppend_ResultsTrueAndUnchangedResultArray()
    On Error GoTo TestFail

    Dim ResultArray() As Long
    Dim ArrayToAppend() As Long
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    
    Dim aExpected(1 To 2) As Long
        aExpected(1) = 8
        aExpected(2) = 9
    '==========================================================================
    
    
    'Arrange:
    ReDim ResultArray(1 To 2)
    ResultArray(1) = 8
    ResultArray(2) = 9
    
    'Act:
    If Not modArraySupport2.ConcatenateVectors( _
            ResultArray, _
            ArrayToAppend, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConcatenateVectors")
Public Sub ConcatenateVectors_IntegerArrayToAppendLongResultArray_ResultsTrueAndResultArray()
    On Error GoTo TestFail

    Dim ResultArray() As Long
    Dim ArrayToAppend(1 To 3) As Integer
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    
    Dim aExpected(1 To 6) As Long
        aExpected(1) = 8
        aExpected(2) = 9
        aExpected(3) = 10
        aExpected(4) = 111
        aExpected(5) = 112
        aExpected(6) = 113
    '==========================================================================
    
    
    'Arrange:
    ReDim ResultArray(1 To 3)
    ResultArray(1) = 8
    ResultArray(2) = 9
    ResultArray(3) = 10
    
    ArrayToAppend(1) = 111
    ArrayToAppend(2) = 112
    ArrayToAppend(3) = 113
    
    'Act:
    If Not modArraySupport2.ConcatenateVectors( _
            ResultArray, _
            ArrayToAppend, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConcatenateVectors")
Public Sub ConcatenateVectors_LongArrayToAppendIntegerResultArray_ResultsFalse()
    On Error GoTo TestFail

    Dim ResultArray() As Integer
    Dim ArrayToAppend(1 To 3) As Long
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    '==========================================================================
    
    
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
    Assert.IsFalse modArraySupport2.ConcatenateVectors( _
            ResultArray, _
            ArrayToAppend, _
            CompatibilityCheck _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConcatenateVectors")
Public Sub ConcatenateVectors_LongArrayToAppendIntegerResultArrayFalseCompatibilityCheck_ResultsTrueAndResultArray()
    On Error GoTo TestFail

    Dim ResultArray() As Integer
    Dim ArrayToAppend(1 To 3) As Long
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = False
    
    Dim aExpected(1 To 6) As Integer
        aExpected(1) = 8
        aExpected(2) = 9
        aExpected(3) = 10
        aExpected(4) = 111
        aExpected(5) = 112
        aExpected(6) = 113
    '==========================================================================
    
    
    'Arrange:
    ReDim ResultArray(1 To 3)
    ResultArray(1) = 8
    ResultArray(2) = 9
    ResultArray(3) = 10
    
    ArrayToAppend(1) = 111
    ArrayToAppend(2) = 112
    ArrayToAppend(3) = 113
    
    'Act:
    If Not modArraySupport2.ConcatenateVectors( _
            ResultArray, _
            ArrayToAppend, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConcatenateVectors")
Public Sub ConcatenateVectors_LongArrayToAppendWithLongNumberIntegerResultArrayFalseCompatibilityCheck_ResultsFalse()
    On Error GoTo TestFail

    Dim ResultArray() As Integer
    Dim ArrayToAppend(1 To 3) As Long
    Dim Success As Boolean
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = False
    
    Const ExpectedError As Long = 6
    '==========================================================================
    
    
    'Arrange:
    ReDim ResultArray(1 To 3)
    ResultArray(1) = 8
    ResultArray(2) = 9
    ResultArray(3) = 10
    
    ArrayToAppend(1) = 111
    ArrayToAppend(2) = 32768   'no valid Integer
    ArrayToAppend(3) = 113
    
    'Act:
    Success = modArraySupport2.ConcatenateVectors( _
            ResultArray, _
            ArrayToAppend, _
            CompatibilityCheck _
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


''TODO: add a test that involves objects
''     (have a look at <https://stackoverflow.com/a/11254505>
''@TestMethod("ConcatenateVectors")
'Public Sub ConcatenateVectors_LegalVariant_ResultsTrueAndResultArray()
'    On Error GoTo TestFail
'
'    Dim ResultArray() As Range          'MUST be dynamic
'    Dim ArrayToAppend(0 To 0) As Range
'    Dim i As Long
'
'    '=========================================================================
'    Const CompatibilityCheck As Boolean = True
'
'    Dim wks As Worksheet
'    Set wks = tblFunctions
'    Dim aExpected(1 To 2) As Range
'    With wks
'        Set aExpected(1) = .Cells(1, 1)
'        Set aExpected(2) = .Cells(1, 2)
'    End With
'    '=========================================================================
'
'
'    'Arrange:
'    With wks
'        ReDim ResultArray(1 To 1)
'        Set ResultArray(1) = .Cells(1, 1)
'        Set ArrayToAppend(0) = .Cells(1, 2)
'    End With
'
'    'Act:
'    If Not modArraySupport2.ConcatenateVectors( _
'            ResultArray, _
'            ArrayToAppend, _
'            CompatibilityCheck _
'    ) Then _
'            GoTo TestFail
'
'    'Assert:
'    For i = LBound(ResultArray) To UBound(ResultArray)
'Debug.Print aExpected(i) Is ResultArray(i)
'        Assert.AreSame aExpected(i), ResultArray(i)
'    Next
'
''    If B = True Then
''        If modArraySupport2.IsArrayAllocated(ResultArray) = True Then
''            For i = LBound(ResultArray) To UBound(ResultArray)
''                If IsObject(ResultArray(i)) = True Then
''Debug.Print CStr(i), "is object", TypeName(ResultArray(i))
''                Else
''Debug.Print CStr(i), ResultArray(i)
''                End If
''            Next
''        Else
''Debug.Print "Result Array Is Not Allocated."
''        End If
''    Else
''Debug.Print "ConcatenateVectors returned False"
''    End If
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub
