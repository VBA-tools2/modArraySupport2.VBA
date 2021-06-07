Attribute VB_Name = "InsertElementIntoVectorTest"

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
'unit tests for 'InsertElementIntoVector'
'==============================================================================

'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_StaticInputArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6) As Long
    
    '==========================================================================
    Const Index As Long = 6
    Const Value As Long = 33
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_2DInputArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    
    '==========================================================================
    Const Index As Long = 6
    Const Value As Long = 33
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 6, 3 To 4)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_TooSmallIndex_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    
    '==========================================================================
    Const Index As Long = 4
    Const Value As Long = 33
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 6)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_TooLargeIndex_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    
    '==========================================================================
    Const Index As Long = 8
    Const Value As Long = 33
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 6)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_WrongValueType_ReturnsFalseAndUnchangedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As Long
    
    '==========================================================================
    Const Index As Long = 6
    Const Value As String = "abc"
    
    Dim aExpected(5 To 6) As Long
        aExpected(5) = 10
        aExpected(6) = 11
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 6)
    InputArray(5) = 10
    InputArray(6) = 11
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
    )
    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_ValidTestWithLongs_ReturnsTrueAndChangedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As Long
    
    '==========================================================================
    Const Index As Long = 6
    Const Value As Long = 33
    
    Dim aExpected(5 To 7) As Long
        aExpected(5) = 10
        aExpected(6) = 33
        aExpected(7) = 11
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 6)
    InputArray(5) = 10
    InputArray(6) = 11
    
    'Act:
    If Not modArraySupport2.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_ValidTestWithStrings_ReturnsTrueAndChangedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As String
    Dim i As Long
    
    '==========================================================================
    Const Index As Long = 7
    Const Value As String = "XYZ"
    
    Dim aExpected(5 To 7) As String
        aExpected(5) = "abc"
        aExpected(6) = vbNullString
        aExpected(7) = "XYZ"
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 6)
    InputArray(5) = "abc"
    InputArray(6) = vbNullString
    
    'Act:
    If Not modArraySupport2.InsertElementIntoVector( _
            InputArray, _
            Index, _
            Value _
    ) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(InputArray) To UBound(InputArray)
        Assert.AreEqual aExpected(i), InputArray(i)
    Next
'TODO: why does the following line result in an error?
'   Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("InsertElementIntoVector")
Public Sub InsertElementIntoVector_ValidTestWithObjects_ReturnsTrueAndChangedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As Object
    Dim wks As Worksheet
        Set wks = ThisWorkbook.Worksheets(1)
    Dim i As Long
    
    
    With wks
        
        '======================================================================
        Const Index As Long = 6
        Dim Value As Object
            Set Value = .Range("B2")
        
        Dim aExpected(5 To 7) As Object
            Set aExpected(5) = .Range("B5")
            Set aExpected(6) = .Range("B2")
            Set aExpected(7) = Nothing
        '======================================================================
        
        
        'Arrange:
        ReDim InputArray(5 To 6)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = Nothing
        
        'Act:
        If Not modArraySupport2.InsertElementIntoVector( _
                InputArray, _
                Index, _
                Value _
        ) Then _
                GoTo TestFail
        
        'Assert:
        For i = LBound(InputArray) To UBound(InputArray)
            If InputArray(i) Is Nothing Then
                Assert.IsNothing aExpected(i)
            Else
                Assert.AreEqual aExpected(i).Address, InputArray(i).Address
            End If
        Next
    End With

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
