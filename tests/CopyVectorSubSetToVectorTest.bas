Attribute VB_Name = "CopyVectorSubSetToVectorTest"

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
'unit tests for 'CopyVectorSubSetToVector'
'==============================================================================

'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_ScalarInput_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 1
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            Scalar, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_ScalarResult_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray() As Long
    Dim ScalarResult As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 1
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ScalarResult, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_UnallocatedInputArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray() As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 1
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_2DInputArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 1
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(0 To 1, 0 To 1)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_2DResultArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 1
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(0 To 1)
    ReDim ResultArray(0 To 1, 0 To 1)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_TooSmallFirstElementToCopy_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = -1
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(0 To 1)
    ReDim ResultArray(0 To 1, 0 To 1)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_TooLargeLastElementToCopy_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 1
    Const LastElementToCopy As Long = 2
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(0 To 1)
    ReDim ResultArray(0 To 1, 0 To 1)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_FirstElementLargerLastElement_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray() As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 1
    Const LastElementToCopy As Long = 0
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(0 To 1)
    ReDim ResultArray(0 To 1, 0 To 1)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_NotEnoughRoomInStaticResultArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(0 To 1) As Long
    Dim ResultArray(0 To 1) As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 0
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_TooSmallDestinationElementInStaticResultArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(0 To 1) As Long
    Dim ResultArray(5 To 7) As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 0
    Const LastElementToCopy As Long = 1
    Const DestinationElement As Long = 1
    '==========================================================================
    
    
    'Arrange:
    InputArray(0) = 0
    InputArray(1) = 1
    
    ResultArray(5) = 10
    ResultArray(6) = 20
    ResultArray(7) = 30
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_UnallocatedResultArrayDestinationElementLargerBase_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 11) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 10
    Const DestinationElement As Long = 5
    
    Dim aExpected(1 To 5) As Long
        aExpected(1) = 0
        aExpected(2) = 0
        aExpected(3) = 0
        aExpected(4) = 0
        aExpected(5) = 10
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 10
    InputArray(11) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_UnallocatedResultArrayLastDestinationElementSmallerBase_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 11) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 10
    Const DestinationElement As Long = -5
    
    Dim aExpected(-5 To 1) As Long
        aExpected(-5) = 10
        aExpected(-4) = 0
        aExpected(-3) = 0
        aExpected(-2) = 0
        aExpected(-1) = 0
        aExpected(0) = 0
        aExpected(1) = 0
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 10
    InputArray(11) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_UnallocatedResultArrayFromNegToPos_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 13) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 13
    Const DestinationElement As Long = -1
    
    Dim aExpected(-1 To 2) As Long
        aExpected(-1) = 10
        aExpected(0) = 20
        aExpected(1) = 30
        aExpected(2) = 40
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 10
    InputArray(11) = 20
    InputArray(12) = 30
    InputArray(13) = 40
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_UnallocatedResultArray_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 11) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 10
    Const DestinationElement As Long = 1
    
    Dim aExpected(1 To 1) As Long
        aExpected(1) = 0
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 0
    InputArray(11) = 1
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_SubArrayLargerThanAllocatedResultArray1_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 13) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 13
    Const DestinationElement As Long = -1
    
    Dim aExpected(-1 To 2) As Long
        aExpected(-1) = 0
        aExpected(0) = 1
        aExpected(1) = 2
        aExpected(2) = 3
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 0
    InputArray(11) = 1
    InputArray(12) = 2
    InputArray(13) = 3
    
    ReDim ResultArray(0 To 1)
    ResultArray(0) = 10
    ResultArray(1) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_SubArrayLargerThanAllocatedResultArray2_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 12) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 12
    Const DestinationElement As Long = -1
    
    Dim aExpected(-1 To 1) As Long
        aExpected(-1) = 0
        aExpected(0) = 1
        aExpected(1) = 2
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 0
    InputArray(11) = 1
    InputArray(12) = 2
    
    ReDim ResultArray(0 To 1)
    ResultArray(0) = 10
    ResultArray(1) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_SubArrayLargerThanAllocatedResultArray3_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 12) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 12
    Const DestinationElement As Long = 1
    
    Dim aExpected(1 To 3) As Long
        aExpected(1) = 0
        aExpected(2) = 1
        aExpected(3) = 2
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 0
    InputArray(11) = 1
    InputArray(12) = 2
    
    ReDim ResultArray(1 To 2)
    ResultArray(1) = 10
    ResultArray(2) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_TooSmallFirstDestinationElementInDynamicAllocatedResultArray_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 11) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 11
    Const DestinationElement As Long = -1
    
    Dim aExpected(-1 To 1) As Long
        aExpected(-1) = 0
        aExpected(0) = 1
        aExpected(1) = 20
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 0
    InputArray(11) = 1
    
    ReDim ResultArray(0 To 1)
    ResultArray(0) = 10
    ResultArray(1) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_TooLargeLastDestinationElementInDynamicAllocatedResultArray_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 11) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 11
    Const DestinationElement As Long = 1
    
    Dim aExpected(0 To 2) As Long
        aExpected(0) = 10
        aExpected(1) = 0
        aExpected(2) = 1
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 0
    InputArray(11) = 1
    
    ReDim ResultArray(0 To 1)
    ResultArray(0) = 10
    ResultArray(1) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_DestinationElementEvenLargerThanUboundInDynamicAllocatedResultArray_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 11) As Long
    Dim ResultArray() As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 11
    Const DestinationElement As Long = 5
    
    Dim aExpected(0 To 6) As Long
        aExpected(0) = 10
        aExpected(1) = 20
        aExpected(2) = 0
        aExpected(3) = 0
        aExpected(4) = 0
        aExpected(5) = 11
        aExpected(6) = 12
    '==========================================================================
    
    
    'Arrange:
    InputArray(10) = 11
    InputArray(11) = 12
    
    ReDim ResultArray(0 To 1)
    ResultArray(0) = 10
    ResultArray(1) = 20
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResultArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyVectorSubSetToVector")
Public Sub CopyVectorSubSetToVector_TestWithObjects_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim InputArray(10 To 11) As Object
    Dim ResultArray() As Object
    Dim i As Long
    
    '==========================================================================
    Const FirstElementToCopy As Long = 10
    Const LastElementToCopy As Long = 11
    Const DestinationElement As Long = 6
    
    Dim aExpected(5 To 7) As Object
    With ThisWorkbook.Worksheets(1)
        Set aExpected(5) = .Range("B5")
        Set aExpected(6) = .Range("B10")
        Set aExpected(7) = .Range("B11")
    End With
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(10) = .Range("B10")
        Set InputArray(11) = .Range("B11")
        
        ReDim ResultArray(5 To 6)
        Set ResultArray(5) = .Range("B5")
        Set ResultArray(6) = .Range("B6")
    End With
    
    'Act:
    If Not modArraySupport2.CopyVectorSubSetToVector( _
            InputArray, _
            ResultArray, _
            FirstElementToCopy, _
            LastElementToCopy, _
            DestinationElement _
    ) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(ResultArray) To UBound(ResultArray)
        If ResultArray(i) Is Nothing Then
            Assert.IsNothing aExpected(i)
        Else
            Assert.AreEqual aExpected(i).Address, ResultArray(i).Address
        End If
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
