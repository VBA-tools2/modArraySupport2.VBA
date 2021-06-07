Attribute VB_Name = "DeleteVectorElementTest"

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
'unit tests for 'DeleteVectorElement'
'==============================================================================

'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = False
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.DeleteVectorElement( _
            Scalar, _
            ElementNumber, _
            ResizeDynamic _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_UnallocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray() As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = False
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_2DArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 7, 1 To 1) As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = False
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_TooLowElementNumber_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 7) As Long
    
    '==========================================================================
    Const ElementNumber As Long = 3
    Const ResizeDynamic As Boolean = False
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_TooHighElementNumber_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 7) As Long
    
    '==========================================================================
    Const ElementNumber As Long = 9
    Const ResizeDynamic As Boolean = False
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveElementOfStaticArray_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = False
    
    Dim aExpected(5 To 7) As Long
        aExpected(5) = 10
        aExpected(6) = 30
        aExpected(7) = 0
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 10
    InputArray(6) = 20
    InputArray(7) = 30
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveElementOfStaticArrayResizeDynamicTrue_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = True
    
    Dim aExpected(5 To 7) As Long
        aExpected(5) = 10
        aExpected(6) = 30
        aExpected(7) = 0
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 10
    InputArray(6) = 20
    InputArray(7) = 30
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveElementOfStaticObjectArray_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As Object
    Dim i As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = False
    
    Dim aExpected(5 To 7) As Object
        With ThisWorkbook.Worksheets(1)
            Set aExpected(5) = .Range("B5")
            Set aExpected(6) = .Range("B7")
            Set aExpected(7) = Nothing
        End With
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = .Range("B6")
        Set InputArray(7) = .Range("B7")
    End With
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
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

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveElementOfDynamicArrayDontResize_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = False
    
    Dim aExpected(5 To 7) As Long
        aExpected(5) = 10
        aExpected(6) = 30
        aExpected(7) = 0
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 7)
    InputArray(5) = 10
    InputArray(6) = 20
    InputArray(7) = 30
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'TODO: why does this test fail?
''@TestMethod("DeleteVectorElement")
'Public Sub DeleteVectorElement_RemoveElementOfDynamicArrayDontResize2_ReturnsTrueAndModifiedInputArray()
'   On Error GoTo TestFail
'
'    Dim InputArray() As Variant
'
'    '==========================================================================
'    Const ElementNumber As Long = 6
'    Const ResizeDynamic As Boolean = False
'
'    Dim aExpected(5 To 7) As Variant
'        aExpected(5) = "abc"
'        aExpected(6) = "ABC"
'        aExpected(7) = vbNullString
'    '==========================================================================
'
'
'    'Arrange:
'    ReDim InputArray(5 To 7)
'    InputArray(5) = "abc"
'    InputArray(6) = 1234
'    InputArray(7) = "ABC"
'
'    'Act:
'    If Not modArraySupport2.DeleteVectorElement( _
'            InputArray, _
'            ElementNumber, _
'            ResizeDynamic _
'    ) Then _
'            GoTo TestFail
'
'    'Assert:
'    Assert.SequenceEquals aExpected, InputArray
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveElementOfDynamicObjectArrayDontResize_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As Object
    Dim i As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = False
    
    Dim aExpected(5 To 7) As Object
        With ThisWorkbook.Worksheets(1)
            Set aExpected(5) = .Range("B5")
            Set aExpected(6) = .Range("B7")
            Set aExpected(7) = Nothing
        End With
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 7)
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = .Range("B6")
        Set InputArray(7) = .Range("B7")
    End With
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
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

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveElementOfDynamicArrayResize_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = True
    
    Dim aExpected(5 To 6) As Long
        aExpected(5) = 10
        aExpected(6) = 30
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 7)
    InputArray(5) = 10
    InputArray(6) = 20
    InputArray(7) = 30
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveElementOfDynamicObjectArrayResize_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As Object
    Dim i As Long
    
    '==========================================================================
    Const ElementNumber As Long = 6
    Const ResizeDynamic As Boolean = True
    
    Dim aExpected(5 To 6) As Object
        With ThisWorkbook.Worksheets(1)
            Set aExpected(5) = .Range("B5")
            Set aExpected(6) = .Range("B7")
        End With
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 7)
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = .Range("B6")
        Set InputArray(7) = .Range("B7")
    End With
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
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

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveOnlyElementOfDynamicObjectArrayResize_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As String
    Dim i As Long
    
    '==========================================================================
    Const ElementNumber As Long = 5
    Const ResizeDynamic As Boolean = True
    
    Dim aExpected() As String
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 5)
    InputArray(5) = "abc"
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.AreEqual aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DeleteVectorElement")
Public Sub DeleteVectorElement_RemoveOnlyElementOfDynamicObjectArrayDontResize_ReturnsTrueAndModifiedInputArray()
    On Error GoTo TestFail

    Dim InputArray() As String
    Dim i As Long
    
    '==========================================================================
    Const ElementNumber As Long = 5
    Const ResizeDynamic As Boolean = False
    
    Dim aExpected(5 To 5) As String
    aExpected(5) = vbNullString
    '==========================================================================
    
    
    'Arrange:
    ReDim InputArray(5 To 5)
    InputArray(5) = "abc"
    
    'Act:
    If Not modArraySupport2.DeleteVectorElement( _
            InputArray, _
            ElementNumber, _
            ResizeDynamic _
    ) Then _
            GoTo TestFail
    
    'Assert:
'    Assert.AreEqual aExpected(5), InputArray(5)
    Assert.SequenceEquals aExpected, InputArray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
