Attribute VB_Name = "ChangeBoundsOfVectorTest"

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
'unit tests for 'ChangeBoundsOfVector'
'==============================================================================

'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_LBGreaterUB_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(2 To 4) As Long
    
    '==========================================================================
    Const NewLB As Long = 5
    Const NewUB As Long = 3
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_ScalarInput_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Const Scalar As Long = 1
    
    '==========================================================================
    Const NewLB As Long = 3
    Const NewUB As Long = 5
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ChangeBoundsOfVector(Scalar, NewLB, NewUB)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_StaticArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(2 To 4) As Long
    
    '==========================================================================
    Const NewLB As Long = 3
    Const NewUB As Long = 5
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_UnallocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    
    '==========================================================================
    Const NewLB As Long = 3
    Const NewUB As Long = 5
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_2DArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(2 To 5, 1 To 1) As Long
    
    '==========================================================================
    Const NewLB As Long = 3
    Const NewUB As Long = 5
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_LongInputArr_ReturnsTrueAndChangedArr()
    On Error GoTo TestFail

    Dim Arr() As Long
    
    '==========================================================================
    Const NewLB As Long = 20
    Const NewUB As Long = 25
    '==========================================================================
    Dim aExpected(NewLB To NewUB) As Long
        aExpected(20) = 11
        aExpected(21) = 22
        aExpected(22) = 33
        aExpected(23) = 0
        aExpected(24) = 0
        aExpected(25) = 0
    '==========================================================================
    
    'Arrange:
    ReDim Arr(5 To 7)
    Arr(5) = 11
    Arr(6) = 22
    Arr(7) = 33
    
    
    'Act:
    If Not modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
            Then GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Arr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_SmallerUBDiffThanSource_ReturnsTrueAndChangedArr()
    On Error GoTo TestFail

    Dim Arr() As Long
    
    '==========================================================================
    Const NewLB As Long = 20
    Const NewUB As Long = 21
    '==========================================================================
    Dim aExpected(NewLB To NewUB) As Long
        aExpected(20) = 11
        aExpected(21) = 22
    '==========================================================================
    
    'Arrange:
    ReDim Arr(5 To 7)
    Arr(5) = 11
    Arr(6) = 22
    Arr(7) = 33
    
    
    'Act:
    If Not modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
            Then GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Arr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_VariantArr_ReturnsTrueAndChangedArr()
    On Error GoTo TestFail

    Dim Arr() As Variant
    Dim i As Long
    
    '==========================================================================
    Const NewLB As Long = 20
    Const NewUB As Long = 25
    '==========================================================================
    Dim aExpected(NewLB To NewUB) As Variant
        aExpected(20) = Array(1, 2, 3)
        aExpected(21) = Array(4, 5, 6)
        aExpected(22) = Array(7, 8, 9)
        aExpected(23) = Empty
        aExpected(24) = Empty
        aExpected(25) = Empty
    '==========================================================================
    
    'Arrange:
    ReDim Arr(5 To 7)
    Arr(5) = Array(1, 2, 3)
    Arr(6) = Array(4, 5, 6)
    Arr(7) = Array(7, 8, 9)
    
    
    'Act:
    If Not modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
            Then GoTo TestFail
    
    'Assert:
    For i = NewLB To NewUB
        If IsArray(Arr(i)) Then
            Assert.SequenceEquals aExpected(i), Arr(i)
        Else
            Assert.AreEqual aExpected(i), Arr(i)
        End If
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_LongInputArrNoUpperBound_ReturnsTrueAndChangedArr()
    On Error GoTo TestFail

    Dim Arr() As Long
'    Dim i As Long
    
    '==========================================================================
    Const NewLB As Long = 20
    Const NewUB As Long = 22
    '==========================================================================
    Dim aExpected(NewLB To NewUB) As Long
        aExpected(20) = 11
        aExpected(21) = 22
        aExpected(22) = 33
    '==========================================================================
    
    'Arrange:
    ReDim Arr(5 To 7)
    Arr(5) = 11
    Arr(6) = 22
    Arr(7) = 33
    
    
    'Act:
    If Not modArraySupport2.ChangeBoundsOfVector(Arr, NewLB) _
            Then GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Arr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'TODO: not sure if the test is done right
'     --> is testing for 'Is(Not)Nothing sufficient?
'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_RangeArr_ReturnsTrueAndChangedArr()
    On Error GoTo TestFail

    Dim Arr() As Range
    Dim i As Long
    
    '==========================================================================
    Const NewLB As Long = 20
    Const NewUB As Long = 25
    '==========================================================================
    Dim aExpected(NewLB To NewUB) As Range
    With ThisWorkbook.Worksheets(1)
        Set aExpected(20) = .Range("A1")
        Set aExpected(21) = .Range("A2")
        Set aExpected(22) = .Range("A3")
        Set aExpected(23) = Nothing
        Set aExpected(24) = Nothing
        Set aExpected(25) = Nothing
    End With
    '==========================================================================
    
    'Arrange:
    ReDim Arr(5 To 7)
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = .Range("A1")
        Set Arr(6) = .Range("A2")
        Set Arr(7) = .Range("A3")
    End With
    
    'Act:
    If Not modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
            Then GoTo TestFail
    
    'Assert:
    For i = NewLB To NewUB
        If aExpected(i) Is Nothing Then
            Assert.IsNothing Arr(i)
        Else
            Assert.IsNotNothing Arr(i)
        End If
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ChangeBoundsOfVector")
Public Sub ChangeBoundsOfVector_CustomClass_ReturnsTrueAndChangedArr()
    On Error GoTo TestFail

    Dim Arr() As cls_4Test_modArraySupport2
    Dim i As Long
    
    '==========================================================================
    Const NewLB As Long = 20
    Const NewUB As Long = 25
    '==========================================================================
    Dim aExpected(NewLB To NewUB) As cls_4Test_modArraySupport2
    Set aExpected(20) = New cls_4Test_modArraySupport2
    Set aExpected(21) = New cls_4Test_modArraySupport2
    Set aExpected(22) = New cls_4Test_modArraySupport2
    aExpected(20).Name = "Name 1"
    aExpected(20).Value = 1
    aExpected(21).Name = "Name 2"
    aExpected(21).Value = 3
    aExpected(22).Name = "Name 3"
    aExpected(22).Value = 3
    Set aExpected(23) = Nothing
    Set aExpected(24) = Nothing
    Set aExpected(25) = Nothing
    '==========================================================================
    
    'Arrange:
    ReDim Arr(5 To 7)
    Set Arr(5) = New cls_4Test_modArraySupport2
    Set Arr(6) = New cls_4Test_modArraySupport2
    Set Arr(7) = New cls_4Test_modArraySupport2
    Arr(5).Name = "Name 1"
    Arr(5).Value = 1
    Arr(6).Name = "Name 2"
    Arr(6).Value = 3
    Arr(7).Name = "Name 3"
    Arr(7).Value = 3
    
    'Act:
    If Not modArraySupport2.ChangeBoundsOfVector(Arr, NewLB, NewUB) _
            Then GoTo TestFail
    
    'Assert:
    For i = NewLB To NewUB
        If aExpected(i) Is Nothing Then
            Assert.IsNothing Arr(i)
        Else
            Assert.IsNotNothing Arr(i)
        End If
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
