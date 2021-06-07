Attribute VB_Name = "SetObjectArrayToNothingTest"

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
'unit tests for 'SetObjectArrayToNothing'
'==============================================================================

'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.SetObjectArrayToNothing(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_UnallocatedLongArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.SetObjectArrayToNothing(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_UnallocatedObjectArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.SetObjectArrayToNothing(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_UnallocatedVariantArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Variant
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.SetObjectArrayToNothing(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_1DLongArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.SetObjectArrayToNothing(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_1DObjectArr_ReturnsTrueAndNothingArr()
    On Error GoTo TestFail

    Dim Arr(5 To 7) As Object
    Dim Element As Variant
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = .Range("B5")
        Set Arr(6) = Nothing
        Set Arr(7) = .Range("B7")
    End With
    
    'Act:
    If Not modArraySupport2.SetObjectArrayToNothing(Arr) Then _
            GoTo TestFail
    
    'Assert:
    For Each Element In Arr
        Assert.IsNothing Element
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_1DVariantArr_ReturnsTrueAndNothingArr()
    On Error GoTo TestFail

    Dim Arr(5 To 7) As Variant
    Dim Element As Variant
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = .Range("B5")
        Set Arr(6) = Nothing
        Set Arr(7) = .Range("B7")
    End With
    
    'Act:
    If Not modArraySupport2.SetObjectArrayToNothing(Arr) Then _
            GoTo TestFail
    
    'Assert:
    For Each Element In Arr
        Assert.IsNothing Element
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_1DVariantArrWithEmptyElement_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(5 To 7) As Variant
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = .Range("B5")
        Set Arr(6) = Nothing
        Arr(7) = Empty
    End With
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.SetObjectArrayToNothing(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_2DObjectArr_ReturnsTrueAndNothingArr()
    On Error GoTo TestFail

    Dim Arr(5 To 7, 3 To 4) As Object
    Dim Element As Variant
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5, 3) = .Range("B5")
        Set Arr(6, 3) = Nothing
        Set Arr(7, 3) = .Range("B7")
    
        Set Arr(5, 4) = .Range("B9")
        Set Arr(6, 4) = Nothing
        Set Arr(7, 4) = .Range("B11")
    End With
    
    'Act:
    If Not modArraySupport2.SetObjectArrayToNothing(Arr) Then _
            GoTo TestFail
    
    'Assert:
    For Each Element In Arr
        Assert.IsNothing Element
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_3DObjectArr_ReturnsTrueAndNothingArr()
    On Error GoTo TestFail

    Dim Arr(5 To 7, 3 To 4, 2 To 2) As Object
    Dim Element As Variant
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5, 3, 2) = .Range("B5")
        Set Arr(6, 3, 2) = Nothing
        Set Arr(7, 3, 2) = .Range("B7")
    
        Set Arr(5, 4, 2) = .Range("B9")
        Set Arr(6, 4, 2) = Nothing
        Set Arr(7, 4, 2) = .Range("B11")
    End With
    
    'Act:
    If Not modArraySupport2.SetObjectArrayToNothing(Arr) Then _
            GoTo TestFail
    
    'Assert:
    For Each Element In Arr
        Assert.IsNothing Element
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SetObjectArrayToNothing")
Public Sub SetObjectArrayToNothing_4DObjectArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4, 2 To 2, 1 To 1) As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.SetObjectArrayToNothing(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
