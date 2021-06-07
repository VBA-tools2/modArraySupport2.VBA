Attribute VB_Name = "IsVariantArrayConsistentTest"

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
'unit tests for 'IsVariantArrayConsistent'
'==============================================================================

'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsVariantArrayConsistent(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_UnallocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_AllocatedLongTypeArray_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_AllocatedObjectTypeArray_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Object
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_AllocatedVariantTypeArrayConsistentIntegers_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(5 To 6) As Variant
    
    
    'Arrange:
    Arr(5) = -100
    Arr(6) = 3
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_AllocatedVariantTypeArrayConsistentObjects_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(5 To 7) As Variant
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5) = .Range("B5")
        Set Arr(6) = Nothing
        Set Arr(7) = .Range("B7")
    End With
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_AllocatedVariantTypeArrayInconsistentTypes_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(5 To 6) As Variant
    
    
    'Arrange:
    Arr(5) = -100
    Arr(6) = "abc"
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_2DAllocatedVariantTypeArrayConsistentIntegers_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Variant
    
    
    'Arrange:
    Arr(5, 3) = 10
    Arr(6, 3) = 11
    Arr(5, 4) = 20
    Arr(6, 4) = 21
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_2DAllocatedVariantTypeArrayConsistentObjects_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Variant
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set Arr(5, 3) = .Range("B5")
        Set Arr(6, 3) = Nothing
        Set Arr(5, 4) = .Range("B7")
        Set Arr(6, 4) = Nothing
    End With
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVariantArrayConsistent")
Public Sub IsVariantArrayConsistent_2DAllocatedVariantTypeArrayInconsistentTypes_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(5 To 6, 3 To 4) As Variant
    
    
    'Arrange:
    Arr(5, 3) = -100
    Arr(6, 3) = "abc"
    Arr(5, 4) = Empty
    Set Arr(6, 4) = Nothing
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsVariantArrayConsistent(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
