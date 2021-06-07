Attribute VB_Name = "IsArrayDynamicTest"

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
'unit tests for 'IsArrayDynamic'
'==============================================================================

'@TestMethod("IsArrayDynamic")
Public Sub IsArrayDynamic_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayDynamic(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayDynamic")
Public Sub IsArrayDynamic_UnallocatedArray_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayDynamic(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayDynamic")
Public Sub IsArrayDynamic_1DDynamicArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr() As Long
    
    
    'Arrange:
    ReDim Arr(5 To 6)
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayDynamic(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayDynamic")
Public Sub IsArrayDynamic_1DStaticArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayDynamic(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayDynamic")
Public Sub IsArrayDynamic_2DDynamicArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr() As Long
    
    
    'Arrange:
    ReDim Arr(5 To 6, 3 To 4)
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayDynamic(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayDynamic")
Public Sub IsArrayDynamic_2DStaticArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6, 3 To 4) As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayDynamic(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
