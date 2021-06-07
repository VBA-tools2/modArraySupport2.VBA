Attribute VB_Name = "IsArrayAllocatedTest"

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
'unit tests for 'IsArrayAllocated'
'==============================================================================

'@TestMethod("IsArrayAllocated")
Public Sub IsArrayAllocated_AllocatedArray_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim AllocatedArray(1 To 3) As Variant
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllocated(AllocatedArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllocated")
Public Sub IsArrayAllocated_UnAllocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim UnAllocatedArray() As Variant
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllocated(UnAllocatedArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
