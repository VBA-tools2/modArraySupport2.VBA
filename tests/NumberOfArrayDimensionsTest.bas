Attribute VB_Name = "NumberOfArrayDimensionsTest"

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
'unit tests for 'NumberOfArrayDimensions'
'==============================================================================

'@TestMethod("NumberOfArrayDimensions")
Public Sub NumberOfArrayDimensions_UnallocatedLongArray_ReturnsZero()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    Dim iNoOfArrDimensions As Long
    
    '==========================================================================
    Const aExpected As Long = 0
    '==========================================================================
    
    
    'Act:
    iNoOfArrDimensions = modArraySupport2.NumberOfArrayDimensions(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumberOfArrayDimensions")
Public Sub NumberOfArrayDimensions_UnallocatedVariantArray_ReturnsZero()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Variant
    Dim iNoOfArrDimensions As Long
    
    '==========================================================================
    Const aExpected As Long = 0
    '==========================================================================
    
    
    'Act:
    iNoOfArrDimensions = modArraySupport2.NumberOfArrayDimensions(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumberOfArrayDimensions")
Public Sub NumberOfArrayDimensions_UnallocatedObjectArray_ReturnsZero()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Object
    Dim iNoOfArrDimensions As Long
    
    '==========================================================================
    Const aExpected As Long = 0
    '==========================================================================
    
    
    'Act:
    iNoOfArrDimensions = modArraySupport2.NumberOfArrayDimensions(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumberOfArrayDimensions")
Public Sub NumberOfArrayDimensions_1DArray_ReturnsOne()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 3) As Long
    Dim iNoOfArrDimensions As Long
    
    '==========================================================================
    Const aExpected As Long = 1
    '==========================================================================
    
    
    'Act:
    iNoOfArrDimensions = modArraySupport2.NumberOfArrayDimensions(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumberOfArrayDimensions")
Public Sub NumberOfArrayDimensions_3DArray_ReturnsThree()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 3, 1 To 2, 1 To 1)
    Dim iNoOfArrDimensions As Long
    
    '==========================================================================
    Const aExpected As Long = 3
    '==========================================================================
    
    
    'Act:
    iNoOfArrDimensions = modArraySupport2.NumberOfArrayDimensions(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfArrDimensions

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
