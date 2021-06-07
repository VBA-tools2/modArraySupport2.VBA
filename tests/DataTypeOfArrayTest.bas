Attribute VB_Name = "DataTypeOfArrayTest"

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
'unit tests for 'DataTypeOfArray'
'==============================================================================

'@TestMethod("DataTypeOfArray")
Public Sub DataTypeOfArray_NoArray_ReturnsMinusOne()
    On Error GoTo TestFail

    'Arrange:
    Dim sTest As String
    Dim aActual As VbVarType
    
    '==========================================================================
    Const aExpected As Long = -1
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.DataTypeOfArray(sTest)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTypeOfArray")
Public Sub DataTypeOfArray_UnallocatedArray_ReturnsVbDouble()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Double
    Dim aActual As VbVarType
    
    '==========================================================================
    Const aExpected As Long = vbDouble
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.DataTypeOfArray(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTypeOfArray")
Public Sub DataTypeOfArray_Test1DStringArray_ReturnsVbString()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 4) As String
    Dim aActual As VbVarType
    
    '==========================================================================
    Const aExpected As Long = vbString
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.DataTypeOfArray(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTypeOfArray")
Public Sub DataTypeOfArray_Test2DStringArray_ReturnsVbString()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 4, 5 To 6) As String
    Dim aActual As VbVarType
    
    '==========================================================================
    Const aExpected As Long = vbString
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.DataTypeOfArray(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DataTypeOfArray")
Public Sub DataTypeOfArray_Test3DStringArray_ReturnsVbString()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 4, 5 To 6, 8 To 8) As String
    Dim aActual As VbVarType
    
    '==========================================================================
    Const aExpected As Long = vbString
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.DataTypeOfArray(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'TODO: Add tests with Objects
