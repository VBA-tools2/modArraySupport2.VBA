Attribute VB_Name = "FirstNonEmptyStringIndexInVectorTest"

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
'unit tests for 'FirstNonEmptyStringIndexInVector'
'==============================================================================

'@TestMethod("FirstNonEmptyStringIndexInVector")
Public Sub FirstNonEmptyStringIndexInVector_NoArray_ReturnsMinusOne()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As String
    Dim aActual As Long
    
    '==========================================================================
    Const aExpected As Long = -1
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.FirstNonEmptyStringIndexInVector(Scalar)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FirstNonEmptyStringIndexInVector")
Public Sub FirstNonEmptyStringIndexInVector_UnallocatedArray_ReturnsMinusOne()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray() As String
    Dim aActual As Long
    
    '==========================================================================
    Const aExpected As Long = -1
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.FirstNonEmptyStringIndexInVector(InputArray)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FirstNonEmptyStringIndexInVector")
Public Sub FirstNonEmptyStringIndexInVector_2DArray_ReturnsMinusOne()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6, 3 To 4) As String
    Dim aActual As Long
    
    '==========================================================================
    Const aExpected As Long = -1
    '==========================================================================
    
    
    'Act:
    aActual = modArraySupport2.FirstNonEmptyStringIndexInVector(InputArray)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FirstNonEmptyStringIndexInVector")
Public Sub FirstNonEmptyStringIndexInVector_NoNonEmptyString_ReturnsMinusOne()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As String
    Dim aActual As Long
    
    '==========================================================================
    Const aExpected As Long = -1
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = vbNullString
    InputArray(6) = vbNullString
    InputArray(7) = vbNullString
    
    'Act:
    aActual = modArraySupport2.FirstNonEmptyStringIndexInVector(InputArray)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("FirstNonEmptyStringIndexInVector")
Public Sub FirstNonEmptyStringIndexInVector_WithNonEmptyStringEntry_ReturnsSeven()
    On Error GoTo TestFail

    Dim InputArray(5 To 7) As String
    Dim aActual As Long
    
    '==========================================================================
    Const aExpected As Long = 7
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = vbNullString
    InputArray(6) = ""
    InputArray(7) = "ghi"
    
    'Act:
    aActual = modArraySupport2.FirstNonEmptyStringIndexInVector(InputArray)
    
    'Assert:
    Assert.AreEqual aExpected, aActual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
