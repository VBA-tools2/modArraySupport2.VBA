Attribute VB_Name = "NumElementsTest"

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
'unit tests for 'NumElements'
'==============================================================================

'@TestMethod("NumElements")
Public Sub NumElements_NoArray_ReturnsZero()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const Dimension As Long = 1
    
    Const aExpected As Long = 0
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Scalar, Dimension)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumElements")
Public Sub NumElements_UnallocatedArray_ReturnsZero()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const Dimension As Long = 1
    
    Const aExpected As Long = 0
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Arr, Dimension)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumElements")
Public Sub NumElements_DimensionLowerOne_ReturnsZero()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4, 1 To 1) As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const Dimension As Long = 0
    
    Const aExpected As Long = 0
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Arr, Dimension)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumElements")
Public Sub NumElements_DimensionHigherNoOfArrDimensions_ReturnsZero()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4, 1 To 1) As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const Dimension As Long = 4
    
    Const aExpected As Long = 0
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Arr, Dimension)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumElements")
Public Sub NumElements_DimensionOne_ReturnsThree()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4, 1 To 1) As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const Dimension As Long = 1
    
    Const aExpected As Long = 3
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Arr, Dimension)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumElements")
Public Sub NumElements_DimensionTwo_ReturnsTwo()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4, 1 To 1) As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const Dimension As Long = 2
    
    Const aExpected As Long = 2
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Arr, Dimension)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumElements")
Public Sub NumElements_DimensionThree_ReturnsOne()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4, 1 To 1) As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const Dimension As Long = 3
    
    Const aExpected As Long = 1
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Arr, Dimension)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("NumElements")
Public Sub NumElements_DefaultDimension_ReturnsThree()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 7, 3 To 4, 1 To 1) As Long
    Dim iNoOfElements As Long
    
    '==========================================================================
    Const aExpected As Long = 3
    '==========================================================================
    
    
    'Act:
    iNoOfElements = modArraySupport2.NumElements(Arr)
    
    'Assert:
    Assert.AreEqual aExpected, iNoOfElements

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
