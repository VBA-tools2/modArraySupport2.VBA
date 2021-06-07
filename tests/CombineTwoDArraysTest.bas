Attribute VB_Name = "CombineTwoDArraysTest"

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
'unit tests for 'CombineTwoDArrays'
'==============================================================================

'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_ScalarArr1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar1 As Long
    Dim Arr2(1 To 2, 2 To 3) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Scalar1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_ScalarArr2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 1 To 2) As Long
    Dim Scalar2 As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Scalar2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_1DArr1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3) As Long
    Dim Arr2(1 To 3, 1 To 2) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_3DArr1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 1 To 2, 1 To 4) As Long
    Dim Arr2(1 To 3, 1 To 2) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_1DArr2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 1 To 2) As Long
    Dim Arr2(1 To 3) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_3DArr2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 1 To 2) As Long
    Dim Arr2(1 To 3, 1 To 2, 1 To 4) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_DifferentColNumbers_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 1 To 2) As Long
    Dim Arr2(1 To 3, 1 To 3) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_DifferentLBoundRows_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 1 To 2) As Long
    Dim Arr2(2 To 3, 1 To 2) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_DifferentLBoundCol1_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 2 To 3) As Long
    Dim Arr2(1 To 3, 1 To 2) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_DifferentLBoundCol2_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1(1 To 3, 1 To 2) As Long
    Dim Arr2(1 To 3, 2 To 3) As Long
    Dim ResArr As Variant
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.IsTrue IsNull(ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_1BasedStringArrays_ReturnsCombinedResultArr()
    On Error GoTo TestFail

    Dim Arr1(1 To 2, 1 To 2) As String
    Dim Arr2(1 To 2, 1 To 2) As String
    Dim ResArr As Variant
    
    '==========================================================================
    Dim aExpected(1 To 4, 1 To 2) As Variant
        aExpected(1, 1) = "a"
        aExpected(1, 2) = "b"
        aExpected(2, 1) = "c"
        aExpected(2, 2) = "d"
        
        aExpected(3, 1) = "e"
        aExpected(3, 2) = "f"
        aExpected(4, 1) = "g"
        aExpected(4, 2) = "h"
    '==========================================================================
    
    
    'Arrange:
    Arr1(1, 1) = "a"
    Arr1(1, 2) = "b"
    Arr1(2, 1) = "c"
    Arr1(2, 2) = "d"
    
    Arr2(1, 1) = "e"
    Arr2(1, 2) = "f"
    Arr2(2, 1) = "g"
    Arr2(2, 2) = "h"
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.SequenceEquals aExpected, ResArr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_0BasedStringArrays_ReturnsCombinedResultArr()
    On Error GoTo TestFail

    Dim Arr1(0 To 1, 0 To 1) As String
    Dim Arr2(0 To 1, 0 To 1) As String
    Dim ResArr As Variant
    
    '==========================================================================
    Dim aExpected(0 To 3, 0 To 1) As Variant
        aExpected(0, 0) = "a"
        aExpected(0, 1) = "b"
        aExpected(1, 0) = "c"
        aExpected(1, 1) = "d"
        
        aExpected(2, 0) = "e"
        aExpected(2, 1) = "f"
        aExpected(3, 0) = "g"
        aExpected(3, 1) = "h"
    '==========================================================================
    
    
    'Arrange:
    Arr1(0, 0) = "a"
    Arr1(0, 1) = "b"
    Arr1(1, 0) = "c"
    Arr1(1, 1) = "d"
    
    Arr2(0, 0) = "e"
    Arr2(0, 1) = "f"
    Arr2(1, 0) = "g"
    Arr2(1, 1) = "h"
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.SequenceEquals aExpected, ResArr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_PositiveBasedStringArrays_ReturnsCombinedResultArr()
    On Error GoTo TestFail

    Dim Arr1(5 To 6, 5 To 6) As String
    Dim Arr2(5 To 6, 5 To 6) As String
    Dim ResArr As Variant
    
    '==========================================================================
    Dim aExpected(5 To 8, 5 To 6) As Variant
        aExpected(5, 5) = "a"
        aExpected(5, 6) = "b"
        aExpected(6, 5) = "c"
        aExpected(6, 6) = "d"
        
        aExpected(7, 5) = "e"
        aExpected(7, 6) = "f"
        aExpected(8, 5) = "g"
        aExpected(8, 6) = "h"
    '==========================================================================
    
    
    'Arrange:
    Arr1(5, 5) = "a"
    Arr1(5, 6) = "b"
    Arr1(6, 5) = "c"
    Arr1(6, 6) = "d"
    
    Arr2(5, 5) = "e"
    Arr2(5, 6) = "f"
    Arr2(6, 5) = "g"
    Arr2(6, 6) = "h"
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.SequenceEquals aExpected, ResArr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_NegativeBasedStringArrays_ReturnsCombinedResultArr()
    On Error GoTo TestFail

    Dim Arr1(-6 To -5, -6 To -5) As String
    Dim Arr2(-6 To -5, -6 To -5) As String
    Dim ResArr As Variant
    
    '==========================================================================
    Dim aExpected(-6 To -3, -6 To -5) As Variant
        aExpected(-6, -6) = "a"
        aExpected(-6, -5) = "b"
        aExpected(-5, -6) = "c"
        aExpected(-5, -5) = "d"
        
        aExpected(-4, -6) = "e"
        aExpected(-4, -5) = "f"
        aExpected(-3, -6) = "g"
        aExpected(-3, -5) = "h"
    '==========================================================================
    
    
    'Arrange:
    Arr1(-6, -6) = "a"
    Arr1(-6, -5) = "b"
    Arr1(-5, -6) = "c"
    Arr1(-5, -5) = "d"
    
    Arr2(-6, -6) = "e"
    Arr2(-6, -5) = "f"
    Arr2(-5, -6) = "g"
    Arr2(-5, -5) = "h"
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays(Arr1, Arr2)
    
    'Assert:
    Assert.SequenceEquals aExpected, ResArr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CombineTwoDArrays")
Public Sub CombineTwoDArrays_NestedStringArrays_ReturnsCombinedResultArr()
    On Error GoTo TestFail

    Dim Arr1(1 To 2, 1 To 2) As String
    Dim Arr2(1 To 2, 1 To 2) As String
    Dim Arr3(1 To 2, 1 To 2) As String
    Dim Arr4(1 To 2, 1 To 2) As String
    Dim ResArr As Variant
    
    '==========================================================================
    Dim aExpected(1 To 8, 1 To 2) As Variant
        aExpected(1, 1) = "a"
        aExpected(1, 2) = "b"
        aExpected(2, 1) = "c"
        aExpected(2, 2) = "d"
        
        aExpected(3, 1) = "e"
        aExpected(3, 2) = "f"
        aExpected(4, 1) = "g"
        aExpected(4, 2) = "h"
        
        aExpected(5, 1) = "i"
        aExpected(5, 2) = "j"
        aExpected(6, 1) = "k"
        aExpected(6, 2) = "l"
        
        aExpected(7, 1) = "m"
        aExpected(7, 2) = "n"
        aExpected(8, 1) = "o"
        aExpected(8, 2) = "p"
    '==========================================================================
    
    
    'Arrange:
    Arr1(1, 1) = "a"
    Arr1(1, 2) = "b"
    Arr1(2, 1) = "c"
    Arr1(2, 2) = "d"
    
    Arr2(1, 1) = "e"
    Arr2(1, 2) = "f"
    Arr2(2, 1) = "g"
    Arr2(2, 2) = "h"
    
    Arr3(1, 1) = "i"
    Arr3(1, 2) = "j"
    Arr3(2, 1) = "k"
    Arr3(2, 2) = "l"
    
    Arr4(1, 1) = "m"
    Arr4(1, 2) = "n"
    Arr4(2, 1) = "o"
    Arr4(2, 2) = "p"
    
    'Act:
    ResArr = modArraySupport2.CombineTwoDArrays( _
            modArraySupport2.CombineTwoDArrays( _
                    modArraySupport2.CombineTwoDArrays(Arr1, Arr2), _
                    Arr3), _
            Arr4 _
    )
    
    'Assert:
    Assert.SequenceEquals aExpected, ResArr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
