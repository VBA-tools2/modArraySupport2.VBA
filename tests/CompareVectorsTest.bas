Attribute VB_Name = "CompareVectorsTest"

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
'unit tests for 'CompareVectors'
'==============================================================================

'@TestMethod("CompareVectors")
Public Sub CompareVectors_UnallocatedArrays_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr1() As String
    Dim Arr2() As String
    Dim ResArr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CompareVectors(Arr1, Arr2, ResArr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CompareVectors")
Public Sub CompareVectors_LegalAndTextCompare_ReturnsTrueAndResArr()
    On Error GoTo TestFail

    Dim Arr1(1 To 5) As String
    Dim Arr2(1 To 5) As String
    Dim ResArr() As Long
    
    '==========================================================================
    Dim aExpected(1 To 5) As Long
        aExpected(1) = -1
        aExpected(2) = 1
        aExpected(3) = -1
        aExpected(4) = 0
        aExpected(5) = 0
    '==========================================================================
    
    
    'Arrange:
    Arr1(1) = "2"
    Arr1(2) = "c"
    Arr1(3) = vbNullString
    Arr1(4) = "."
    Arr1(5) = "B"
    
    Arr2(1) = "4"
    Arr2(2) = "a"
    Arr2(3) = "x"
    Arr2(4) = "."
    Arr2(5) = "b"
    
    'Act:
    If Not modArraySupport2.CompareVectors(Arr1, Arr2, ResArr, vbTextCompare) _
            Then GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResArr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CompareVectors")
Public Sub CompareVectors_LegalAndBinaryCompare_ReturnsTrueAndResArr()
    On Error GoTo TestFail

    Dim Arr1(1 To 5) As String
    Dim Arr2(1 To 5) As String
    Dim ResArr() As Long
    
    '==========================================================================
    Dim aExpected(1 To 5) As Long
        aExpected(1) = -1
        aExpected(2) = 1
        aExpected(3) = -1
        aExpected(4) = 0
        aExpected(5) = -1
    '==========================================================================
    
    
    'Arrange:
    Arr1(1) = "2"
    Arr1(2) = "c"
    Arr1(3) = vbNullString
    Arr1(4) = "."
    Arr1(5) = "B"
    
    Arr2(1) = "4"
    Arr2(2) = "a"
    Arr2(3) = "x"
    Arr2(4) = "."
    Arr2(5) = "b"
    
    'Act:
    If Not modArraySupport2.CompareVectors(Arr1, Arr2, ResArr, vbBinaryCompare) _
            Then GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, ResArr

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
