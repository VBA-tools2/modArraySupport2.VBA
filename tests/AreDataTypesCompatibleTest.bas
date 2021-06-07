Attribute VB_Name = "AreDataTypesCompatibleTest"

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
'unit tests for 'AreDataTypesCompatible'
'==============================================================================

'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_ScalarSourceArrayDest_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Source As Long
    Dim Dest() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_BothStringScalars_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Source As String
    Dim Dest As String
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_BothStringArrays_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Source() As String
    Dim Dest() As String
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_LongSourceIntegerDest_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Source As Long
    Dim Dest As Integer
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_IntegerSourceLongDest_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Source As Integer
    Dim Dest As Long
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_DoubleSourceLongDest_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Source As Double
    Dim Dest As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_BothObjects_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Source As Object
    Dim Dest As Object
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AreDataTypesCompatible")
Public Sub AreDataTypesCompatible_SingleSourceDateDest_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Source As Single
    Dim Dest As Date
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.AreDataTypesCompatible(Source, Dest)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


''TODO: How to do this test?
''     --> in 'ChangeBoundsOfVector_VariantArr_ReturnsTrueAndChangedArr' are
''         'Empty' entries added at the end of the array
''@TestMethod("AreDataTypesCompatible")
'Public Sub AreDataTypesCompatible_VariantSourceEmptyDest_ReturnsTrue()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim Source(0) As Variant
'    Dim Dest(0) As Variant
'    Dim vDummy As Variant
'
'
'    'Act:
'    vDummy = 4534
'    Source(0) = CVar(vDummy)
'    Dest(0) = Empty
'
'    'Assert:
'    Assert.IsTrue modArraySupport2.AreDataTypesCompatible(Source(0), Dest(0))
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub
