Attribute VB_Name = "CopyNonNothingObjectsToVectorTest"

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
'unit tests for 'CopyNonNothingObjectsToVector'
'==============================================================================

'@TestMethod("CopyNonNothingObjectsToVector")
Public Sub CopyNonNothingObjectsToVector_ScalarResultArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim SourceArray() As Object
    Dim ScalarResult As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyNonNothingObjectsToVector( _
            SourceArray, _
            ScalarResult _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyNonNothingObjectsToVector")
Public Sub CopyNonNothingObjectsToVector_StaticResultArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim SourceArray() As Object
    Dim ResultArray(1 To 2) As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyNonNothingObjectsToVector( _
            SourceArray, _
            ResultArray _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyNonNothingObjectsToVector")
Public Sub CopyNonNothingObjectsToVector_2DResultArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim SourceArray() As Object
    Dim ResultArray() As Object
    
    
    'Arrange:
    ReDim ResultArray(1 To 2, 1 To 1)
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyNonNothingObjectsToVector( _
            SourceArray, _
            ResultArray _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyNonNothingObjectsToVector")
Public Sub CopyNonNothingObjectsToVector_NonObjectOnlySourceArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim SourceArray(5 To 6) As Variant
    Dim ResultArray() As Object
    
    
    'Arrange:
    Set SourceArray(5) = Nothing
    SourceArray(6) = 1
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyNonNothingObjectsToVector( _
            SourceArray, _
            ResultArray _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyNonNothingObjectsToVector")
Public Sub CopyNonNothingObjectsToVector_ValidNonNothingOnlySourceArray_ReturnsTrueAndResultArray()
    On Error GoTo TestFail

    Dim SourceArray(5 To 6) As Variant
    Dim ResultArray() As Object
    Dim i As Long
    
    
    'Arrange:
    Set SourceArray(5) = Nothing
    Set SourceArray(6) = ThisWorkbook.Worksheets(1).Range("B2")
    
    'Act:
    If Not modArraySupport2.CopyNonNothingObjectsToVector( _
            SourceArray, _
            ResultArray _
    ) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(ResultArray) To UBound(ResultArray)
        Assert.IsNotNothing ResultArray(i)
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyNonNothingObjectsToVector")
Public Sub CopyNonNothingObjectsToVector_NothingOnlySourceArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim SourceArray(5 To 6) As Variant
    Dim ResultArray() As Object
    Dim i As Long
    
    
    'Arrange:
    Set SourceArray(5) = Nothing
    Set SourceArray(6) = Nothing
    
    'Act:
    If Not modArraySupport2.CopyNonNothingObjectsToVector( _
            SourceArray, _
            ResultArray _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllocated(ResultArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
