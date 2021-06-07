Attribute VB_Name = "IsArrayAllDefaultTest"

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
'unit tests for 'IsArrayAllDefault'
'==============================================================================

'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllDefault(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub IsArrayAllDefault_UnallocatedArray_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray() As Long
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_DefaultVariantArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    
    
    'Arrange:
    InputArray(5) = Empty
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_NonDefaultVariantArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 5) As Variant
    
    
    'Arrange:
    InputArray(5) = 10
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_DefaultStringArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As String
    
    
    'Arrange:
    InputArray(5) = vbNullString
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_NonDefaultStringArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 5) As String
    
    
    'Arrange:
    InputArray(5) = "abc"
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_DefaultNumericArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Long
    
    
    'Arrange:
    InputArray(5) = 0
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_NonDefaultNumericArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 5) As Long
    
    
    'Arrange:
    InputArray(5) = -1
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_Default3DNumericArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6, 3 To 4, -2 To -1) As Long
    
    
    'Arrange:
    InputArray(5, 3, -2) = 0
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_NonDefault3DNumericArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 6, 3 To 4, -2 To -1) As Long
    
    
    'Arrange:
    InputArray(6, 4, -1) = -1
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_DefaultObjectArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Object
    
    
    'Arrange:
    Set InputArray(5) = Nothing
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllDefault")
Public Sub IsArrayAllDefault_NonDefaultObjectArray_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 5) As Object
    
    
    'Arrange:
    Set InputArray(5) = ThisWorkbook.Worksheets(1).Range("B5")
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllDefault(InputArray)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
