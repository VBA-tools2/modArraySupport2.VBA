Attribute VB_Name = "IsArrayObjectsTest"

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
'unit tests for 'IsArrayObjects'
'==============================================================================

'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    '==========================================================================
    Const AllowNothing As Boolean = True
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayObjects(Scalar, AllowNothing)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_LongPtrInputArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6) As Long
    
    '==========================================================================
    Const AllowNothing As Boolean = True
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayObjects(InputArray, AllowNothing)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_ObjectInputArrayNothingOnlyAllowNothingTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Object
    Dim Element As Variant
    
    '==========================================================================
    Const AllowNothing As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    Set InputArray(5) = Nothing
    Set InputArray(6) = Nothing
    
    'Act:
    If Not modArraySupport2.IsArrayObjects(InputArray, AllowNothing) Then _
            GoTo TestFail
    
    'Assert:
    For Each Element In InputArray
        Assert.IsNothing Element
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_ObjectInputArrayNothingOnlyAllowNothingFalse_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6) As Object
    
    '==========================================================================
    Const AllowNothing As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Set InputArray(5) = Nothing
    Set InputArray(6) = Nothing
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayObjects(InputArray, AllowNothing)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_ObjectInputArrayNonNothingOnlyAllowNothingTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Object
    Dim Element As Variant
    
    '==========================================================================
    Const AllowNothing As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = .Range("B6")
    End With
    
    'Act:
    If Not modArraySupport2.IsArrayObjects(InputArray, AllowNothing) Then _
            GoTo TestFail
    
    'Assert:
    For Each Element In InputArray
        Assert.IsNotNothing Element
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_ObjectInputArrayNonNothingOnlyAllowNothingFalse_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6) As Object
    Dim Element As Variant
    
    '==========================================================================
    Const AllowNothing As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = .Range("B6")
    End With
    
    'Act:
    If Not modArraySupport2.IsArrayObjects(InputArray, AllowNothing) Then _
            GoTo TestFail
    
    'Assert:
    For Each Element In InputArray
        Assert.IsNotNothing Element
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_VariantInputArrayAllowNothingFalse_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6) As Variant
    Dim Element As Variant
    
    '==========================================================================
    Const AllowNothing As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = Nothing
    End With
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayObjects(InputArray, AllowNothing)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_VariantInputArrayAllowNothingTrue_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6) As Variant
    Dim Element As Variant
    
    '==========================================================================
    Const AllowNothing As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5) = .Range("B5")
        Set InputArray(6) = Nothing
    End With
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayObjects(InputArray, AllowNothing)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_2DVariantInputArrayAllowNothingFalse_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6, 3 To 4) As Variant
    Dim Element As Variant
    
    '==========================================================================
    Const AllowNothing As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5, 3) = .Range("B5")
        Set InputArray(6, 3) = .Range("B6")
        Set InputArray(5, 4) = Nothing
        Set InputArray(6, 4) = Nothing
    End With
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayObjects(InputArray, AllowNothing)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayObjects")
Public Sub IsArrayObjects_2DVariantInputArrayAllowNothingTrue_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6, 3 To 4) As Variant
    Dim Element As Variant
    
    '==========================================================================
    Const AllowNothing As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    With ThisWorkbook.Worksheets(1)
        Set InputArray(5, 3) = .Range("B5")
        Set InputArray(6, 3) = .Range("B6")
        Set InputArray(5, 4) = Nothing
        Set InputArray(6, 4) = Nothing
    End With
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayObjects(InputArray, AllowNothing)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
