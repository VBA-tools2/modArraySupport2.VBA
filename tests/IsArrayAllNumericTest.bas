Attribute VB_Name = "IsArrayAllNumericTest"

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
'unit tests for 'IsArrayAllNumeric'
'==============================================================================

'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Scalar, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_UnallocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_IncludingNumericStringAllowNumericStringsFalse_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = "100"
    Arr(2) = 2
    Arr(3) = Empty
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_IncludingNumericStringAllowNumericStringsTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = True
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = "100"
    Arr(2) = 2
    Arr(3) = Empty
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_IncludingNonNumericString_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = True
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = "abc"
    Arr(2) = 2
    Arr(3) = Empty
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_Numeric1DVariantArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = 123
    Arr(2) = 456
    Arr(3) = 789
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_1DVariantArrayWithObject_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = 123
    Set Arr(2) = ThisWorkbook.Worksheets(1).Range("A1")
    Arr(3) = 789
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_1DVariantArrayWithUnallocatedEntry_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = 123
    Arr(3) = 789
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_Numeric2DVariantArray_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(1 To 3, 4 To 5) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1, 4) = 123
    Arr(2, 4) = 456
    Arr(3, 4) = 789
    
    Arr(1, 5) = -5
    Arr(3, 5) = -10
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_2DVariantArrayWithObject_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(1 To 3, 4 To 5) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1, 4) = 123
    Set Arr(2, 4) = ThisWorkbook.Worksheets(1).Range("A1")
    Arr(3, 4) = 789
    
    Arr(1, 5) = -5
    Arr(3, 5) = -10
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowArrayElementsFalse_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = 123
    Arr(2) = Array(-5)
    Arr(3) = 789
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowArrayElementsTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = 123
    Arr(2) = Array(-5)
    Arr(3) = 789
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowArrayElementsTrue_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = False
    Const AllowArrayElements As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = 123
    Arr(2) = Array(-5, "-5")
    Arr(3) = 789
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsArrayAllNumeric")
Public Sub IsArrayAllNumeric_1DVariantArrayWithArrayAllowNumericStringsTrueAllowArrayElementsTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(1 To 3) As Variant
    
    '==========================================================================
    Const AllowNumericStrings As Boolean = True
    Const AllowArrayElements As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    Arr(1) = 123
    Arr(2) = Array(-5, "-5")
    Arr(3) = 789
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsArrayAllNumeric( _
            Arr, _
            AllowNumericStrings, _
            AllowArrayElements _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
