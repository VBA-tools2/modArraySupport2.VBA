Attribute VB_Name = "IsVectorSortedTest"

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
'unit tests for 'IsVectorSorted'
'==============================================================================

'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_NoArray_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            Scalar, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue IsNull(aResult)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_UnallocatedArray_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray() As Long
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue IsNull(aResult)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_2DArray_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6, 3 To 4) As Long
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue IsNull(aResult)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_ObjectArray_ReturnsNull()
    On Error GoTo TestFail

    'Arrange:
    Dim InputArray(5 To 6) As Object
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue IsNull(aResult)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_StringArrayDescendingFalse_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As String
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = "ABC"
    InputArray(6) = "abc"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayContainingObjectDescendingFalse_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    Set InputArray(5) = ThisWorkbook.Worksheets(1).Range("B5")
    InputArray(6) = vbNullString
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArraySmallNumericStringPlusLargerNumberDescendingFalse_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = "45"
    InputArray(6) = 123
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArraySmallNumberPlusLargerNumericStringDescendingFalse_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 45
    InputArray(6) = "123"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayLargeNumberPlusSmallNumericStringDescendingFalse_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    '(it seems that the numbers are always considered smaller than any string)
    InputArray(5) = 9
    InputArray(6) = ""
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayNumberPlusStringDescendingFalse_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 45
    InputArray(6) = "abc"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayNumberPlusStringsDescendingFalse_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 8) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    '(but then strings seem to be compared as usual)
    InputArray(5) = 5
    InputArray(6) = "1"
    InputArray(7) = "Abc"
    InputArray(8) = "defg"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayNumberPlusStrings2DescendingFalse_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 8) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = False
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 5
    InputArray(6) = "zbc"
    InputArray(7) = "defg"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_StringArrayDescendingTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As String
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = "ABC"
    InputArray(6) = "abc"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayContainingObjectDescendingTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    Set InputArray(5) = ThisWorkbook.Worksheets(1).Range("B5")
    InputArray(6) = vbNullString
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArraySmallNumericStringPlusLargerNumberDescendingTrue_ReturnsTrue()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = "45"
    InputArray(6) = 123
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsTrue aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArraySmallNumberPlusLargerNumericStringDescendingTrue_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 45
    InputArray(6) = "123"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayLargeNumberPlusSmallNumericStringDescendingTrue_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    '(it seems that the numbers are always considered smaller than any string)
    InputArray(5) = 9
    InputArray(6) = ""
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayNumberPlusStringDescendingTrue_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 6) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 45
    InputArray(6) = "abc"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayNumberPlusStringsDescendingTrue_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 8) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    '(but then strings seem to be compared as usual)
    InputArray(5) = 5
    InputArray(6) = "1"
    InputArray(7) = "Abc"
    InputArray(8) = "defg"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsVectorSorted")
Public Sub IsVectorSorted_VariantArrayNumberPlusStrings2DescendingTrue_ReturnsFalse()
    On Error GoTo TestFail

    Dim InputArray(5 To 8) As Variant
    Dim aResult As Variant
    
    '==========================================================================
    Const Descending As Boolean = True
    '==========================================================================
    
    
    'Arrange:
    InputArray(5) = 5
    InputArray(6) = "zbc"
    InputArray(7) = "defg"
    
    'Act:
    aResult = modArraySupport2.IsVectorSorted( _
            InputArray, _
            Descending _
    )
    
    'Assert:
    Assert.IsFalse aResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
