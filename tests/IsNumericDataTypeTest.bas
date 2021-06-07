Attribute VB_Name = "IsNumericDataTypeTest"

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
'unit tests for 'IsNumericDataType'
'==============================================================================

'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_LongPtrScalar_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsNumericDataType(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_CurrencyScalar_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Currency
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsNumericDataType(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_StringScalar_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As String
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_ObjectScalar_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_VariantScalarUninitialized_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Variant
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_VariantScalarNumericContent_ReturnsTrue()
    On Error GoTo TestFail

    Dim Scalar As Variant
    
    
    'Arrange:
    Scalar = 3
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsNumericDataType(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_VariantScalarNonNumericContent_ReturnsFalse()
    On Error GoTo TestFail

    Dim Scalar As Variant
    
    
    'Arrange:
    Scalar = "abc"
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_LongPtrArrayUnallocated_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_LongPtrStaticArray_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(5 To 6) As Long
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_CurrencyArray_ReturnsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Currency
    
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_StringArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As String
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_ObjectArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Object
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_VariantArrayUnallocated_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Variant
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_StaticVariantArrayNumericContent_ReturnsTrue()
    On Error GoTo TestFail

    Dim Arr(5 To 6) As Variant
    
    
    'Arrange:
    Arr(5) = 3
    Arr(6) = 7.8
    
    'Act:
    'Assert:
    Assert.IsTrue modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_StaticVariantArrayMixedContentNumericFirst_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(5 To 6) As Variant
    
    
    'Arrange:
    Arr(5) = -2
    Arr(6) = "abc"
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("IsNumericDataType")
Public Sub IsNumericDataType_StaticVariantArrayMixedContentNonNumericFirst_ReturnsFalse()
    On Error GoTo TestFail

    Dim Arr(5 To 6) As Variant
    
    
    'Arrange:
    Arr(5) = "abc"
    Arr(6) = -2
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.IsNumericDataType(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
