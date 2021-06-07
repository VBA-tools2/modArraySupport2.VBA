Attribute VB_Name = "ResetVariantArrayToDefaultsTest"

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
'unit tests for 'ResetVariantArrayToDefaults' 
'(as well as for 'SetVariableToDefault')
'==============================================================================

'@TestMethod("ResetVariantArrayToDefaults")
Public Sub ResetVariantArrayToDefaults_NoArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Scalar As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ResetVariantArrayToDefaults(Scalar)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ResetVariantArrayToDefaults")
Public Sub ResetVariantArrayToDefaults_UnallocatedArray_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr() As Long
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ResetVariantArrayToDefaults(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ResetVariantArrayToDefaults")
Public Sub ResetVariantArrayToDefaults_4DArr_ReturnsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Arr(1 To 8, 4 To 5, 3 To 3, 2 To 2) As Variant
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.ResetVariantArrayToDefaults(Arr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ResetVariantArrayToDefaults")
Public Sub ResetVariantArrayToDefaults_AllSetVariableToDefaultElementsIn1DArr_ReturnsTrueAndResettedArr()
    On Error GoTo TestFail

    Dim Arr(1 To 15) As Variant
    Dim i As Long
    
    '==========================================================================
    Dim aExpected(1 To 15) As Variant
        Set aExpected(1) = Nothing
        aExpected(2) = Array()
            SetVariableToDefault aExpected(2)
        aExpected(3) = False
        aExpected(4) = CByte(0)
        aExpected(5) = CCur(0)
        aExpected(6) = CDate(0)
        aExpected(7) = CDec(0)
        aExpected(8) = CDbl(0)
        aExpected(9) = Empty
        aExpected(10) = Empty
        aExpected(11) = CInt(0)
        aExpected(12) = CLng(0)
        aExpected(13) = Empty
        aExpected(14) = CSng(0)
        aExpected(15) = vbNullString
    '==========================================================================
    
    
    'Arrange:
    Set Arr(1) = ThisWorkbook.Worksheets(1).Range("B5")
    Arr(2) = Array(123)
    Arr(3) = True
    Arr(4) = CByte(1)
    Arr(5) = CCur(1)
    Arr(6) = #2/12/1969#
    Arr(7) = CDec(10000000.0587)
    Arr(8) = CDbl(-123.456)
    Arr(9) = Empty
    Arr(10) = CVErr(xlErrNA)
    Arr(11) = CInt(2345.5678)
    Arr(12) = CLng(123456789)
    Arr(13) = Null
    Arr(14) = CSng(654.321)
    Arr(15) = "abc"
    
    'Act:
    If Not modArraySupport2.ResetVariantArrayToDefaults(Arr) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(Arr) To UBound(Arr)
        If IsObject(Arr(i)) Then
            Assert.IsNothing Arr(i)
        ElseIf IsNull(Arr(i)) Then
            Assert.IsTrue IsNull(Arr(i))
        Else
            Assert.AreEqual aExpected(i), Arr(i)
        End If
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ResetVariantArrayToDefaults")
Public Sub ResetVariantArrayToDefaults_AllSetVariableToDefaultElementsIn2DArr_ReturnsTrueAndResettedArr()
    On Error GoTo TestFail

    Dim Arr(1 To 8, 4 To 5) As Variant
    Dim i As Long
    Dim j As Long
    
    '==========================================================================
    Dim aExpected(1 To 8, 4 To 5) As Variant
        Set aExpected(1, 4) = Nothing
        aExpected(2, 4) = Array()
            SetVariableToDefault aExpected(2, 4)
        aExpected(3, 4) = False
        aExpected(4, 4) = CByte(0)
        aExpected(5, 4) = CCur(0)
        aExpected(6, 4) = CDate(0)
        aExpected(7, 4) = CDec(0)
        aExpected(8, 4) = CDbl(0)
        
        aExpected(1, 5) = Empty
        aExpected(2, 5) = Empty
        aExpected(3, 5) = CInt(0)
        aExpected(4, 5) = CLng(0)
        aExpected(5, 5) = Empty
        aExpected(6, 5) = CSng(0)
        aExpected(7, 5) = vbNullString
        aExpected(8, 5) = Empty     'non-initialized Variant entry
    '==========================================================================
    
    
    'Arrange:
    Set Arr(1, 4) = ThisWorkbook.Worksheets(1).Range("B5")
    Arr(2, 4) = Array(123)
    Arr(3, 4) = True
    Arr(4, 4) = CByte(1)
    Arr(5, 4) = CCur(1)
    Arr(6, 4) = #2/12/1969#
    Arr(7, 4) = CDec(10000000.0587)
    Arr(8, 4) = CDbl(-123.456)
    
    Arr(1, 5) = Empty
    Arr(2, 5) = CVErr(xlErrNA)
    Arr(3, 5) = CInt(2345.5678)
    Arr(4, 5) = CLng(123456789)
    Arr(5, 5) = Null
    Arr(6, 5) = CSng(654.321)
    Arr(7, 5) = "abc"
        
    'Act:
    If Not modArraySupport2.ResetVariantArrayToDefaults(Arr) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            If IsObject(Arr(i, j)) Then
                Assert.IsNothing Arr(i, j)
            Else
                Assert.AreEqual aExpected(i, j), Arr(i, j)
            End If
        Next
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ResetVariantArrayToDefaults")
Public Sub ResetVariantArrayToDefaults_AllSetVariableToDefaultElementsIn3DArr_ReturnsTrueAndResettedArr()
    On Error GoTo TestFail

    Dim Arr(1 To 8, 4 To 5, 3 To 3) As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    '==========================================================================
    Dim aExpected(1 To 8, 4 To 5, 3 To 3) As Variant
        Set aExpected(1, 4, 3) = Nothing
        aExpected(2, 4, 3) = Array()
            SetVariableToDefault aExpected(2, 4, 3)
        aExpected(3, 4, 3) = False
        aExpected(4, 4, 3) = CByte(0)
        aExpected(5, 4, 3) = CCur(0)
        aExpected(6, 4, 3) = CDate(0)
        aExpected(7, 4, 3) = CDec(0)
        aExpected(8, 4, 3) = CDbl(0)
    
        aExpected(1, 5, 3) = Empty
        aExpected(2, 5, 3) = Empty
        aExpected(3, 5, 3) = CInt(0)
        aExpected(4, 5, 3) = CLng(0)
        aExpected(5, 5, 3) = Empty
        aExpected(6, 5, 3) = CSng(0)
        aExpected(7, 5, 3) = vbNullString
        aExpected(8, 5, 3) = Empty     'non-initialized Variant entry
    '==========================================================================
    
    
    'Arrange:
    Set Arr(1, 4, 3) = ThisWorkbook.Worksheets(1).Range("B5")
    Arr(2, 4, 3) = Array(123)
    Arr(3, 4, 3) = True
    Arr(4, 4, 3) = CByte(1)
    Arr(5, 4, 3) = CCur(1)
    Arr(6, 4, 3) = #2/12/1969#
    Arr(7, 4, 3) = CDec(10000000.0587)
    Arr(8, 4, 3) = CDbl(-123.456)
    
    Arr(1, 5, 3) = Empty
    Arr(2, 5, 3) = CVErr(xlErrNA)
    Arr(3, 5, 3) = CInt(2345.5678)
    Arr(4, 5, 3) = CLng(123456789)
    Arr(5, 5, 3) = Null
    Arr(6, 5, 3) = CSng(654.321)
    Arr(7, 5, 3) = "abc"
    
    'Act:
    If Not modArraySupport2.ResetVariantArrayToDefaults(Arr) Then _
            GoTo TestFail
    
    'Assert:
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            For k = LBound(Arr, 3) To UBound(Arr, 3)
                If IsObject(Arr(i, j, k)) Then
                    Assert.IsNothing Arr(i, j, k)
                Else
                    Assert.AreEqual aExpected(i, j, k), Arr(i, j, k)
                End If
            Next
        Next
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
