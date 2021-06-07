Attribute VB_Name = "CopyArrayTest"

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
'unit tests for 'CopyArray'
'==============================================================================

'@TestMethod("CopyArray")
Public Sub CopyArray_UnallocatedSrc_ResultsTrueAndUnchangedDest()
    On Error GoTo TestFail

    Dim Src() As Long
    Dim Dest(0) As Integer
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    
    Dim aExpected(0) As Integer
        aExpected(0) = 50
    '==========================================================================
    
    
    'Arrange:
    Dest(0) = 50
    
    'Act:
    If Not modArraySupport2.CopyArray( _
            Src, _
            Dest, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Dest

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyArray")
Public Sub CopyArray_IncompatibleDest_ResultsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim Src(1 To 2) As Long
    Dim Dest(1 To 2) As Integer
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    '==========================================================================
    
    
    'Act:
    'Assert:
    Assert.IsFalse modArraySupport2.CopyArray( _
            Src, _
            Dest, _
            CompatibilityCheck _
    )

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyArray")
Public Sub CopyArray_AllocatedDestLessElementsThenSrc_ResultsTrueAndDestArray()
    On Error GoTo TestFail

    Dim Src(1 To 3) As Long
    Dim Dest(10 To 11) As Long
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    
    Dim aExpected(10 To 11) As Long
        aExpected(10) = 1
        aExpected(11) = 2
    '==========================================================================
    
    
    'Arrange:
    Src(1) = 1
    Src(2) = 2
    Src(3) = 3
    
    'Act:
    If Not modArraySupport2.CopyArray( _
            Src, _
            Dest, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Dest

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyArray")
Public Sub CopyArray_AllocatedDestMoreElementsThenSrc_ResultsTrueAndDestArray()
    On Error GoTo TestFail

    Dim Src(1 To 3) As Long
    Dim Dest(10 To 13) As Long
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = True
    
    Dim aExpected(10 To 13) As Long
        aExpected(10) = 1
        aExpected(11) = 2
        aExpected(12) = 3
        aExpected(13) = 0
    '==========================================================================
    
    
    'Arrange:
    Src(1) = 1
    Src(2) = 2
    Src(3) = 3
    
    'Act:
    If Not modArraySupport2.CopyArray( _
            Src, _
            Dest, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Dest

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("CopyArray")
Public Sub CopyArray_NoCompatibilityCheck_ResultsTrueAndDestArrayWithOverflow()
    On Error GoTo TestFail

    Dim Src(1 To 2) As Long
    Dim Dest(1 To 2) As Integer
    
    '==========================================================================
    Const CompatibilityCheck As Boolean = False
    
    Dim aExpected(1 To 2) As Integer
        aExpected(1) = 1234
        aExpected(2) = 0
    '==========================================================================
    
    
    'Arrange:
    Src(1) = 1234
    Src(2) = 32768       'no valid Integer
    
    'Act:
    If Not modArraySupport2.CopyArray( _
            Src, _
            Dest, _
            CompatibilityCheck _
    ) Then _
            GoTo TestFail
    
    'Assert:
    Assert.SequenceEquals aExpected, Dest

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'TODO: Add tests with Objects
