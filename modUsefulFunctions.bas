Attribute VB_Name = "modUsefulFunctions"

Option Explicit
Option Base 1


'to make 'vbLongLong' work also on 32-bit systems
'(inspired by <https://stackoverflow.com/a/36967283>)
Public Function DeclareLongLong() As Byte

    '===============================
    Const vbLongLong As Byte = 20
    '===============================
    
    
    #If Win32 Then
        DeclareLongLong = vbLongLong
    #End If
    #If Win64 Then
        DeclareLongLong = VBA.vbLongLong
    #End If

End Function
