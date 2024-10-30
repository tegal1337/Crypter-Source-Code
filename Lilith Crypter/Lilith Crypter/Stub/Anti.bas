Attribute VB_Name = "Anti"
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String) As Long

'Yet Another Anti Sandboxie Method - Visual Basic
'by steve10120
'http://hackhound.org
Public Function Sandboxed() As Boolean
    Dim hMod As Long
    hMod = GetModuleHandle("SbieDll.dll")
    If hMod = 0 Then
        Sandboxed = False
    Else
        Sandboxed = True
    End If
End Function

'Anti-Anubis
Public Function Anubis() As Boolean

Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String
 
Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
Res = GetVolumeInformation(Environ("SystemDrive") & "\", Temp1, Len(Temp1), SerialNum, 0, 0, _
Temp2, Len(Temp2))
 
If SerialNum = 1824245000 Then
    Anubis = True
Else
    Anubis = False
End If

End Function

'Anti-Debuggers (OllyDbG, ..)
Public Function Debugger() As Boolean
    Debugger = Not (OutputDebugString(VarPtr(ByVal "=)")) = 1)
End Function
