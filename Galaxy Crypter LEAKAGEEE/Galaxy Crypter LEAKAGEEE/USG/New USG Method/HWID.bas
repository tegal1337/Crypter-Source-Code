Attribute VB_Name = "HWID"
Dim reg As Object, Pid As Variant, GUID As Variant
Dim LENGUID As Long, LENPID As Long, TempS As String
Dim X As Long, SPID As String, SGUID As String, HWID As String
Const regPID = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductId"
Const regGUID = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography\MachineGuid"

Public Function CREATEID() As String
'On Error Resume Next

Set reg = CreateObject("wscript.shell")
Pid = Replace(reg.regread(regPID), "-", "")
GUID = Replace(reg.regread(regGUID), "-", "")

LENPID = Len(Pid)
LENGUID = Len(GUID)
    
For X = 1 To LENPID
TempS = Hex((Asc(Mid$(Pid, X, 1)) Xor 23) Xor 14)
SPID = SPID & TempS
Next X
SPID = StrReverse(SPID)

For X = 1 To LENGUID
TempS = Hex((Asc(Mid$(GUID, X, 1)) Xor 23) Xor 14)
SGUID = SGUID & TempS
Next X
SGUID = StrReverse(SGUID)
HWID = StrReverse(SGUID & SPID)
CREATEID = HWID
End Function

Public Function dbvgbwdiz(ByVal sData As String) As String
    Dim i       As Long

    For i = 1 To Len(sData)
        dbvgbwdiz = dbvgbwdiz & Chr$(Asc(Mid$(sData, i, 1)) Xor &HFE)
    Next i
End Function
