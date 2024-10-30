Attribute VB_Name = "mResource"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Public Function STRING_TO_BYTES(sString As String) As Byte()
STRING_TO_BYTES = StrConv(sString, vbFromUnicode)
End Function

Public Function BYTES_TO_STRING(bBytes() As Byte) As String
BYTES_TO_STRING = bBytes
BYTES_TO_STRING = StrConv(BYTES_TO_STRING, vbUnicode)
End Function

Public Function GetResData(ByVal ResType As Long, ByVal ResName As Long, Optional EXEPfad As String) As String
GetResData = BYTES_TO_STRING(GetResDataBytes(ResType, ResName, EXEPfad))
End Function

Public Function GetResDataBytes(ByVal ResType As Long, ByVal ResName As Long, Optional EXEPfad As String) As Byte()
Dim hMod As Long
Dim Text As String
Dim hRsrc As Long
Dim b() As Byte
Dim lpData As Long
Dim Size As Long
Dim hGlobal As Long

If EXEPfad = "" Or EXEPfad Like AppExe Or Dir(EXEPfad) = "" Then
hMod = App.hInstance
Else
hMod = LoadLibrary(EXEPfad)
End If

If hMod = 0 Then Exit Function
If IsNumeric(CLng(ResType)) Then hRsrc = CallAPIByName("kernel32", "FindResourceA", CLng(hMod), ResName, CLng(ResType))
If hRsrc = 0 Then Exit Function

hGlobal = CallAPIByName("kernel32", "LoadResource", hMod, hRsrc)
lpData = CallAPIByName("kernel32", "LockResource", hGlobal)
Size = CallAPIByName("kernel32", "SizeofResource", hMod, hRsrc)

If Size = 0 Then Exit Function
Text = Space(Size)
ReDim b(0 To Size - 1) As Byte

Call CopyMemory(b(0), ByVal lpData, Size)
Call CallAPIByName("kernel32", "FreeResource", hGlobal)
GetResDataBytes = b
hMod = CallAPIByName("kernel32", "FreeLibraryA")
End Function

Public Function AppExe() As String
Dim sFile As String * 384
Call GetModuleFileName(GetModuleHandle(vbNullString), sFile, 384)
AppExe = Left(sFile, InStr(1, sFile, Chr(0)) - 1)
End Function


