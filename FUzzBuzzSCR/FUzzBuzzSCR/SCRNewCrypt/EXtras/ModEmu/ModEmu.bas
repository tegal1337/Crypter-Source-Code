Attribute VB_Name = "Module1"
'------------------------------------------------------------------
'Name: Anti-Emulator
'Coder: ChainCoder / translated by Slayer616
'Usage: If AntiEmulator = True Then End
'Give Credits if you use this code!
'------------------------------------------------------------------
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Function AntiEmulator() As Boolean
Dim TimeNow As Long
Dim TimeAfterSleep As Long
TimeNow = GetTickCount
Sleep 500
TimeAfterSleep = GetTickCount
If TimeAfterSleep - TimeNow < 500 Then
AntiEmulator = True
Else
AntiEmulator = False
End If
End Function
