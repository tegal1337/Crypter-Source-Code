Attribute VB_Name = "modSandbox"
Option Explicit
Private Declare Function OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
  ByVal hHandle As Long, _
  ByVal dwMilliseconds As Long) As Long
Public Function Wait(ByVal mSek As Long)
    WaitForSingleObject -1, mSek
End Function
Public Function IsDebuggerActive() As Boolean
IsDebuggerActive = Not (OutputDebugString(VarPtr(ByVal "=)")) = 1)
End Function
Function ZGHUhjbasDhbSDA() As Boolean
Dim SandboxPath As String
Dim SandboxFile As String
Dim BoxedFile As String

SandboxPath = Environ$(Chr$(83) & Chr$(89) & Chr$(83) & Chr$(84) & Chr$(69) & _
              Chr$(77) & Chr$(68) & Chr$(82) & Chr$(73) & Chr$(86) & Chr$(69)) & Chr$(92) & _
              Chr$(83) & Chr$(97) & Chr$(110) & Chr$(100) & Chr$(98) & Chr$(111) & Chr$(120) & _
              Chr$(92) & Environ$(Chr$(85) & Chr$(83) & Chr$(69) & Chr$(82) & Chr$(78) & Chr$(65) & _
              Chr$(77) & Chr$(69)) & Chr$(92) & Chr$(68) & Chr$(101) & Chr$(102) & Chr$(97) & _
              Chr$(117) & Chr$(108) & Chr$(116) & Chr$(66) & Chr$(111) & Chr$(120) & Chr$(92) & _
              Chr$(100) & Chr$(114) & Chr$(105) & Chr$(118) & Chr$(101) & Chr$(92) & _
              Left(Environ$(Chr$(83) & Chr$(89) & Chr$(83) & Chr$(84) & Chr$(69) & Chr$(77) & _
              Chr$(68) & Chr$(82) & Chr$(73) & Chr$(86) & Chr$(69)), 1) & Chr$(92)

SandboxFile = Environ$(Chr$(83) & Chr$(89) & Chr$(83) & Chr$(84) & Chr$(69) & Chr$(77) & Chr$(68) & _
              Chr$(82) & Chr$(73) & Chr$(86) & Chr$(69)) & Chr$(92) & App.ExeName & Chr$(45) & _
              Chr$(66) & Chr$(111) & Chr$(120) & Chr$(101) & Chr$(100) & Chr$(46) & Chr$(116) & _
              Chr$(101) & Chr$(115) & Chr$(116)

BoxedFile = SandboxPath & App.ExeName & Chr$(45) & Chr$(66) & Chr$(111) & Chr$(120) & Chr$(101) & _
            Chr$(100) & Chr$(46) & Chr$(116) & Chr$(101) & Chr$(115) & Chr$(116)

Open SandboxFile For Output As #1
Close #1

    Wait 200

If Not Dir$(BoxedFile) = "" Then
    ZGHUhjbasDhbSDA = True
        Kill BoxedFile
            Else
    ZGHUhjbasDhbSDA = False
End If
End Function
