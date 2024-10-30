Attribute VB_Name = "mKillProcess"
Option Explicit
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
'-------------------------------------------------------
Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapShot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapShot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" ( _
ByVal lFlags As Long, lProcessID As Long) As Long

Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long

Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Sub KillProcess(NameProcess As String)
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = 2&
Dim uProcess  As PROCESSENTRY32
Dim RProcessFound As Long
Dim hSnapShot As Long
Dim SzExename As String
Dim ExitCode As Long
Dim MyProcess As Long
Dim AppKill As Boolean
Dim AppCount As Integer
Dim i As Integer
Dim WinDirEnv As String
       
       If NameProcess <> "" Then
          AppCount = 0

          uProcess.dwSize = Len(uProcess)
          hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapShot, uProcess)
 
          Do
            i = InStr(1, uProcess.szExeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szExeFile, i - 1))
            WinDirEnv = Environ("Windir") + "\"
            WinDirEnv = LCase$(WinDirEnv)
       
            If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
               AppCount = AppCount + 1
               MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
               AppKill = TerminateProcess(MyProcess, ExitCode)
               Call CloseHandle(MyProcess)
            End If
            RProcessFound = ProcessNext(hSnapShot, uProcess)
          Loop While RProcessFound
          Call CloseHandle(hSnapShot)
       End If
End Sub

