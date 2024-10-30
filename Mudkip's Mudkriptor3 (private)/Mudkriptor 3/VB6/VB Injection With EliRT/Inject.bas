Attribute VB_Name = "Module1"
Option Explicit

Public hModule          As Long
Public hProcess         As Long
Public dwSize           As Long
Public dwPid            As Long
Public dwBytesWritten   As Long
Public dwTid            As Long

Public SE               As SECURITY_ATTRIBUTES

'Some constants arnt needed, but hey :P
Public Const PAGE_READONLY              As Long = &H2
Public Const PAGE_READWRITE             As Long = &H4
Public Const PAGE_EXECUTE               As Long = &H10
Public Const PAGE_EXECUTE_READ          As Long = &H20
Public Const PAGE_EXECUTE_READWRITE     As Long = &H40
Public Const MEM_RELEASE                As Long = &H8000
Public Const MEM_COMMIT                 As Long = &H1000
Public Const MEM_RESERVE                As Long = &H2000
Public Const MEM_RESET                  As Long = &H80000
Public Const STANDARD_RIGHTS_REQUIRED   As Long = &HF0000
Public Const SYNCHRONIZE                As Long = &H100000
Public Const PROCESS_ALL_ACCESS         As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Const INFINITE                   As Long = &HFFFFFF

Public Type SECURITY_ATTRIBUTES
        nLength                 As Long
        lpSecurityDescriptor    As Long
        bInheritHandle          As Long
End Type
'Replace RT.dll with the dropped EliRT Lib Name!
Public Declare Function xVirtualAllocEx Lib "RT.dll" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function xVirtualFreeEx Lib "RT.dll" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function xCreateRemoteThread Lib "RT.dll" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
'Public Declare Function xOpenThread Lib "RT.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Sub Main()
Inject "C:\VIP.dll", "IEFrame"
End Sub

Public Function Inject(szDll As String, szTargetWindowClassName As String) As Boolean
Dim hWnd        As Long
Dim k32LL       As Long
Dim Thread      As Long

    SE.nLength = Len(SE)
    SE.lpSecurityDescriptor = False
    
    'Find window and open process
    hWnd = FindWindow(szTargetWindowClassName, vbNullString)
    GetWindowThreadProcessId hWnd, dwPid
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, dwPid)
        If hProcess = 0 Then GoTo Inject_Error
    k32LL = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
        'MsgBox "Process is: " & hProcess
    
    'Do the actual injecting
    hModule = xVirtualAllocEx(hProcess, 0, LenB(szDll), MEM_COMMIT, PAGE_READWRITE)
        'MsgBox "Module is: " & hModule
        If hModule = 0 Then GoTo Inject_Error
    WriteProcessMemory hProcess, ByVal hModule, ByVal szDll, LenB(szDll), dwBytesWritten
        'MsgBox "Bytes Written: " & dwBytesWritten
    Thread = xCreateRemoteThread(hProcess, SE, 0, ByVal k32LL, ByVal hModule, 0, dwTid)
        If Thread = 0 Then GoTo Inject_Error
        'MsgBox "Thread ID: " & dwTid
        'MsgBox "Thread is: " & Thread
    'Clean up a bit
    WaitForSingleObject Thread, INFINITE
    xVirtualFreeEx hProcess, hModule, 0&, MEM_RELEASE
    CloseHandle Thread

Exit Function

Inject_Error:
    Inject = False
    MsgBox "error"
    Exit Function
End Function

'Delphi and MASM examples

'function InjectLibrary(Process: LongWord; DLLPath: PChar): Boolean;
'Var
'  Parameters: Pointer;
'  BytesWritten, Thread, ThreadID: dword;
'begin
'  Result := False;
'  Parameters := xVirtualAllocEx(Process, nil, 4096, MEM_COMMIT, PAGE_READWRITE);
'  if Parameters = nil then Exit;
'  WriteProcessMemory(Process, Parameters, Pointer(DLLPath), 4096, BytesWritten);
'  Thread := xCreateRemoteThread(Process, nil, 0, GetProcAddress(GetModuleHandle('KERNEL32.DLL'), 'LoadLibraryA'), Parameters, 0, @ThreadId);
'  WaitForSingleObject(Thread, INFINITE);
'  xVirtualFreeEx(Process, Parameters, 0, MEM_RELEASE);
'  if Thread = 0 then Exit;
'  CloseHandle(Thread);
'  Result := True;
'end;

'.code
'_entrypoint:
'invoke FindWindow, addr szTarget, 0
'invoke GetWindowThreadProcessId, eax, addr dwPid
'Invoke OpenProcess, PROCESS_ALL_ACCESS, False, dwPid
'mov hProcess, eax
'invoke xVirtualAllocEx, hProcess, 0, sizeof szDll, MEM_COMMIT or MEM_RESERVE, PAGE_EXECUTE_READWRITE
'mov hModule, eax
'invoke WriteProcessMemory, hProcess, hModule, addr szDll, sizeof szDll, addr dwBytesWritten
'invoke GetModuleHandle, addr szKernel32
'invoke GetProcAddress, eax, addr szLoadLibrary
'invoke xCreateRemoteThread, hProcess, 0, 0, eax, hModule, 0, addr dwTid
'Invoke ExitProcess, 0
'end _entrypoint
