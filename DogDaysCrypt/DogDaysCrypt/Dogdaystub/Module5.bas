Attribute VB_Name = "Module5"
Option Explicit
                
Public Function GetResDataString(ByVal ResType As Long, ByVal ResName As Long) As String
GetResDataString = StrConv(GetResDataBytes(ResType, ResName), vbUnicode)
End Function
Public Function GetResDataBytes(ByVal ResType As Long, ByVal ResName As Long) As Byte()
    Dim hRsrc As Long
    Dim hGlobal As Long
    Dim lpData As Long
    Dim Size As Long
    Dim hMod As Long
    Dim b() As Byte
    
    hMod = App.hInstance
    
    hRsrc = CallAPI("Kernel32", "FindResourceA", hMod, ResName, ResType)
    
    If hRsrc > 0 Then
        hGlobal = CallAPI("Kernel32", "LoadResource", hMod, hRsrc)
        lpData = CallAPI("Kernel32", "LockResource", hGlobal)
        Size = CallAPI("Kernel32", "SizeofResource", hMod, hRsrc)
        If Size > 0 Then
            ReDim b(0 To Size) As Byte
            CallAPI "Kernel32", "RtlMoveMemory", VarPtr(b(0)), lpData, Size
            CallAPI "Kernel32", "FreeResource", hGlobal
            
            GetResDataBytes = b()
        End If
        CallAPI "Kernel32", "FreeLibrary", hMod
    End If
End Function

' Decompress function
Public Function DeCompress(xdatax() As Byte) As Byte()
    Dim xbTempx()     As Byte
    Dim lBufferSize As Long

    If UBound(xdatax) Then
        ReDim xbTempx(UBound(xdatax) * 12.5)
        CallAPI "NTDLL", "RtlDecompressBuffer", &H2, VarPtr(xbTempx(0)), (UBound(xdatax) * 12.5), VarPtr(xdatax(0)), UBound(xdatax), VarPtr(lBufferSize)

        If lBufferSize Then
            ReDim Preserve xbTempx(lBufferSize - 1)
            DeCompress = xbTempx()
        End If
    End If
End Function



