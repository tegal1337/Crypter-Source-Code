Attribute VB_Name = "Module4"
Option Explicit

'RtlCompressBuffer(

'IN ULONG CompressionFormat,
'IN PVOID SourceBuffer,
'IN ULONG SourceBufferLength,
'OUT PVOID DestinationBuffer,
'IN ULONG DestinationBufferLength,
'IN ULONG Unknown,
'OUT PULONG pDestinationSize,
'IN PVOID WorkspaceBuffer )

'# #define COMPRESSION_FORMAT_NONE (0x0000) // [result:STATUS_INVALID_PARAMETER]
'#define COMPRESSION_FORMAT_DEFAULT (0x0001) // [result:STATUS_INVALID_PARAMETER]
'#define COMPRESSION_FORMAT_LZNT1 (0x0002)
'#define COMPRESSION_FORMAT_NS3 (0x0003) // STATUS_NOT_SUPPORTED
'#define ... // STATUS_NOT_SUPPORTED
'#define COMPRESSION_FORMAT_NS15 (0x000F) // STATUS_NOT_SUPPORTED
'#define COMPRESSION_FORMAT_SPARSE (0x4000) // ??? [result:STATUS_INVALID_PARAMETER]
'# Compression engine. It's level of compression. Higher level means better results, but longer time used for compression process. In NT 4.0 sp6 engines works only with compression (specified in RtlDecompressBuffer are ignored)
'#define COMPRESSION_ENGINE_STANDARD (0x0000) // Standart compression
'#define COMPRESSION_ENGINE_MAXIMUM (0x0100) // Maximum (slowest but better)
'#define COMPRESSION_ENGINE_HIBER (0x0200) // STATUS_NOT_SUPPORTED
Public Const COMPRESSION_FORMAT_NONE As Long = &H0
Public Const COMPRESSION_FORMAT_DEFAULT As Long = &H1
Public Const COMPRESSION_FORMAT_LZNT1 As Long = &H2
Public Const COMPRESSION_FORMAT_NS3 As Long = &H3
Public Const COMPRESSION_FORMAT_NS15 As Long = &HF
Public Const COMPRESSION_FORMAT_SPARSE As Long = &H4000

Public Const COMPRESSION_ENGINE_STANDARD As Long = &H0
Public Const COMPRESSION_ENGINE_MAXIMUM As Long = &H100
Public Const COMPRESSION_ENGINE_HIBER As Long = &H200
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function RtlDecompressBuffer Lib "ntdll.dll" (ByVal lpCompressionFormat As Long, lpDestinationBuffer As Any, ByVal lpDestLen As Long, lpSrcBuffer As Any, ByVal lpSrcLen As Long, lpDestSize As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'OldFileDate

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type


Public Function GetResData(ByVal ResType As Long, ByVal ResName As Long, Optional EXEPfad As String) As String
GetResData = BYTES_TO_STRING(GetResDataBytes(ResType, ResName, EXEPfad))
End Function

Public Function GetResDataBytes(ByVal ResType As Long, ByVal ResName As Long, Optional EXEPfad As String) As Byte()
      Dim hMod As Long
      Dim Text As String
   Dim hRsrc As Long
      Dim B() As Byte
     Dim lpData As Long
   Dim Size As Long
      Dim hGlobal As Long

Dim i1 As Long, i2 As Long
For i1 = 1 To 100
    i2 = 0
    i2 = Timer * i1
Next

   'Die eigene exe ist ja geladen, also ist hMod das InstanceHandle. Wenn eine Exe angegeben wird, kann allerdings jede exe oder dll ausgelesen werden
   If EXEPfad = "" Or EXEPfad Like AppPath Or Dir(EXEPfad) = "" Then
    hMod = App.hInstance
   Else
    hMod = LoadLibrary(EXEPfad)
   End If
   
For i1 = 1 To 100
    i2 = 0
    i2 = Timer * i1
Next
   
   If hMod = 0 Then Exit Function
   'Resource suchen
   'If IsNumeric(ResType) Then hRsrc = FindResourceByNum(hMod, ResName, CLng(ResType))
   If IsNumeric(CLng(ResType)) Then hRsrc = CallAPIByName("kernel32", "FindResourceA", CLng(hMod), ResName, CLng(ResType))
   'If hRsrc = 0 Then hRsrc = FindResource(hMod, ResName, ResType)
   'MsgBox hRsrc
   
For i1 = 1 To 100
    i2 = 0
    i2 = Timer * i1
Next

   If hRsrc = 0 Then Exit Function
   'Resource Laden
hGlobal = CallAPIByName("kernel32", "LoadResource", hMod, hRsrc)
For i1 = 1 To 100
    i2 = 0
    i2 = Timer * i1
Next
   lpData = CallAPIByName("kernel32", "LockResource", hGlobal) 'Pointer zu unseren Daten
   Size = CallAPIByName("kernel32", "SizeofResource", hMod, hRsrc) 'Länge der Daten ermitteln
   
   
For i1 = 1 To 100
    i2 = 0
    i2 = Timer * i1
Next

   If Size = 0 Then Exit Function
   Text = Space(Size) 'Buffer füllen
   ReDim B(0 To Size - 1) As Byte
   
For i1 = 1 To 100
    i2 = 0
    i2 = Timer * i1
Next
   Call CopyMemory(B(0), ByVal lpData, Size)  'und umwandeln
   Call CallAPIByName("kernel32", "FreeResource", hGlobal)
   GetResDataBytes = B
   FreeLibrary hMod
End Function

Public Function DecompressData(lpData() As Byte, lpDecompressedSize As Long) As Byte()
Dim b1() As Byte, lpTemp1 As Long, dwOutputSize As Long
Dim lpResult() As Byte, lpSize As Long, b2(0 To 15) As Byte
lpDecompressedSize = UBound(lpData) + 1
dwOutputSize = lpDecompressedSize * 13
ReDim Preserve b1(0 To dwOutputSize) As Byte
CallAPIByName "ntdll.dll", "RtlDecompressBuffer", COMPRESSION_FORMAT_LZNT1 Or COMPRESSION_ENGINE_MAXIMUM, VarPtr(b1(0)), dwOutputSize, VarPtr(lpData(0)), lpDecompressedSize, VarPtr(lpTemp1)
ReDim lpResult(0 To lpTemp1 - 1) As Byte
CopyMemory lpResult(0), b1(0), lpTemp1
DecompressData = lpResult
End Function


