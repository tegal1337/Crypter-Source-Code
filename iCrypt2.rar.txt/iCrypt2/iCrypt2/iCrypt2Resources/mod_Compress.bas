Attribute VB_Name = "mod_Compress"
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
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function RtlCompressBuffer Lib "ntdll.dll" (ByVal lpCompressionFormat As Long, lpSourceBuffer As Any, ByVal lpSrcLen As Long, lpDestBuffer As Any, ByVal lpDestLen As Long, ByVal lpUnknown As Long, lpDestSize As Long, lpWorkSpaceBuffer As Any) As Long
Public Declare Function RtlGetCompressionWorkSpaceSize Lib "ntdll.dll" (ByVal lpCompressionFormat As Long, lpUnknown As Long, pNeededBufferSize As Long) As Long
Public Declare Function RtlDecompressBuffer Lib "ntdll.dll" (ByVal lpCompressionFormat As Long, lpDestinationBuffer As Any, ByVal lpDestLen As Long, lpSrcBuffer As Any, ByVal lpSrcLen As Long, lpDestSize As Long) As Long

Public Function CompressData(lpData() As Byte) As Byte()
Dim b1() As Byte, lpTemp1 As Long, dwOutputSize As Long, lpSize As Long, b2(0 To 15) As Byte
Dim lpResult() As Byte
lpSize = UBound(lpData) + 1
b1 = lpData
ZeroMemory b1(0), lpSize
Call RtlGetCompressionWorkSpaceSize(COMPRESSION_FORMAT_LZNT1 Or COMPRESSION_ENGINE_MAXIMUM, lpTemp1, dwOutputSize)
lpTemp1 = 0
RtlCompressBuffer COMPRESSION_FORMAT_LZNT1 Or COMPRESSION_ENGINE_MAXIMUM, lpData(0), lpSize, b1(0), lpSize, 0, lpTemp1, b2(0)

ReDim lpResult(0 To lpTemp1 - 1) As Byte
CopyMemory lpResult(0), b1(0), lpTemp1

CompressData = lpResult
'RtlCompressBuffer(dwCompression, lpMemory, Size, lpOutput, Size, 0, @dwTemp, lpWorkspace);
     
End Function
Public Function CompressInfo(lpData() As Byte, lpSize As Long, lpRatio As Long) As Long
Dim i1 As Long
i1 = UBound(CompressData(lpData))
lpSize = i1 + 1
lpRatio = Round((i1 / UBound(lpData)) * 100, 2)
End Function
Public Function CompressedSize(lpData() As Byte) As Long
CompressedSize = UBound(CompressData(lpData)) + 1
End Function
Public Function CompressedRatio(lpData() As Byte) As Long
CompressedRatio = Round((UBound(CompressData(lpData)) / UBound(lpData)) * 100, 2)
End Function
Public Function DecompressData(lpData() As Byte, lpDecompressedSize As Long) As Byte()
Dim b1() As Byte, lpTemp1 As Long, dwOutputSize As Long, lpSize As Long, b2(0 To 15) As Byte
Dim lpResult() As Byte
lpDecompressedSize = UBound(lpData) + 1
dwOutputSize = lpDecompressedSize * 13
ReDim Preserve b1(0 To dwOutputSize) As Byte
RtlDecompressBuffer COMPRESSION_FORMAT_LZNT1 Or COMPRESSION_ENGINE_MAXIMUM, b1(0), dwOutputSize, lpData(0), lpDecompressedSize, lpTemp1
ReDim lpResult(0 To lpTemp1 - 1) As Byte
CopyMemory lpResult(0), b1(0), lpTemp1
DecompressData = lpResult
End Function

