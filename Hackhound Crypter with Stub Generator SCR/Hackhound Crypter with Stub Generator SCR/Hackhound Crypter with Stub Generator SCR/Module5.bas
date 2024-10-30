Attribute VB_Name = "mCompress"
Option Explicit
Private Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" (ByVal Flags As Integer, WorkSpaceSize As Long, UNKNOWN_PARAMETER As Long) As Long
Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, ByVal NumBits As Long, regionsize As Long, ByVal Flags As Long, ByVal ProtectMode As Long) As Long
Private Declare Function RtlCompressBuffer Lib "NTDLL" (ByVal Flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, ByVal UNKNOWN_PARAMETER As Long, OutputSize As Long, ByVal WorkSpace As Long) As Long
Private Declare Function RtlDecompressBuffer Lib "NTDLL" (ByVal Flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, OutputSize As Long) As Long
Private Declare Function NtFreeVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, regionsize As Long, ByVal Flags As Long) As Long
 
Public Function Compress(Data() As Byte) As Byte()
    Dim WorkSpaceSize As Long
    Dim WorkSpace As Long
    Dim bTemp() As Byte
    Dim lOut As Long
    
    ReDim bTemp(UBound(Data) * 1.13 + 4)
 
    RtlGetCompressionWorkSpaceSize 2, WorkSpaceSize, 0
    NtAllocateVirtualMemory -1, WorkSpace, 0, WorkSpaceSize, 4096, 64
    RtlCompressBuffer 2, VarPtr(Data(0)), UBound(Data) + 1, VarPtr(bTemp(0)), (UBound(Data) * 1.13 + 4), 0, lOut, WorkSpace
    NtFreeVirtualMemory -1, WorkSpace, 0, 16384
    ReDim Preserve bTemp(lOut)
    Compress = bTemp()
 
End Function

