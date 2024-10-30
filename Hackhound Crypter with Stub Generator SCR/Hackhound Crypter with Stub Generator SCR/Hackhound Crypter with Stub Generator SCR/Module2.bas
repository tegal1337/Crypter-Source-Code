Attribute VB_Name = "mResource"
Option Explicit
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function UpdateResource1 Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long

Public Function SetResource(lpType As Long, lpID As Long, lpData As String, lpFile As String) As Long
    Dim pReturn As Long
    Dim rPort As Long
    
    pReturn = BeginUpdateResource(lpFile, False)
    If pReturn <> 0 Then
        rPort = UpdateResource(pReturn, lpType, lpID, 1033, ByVal lpData, Len(lpData))
        EndUpdateResource pReturn, False
        If rPort <> 0 Then SetResource = True
    End If
    
End Function

Public Function SetResourceBytes(lpType As Long, lpID As Long, lpData() As Byte, lpFile As String) As Long
    Dim pReturn As Long
    Dim rPort As Long
    Dim nCount As Long
    
    nCount = UBound(lpData) + 1 - LBound(lpData)
    pReturn = BeginUpdateResource(lpFile, False)
    If pReturn <> 0 Then
        rPort = UpdateResource1(pReturn, lpType, lpID, 1033, lpData(LBound(lpData)), nCount)
        EndUpdateResource pReturn, False
        If rPort <> 0 Then SetResourceBytes = True
    End If
    
End Function
