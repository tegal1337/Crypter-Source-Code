Attribute VB_Name = "modMain"
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long

Public Sub AddToRes(sData As String, sFilePath As String, sType As String, sName As String)
Dim lRes As Long

lRes = BeginUpdateResource(sFilePath, False)
Call UpdateResource(lRes, sType, sName, 0, ByVal sData, Len(sData))
Call EndUpdateResource(lRes, False)
End Sub
