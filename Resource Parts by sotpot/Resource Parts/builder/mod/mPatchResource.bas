Attribute VB_Name = "mPatchResource"
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long

Public Function PatchResource(ExeName As String, ResName As String, ResType As String, ResData As String, Optional ResLang As Long = 1033)
    Dim pReturn As Long
    pReturn = BeginUpdateResource(ExeName, False)
    UpdateResource pReturn, ResType, ResName, ResLang, ByVal ResData, Len(ResData)
    EndUpdateResource pReturn, False
End Function

