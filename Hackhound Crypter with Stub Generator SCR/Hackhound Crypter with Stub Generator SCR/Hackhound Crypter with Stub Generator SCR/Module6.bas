Attribute VB_Name = "mEof"
Public Function GetEOFData(sFilePath As String) As String
    On Error GoTo ErrHandler
    Dim sFileBuffer                 As String
    Dim sEOFBuffer                  As String
    Dim lPos                        As Long
    
    Open sFilePath For Binary As #1
        sFileBuffer = Space(LOF(1))
        Get #1, , sFileBuffer
    Close #1
    
    lPos = InStr(1, StrReverse(sFileBuffer), GetNullBytes(30))
    sEOFBuffer = (Mid(StrReverse(sFileBuffer), 1, lPos - 1))
    GetEOFData = StrReverse(sEOFBuffer)
    Exit Function
    
ErrHandler:
        GetEOFData = vbNullString
End Function

Public Sub WriteEOFData(sFilePath As String, sEOFData As String)
    On Error Resume Next
    Dim sFile           As String
    Dim lFF             As Long
    
    lFF = FreeFile
    
    Open sFilePath For Binary As #lFF
        sFile = Space(LOF(lFF))
        Get #lFF, , sFile
    Close #lFF
    
    Kill sFilePath
    lFF = FreeFile
    
    Open sFilePath For Binary As #lFF
        Put #lFF, , sFile & sEOFData
    Close #lFF
End Sub

Private Function GetNullBytes(lNum) As String
    Dim sBuffer         As String
    Dim i               As Integer
    
    For i = 1 To lNum
        sBuffer = sBuffer & Chr(0)
    Next i
    
    GetNullBytes = sBuffer
End Function
