Attribute VB_Name = "mChangeIcon"

Public Sub ExtractIcon(InFile As String)
    Call ResToFile("Reshacker", "src")
    
       Shell (Environ$("Temp") & "\Src.exe -extract " & InFile & "," & Environ$("Temp") & "\ico.rc" & ",ICONGROUP,,")
       
End Sub
Public Sub ResToFile(ByVal ResID As String, FileName As String)

Dim ByteRA() As Byte
ByteRA = LoadResData(ResID, "custom")

Open Environ$("Temp") & "\" & FileName & ".exe" For Binary As #1
Put #1, , ByteRA
Close #1
End Sub

