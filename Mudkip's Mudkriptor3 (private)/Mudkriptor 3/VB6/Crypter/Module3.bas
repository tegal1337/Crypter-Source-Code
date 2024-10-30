Attribute VB_Name = "Module3"
Public Function changeIcon(file As String, icon As String)
    Shell "ResHacker.exe -addoverwrite " & file & "," & file & ", " & icon & ",ICONGROUP,REBOL,1033"
End Function
