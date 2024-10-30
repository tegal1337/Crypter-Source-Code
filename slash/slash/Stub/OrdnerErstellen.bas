Attribute VB_Name = "OrdnerErstellen"
'angepasste version für stub.....
Public Sub CreateFolder(NewPath As String)
Dim fs As Object ', MyNewWb As Variant

Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.FolderExists(NewPath) = False Then   ' Prüfen, ob der ordner existiert...
     fs.CreateFolder (NewPath)   ' WEnn nicht, erstellen
    'MsgBox "Folder " & NewFolder & " Created"                                                                  'testing
        Else
    'MsgBox "Folder " & NewFolder & " Exists"                                                                   'testing
    End If
NewFolder = ""
End Sub
