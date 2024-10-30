Attribute VB_Name = "mRandomize"
Public Function GenNumKey(mNumber As Integer, Optional MinNumber As Integer) As String

Dim I As Integer, gString As String

Call Randomize

DoEvents

RetryString:
Randomize

gString = ""

For I = 1 To RandomNumber(mNumber, MinNumber)

DoEvents

If I = 1 Or I = 5 Or I = 6 Or I = 10 Or I = 11 Or I = 12 Or I = 18 Or I = 19 Or I = 25 Or I = 27 Or I = 29 Or I = 33 Or I = 35 Or I = 36 Or I = 38 Or I = 40 Or I = 43 Then
    gString = gString & RandomLetter
Else
    gString = gString & RandomNumber(100, 5)
End If

Next I

If Len(gString) < 4 Then GoTo RetryString
GenNumKey = gString
End Function

Public Function RandomNumber(mNumber As Integer, Optional lVal As Integer) As Integer
Dim Mynum As Integer

DoEvents
Randomize

Begin:
Randomize

Mynum = Int(mNumber * Rnd)

If Mynum = 0 Then GoTo Begin

If IsNumeric(lVal) Then
    If lVal <> 0 Then If Mynum <= lVal Then GoTo Begin
End If

RandomNumber = Mynum
End Function

Public Function RandomLetter() As String
Dim KeySet As String, BERJh3uwieojrqu3 As Integer

DoEvents

KeySet = ""

RetryString:

    KeySet = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
    Randomize

BERJh3uwieojrqu3 = Int(52 * Rnd)
If BERJh3uwieojrqu3 = 0 Then GoTo RetryString
RandomLetter = Mid(KeySet, BERJh3uwieojrqu3, 1)
End Function
