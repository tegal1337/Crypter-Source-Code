Attribute VB_Name = "modEN"
Function Decry(MyString As String, MyPassword As String) As String
Dim PWMutex
Dim TempString
For i = 1 To Len(MyPassword)
PWMutex = PWMutex & Asc(Mid(MyPassword, i, 1))
Next i
PWMutex = PWMutex - (255 * Fix((PWMutex / 255)))
For i = 1 To Len(MyString)
If (Asc(Mid(MyString, i, 1)) - PWMutex) < 0 Then
TempString = TempString & Chr((Asc(Mid(MyString, i, 1)) - PWMutex) + 255)
Else
TempString = TempString & Chr((Asc(Mid(MyString, i, 1)) - PWMutex))
End If
Next i
Decry = TempString
End Function
