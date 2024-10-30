Attribute VB_Name = "mEncrypt"
Function XOREncryption(CodeKey As String, DataIn As String) As String
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer, intXOrValue2 As Integer

    For lonDataPtr = 1 To Len(DataIn)
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        strDataOut = strDataOut + Chr$(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   XOREncryption = strDataOut
End Function
