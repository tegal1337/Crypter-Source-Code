Attribute VB_Name = "mSplitParts"
' made by sotpot


Public Function fSplitParts(sString As String, lLength As Long) As String()
 
 Dim lStringSize As Long
 Dim lPartSize As Long
 Dim lStart As Long
 Dim sTempString As String
 
 
    lStringSize = Len(sString)
    lPartSize = lLength
    lStart = 1
    
    
    
     Do While lStringSize > 0
      sTempString = sTempString & Mid$(sString, lStart, lPartSize) & "{Part}"
      lStringSize = lStringSize - lPartSize
      lStart = lStart + lPartSize
     Loop
     
     fSplitParts = Split(sTempString, "{Part}")
    
    
End Function

