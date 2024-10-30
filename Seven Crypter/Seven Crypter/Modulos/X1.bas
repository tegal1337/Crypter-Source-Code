Attribute VB_Name = "X1"
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Function Archivo_Temporal() As String
    Dim sSave As String, hOrgFile As Long, hNewFile As Long, bBytes() As Byte
    Dim sTemp As String, nSize As Long, ret As Long
    
    sTemp = String(260, 0)

    GetTempFileName Environ("temp"), "TTT", 0, sTemp

    Archivo_Temporal = Left$(sTemp, InStr(1, sTemp, Chr$(0)) - 1)

End Function


Function Cargar(ID As Integer) As String
Path = Archivo_Temporal
Dim aDatos() As Byte
aDatos = LoadResData(ID, "CUSTOM")
Open Path For Binary Access Write As #1
Put #1, , aDatos
Close
Cargar = Path
   
End Function

Public Function DirD(Ruta As String) As Boolean
Dim xFile
Set xFile = CreateObject("Scripting.FileSystemObject")
DirD = xFile.FileExists(Ruta)
End Function



Public Function cambiaricono(Archivo As String, Icono As String)
Shell (App.Path & "\reshacker.exe -addoverwrite " & Archivo & ", " & Archivo & ", " & Icono & ", " & "icongroup, 1,0") '
End Function


Public Function Skinear(Form As String)

frmMain.Skin.LoadSkin Cargar(1), ""
frmMain.Skin.ApplyWindow Form
frmMain.Skin.ApplyOptions = frmMain.Skin.ApplyOptions Or xtpSkinApplyMetrics

End Function

Public Function ReadEOFData(sFilePath As String) As String
On Error GoTo Err:
Dim sFileBuf As String, sEOFBuf As String, sChar As String
Dim lFF As Long, lPos As Long, lPos2 As Long, lCount As Long
If Dir(sFilePath) = "" Then GoTo Err:
lFF = FreeFile
Open sFilePath For Binary As #lFF
sFileBuf = Space(LOF(lFF))
Get #lFF, , sFileBuf
Close #lFF
lPos = InStr(1, StrReverse(sFileBuf), GetNullBytes(30))
sEOFBuf = (Mid(StrReverse(sFileBuf), 1, lPos - 1))
ReadEOFData = StrReverse(sEOFBuf)
If ReadEOFData = "" Then
End If
Exit Function
Err:
ReadEOFData = vbNullString
End Function

Sub WriteEOFData(sFilePath As String, sEOFData As String)
Dim sFileBuf As String
Dim lFF As Long
On Error Resume Next
If Dir(sFilePath) = "" Then Exit Sub
lFF = FreeFile
Open sFilePath For Binary As #lFF
sFileBuf = Space(LOF(lFF))
Get #lFF, , sFileBuf
Close #lFF
Kill sFilePath
lFF = FreeFile
Open sFilePath For Binary As #lFF
Put #lFF, , sFileBuf & sEOFData
Close #lFF
End Sub

Public Function GetNullBytes(lNum) As String
Dim sBuf As String
Dim i As Integer
For i = 1 To lNum
sBuf = sBuf & Chr(0)
Next
GetNullBytes = sBuf
End Function


