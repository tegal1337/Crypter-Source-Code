Attribute VB_Name = "GetSysIcon"
Option Explicit
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_USEFILEATTRIBUTES = &H10
Private Const SHGFI_LARGEICON = 0
Private Const SHGFI_SMALLICON = 1
Private Const SHGFI_ICON = &H100
Private Const SHGFI_ATTRIBUTES = &H800
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const MAX_PATH = 260
Private Type SHFILEINFO
hIcon As Long
iIcon As Long
dwAttributes As Long
szDisplayName As String * MAX_PATH
szTypeName As String * 80
End Type
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, _
ByVal y As Long, ByVal hIcon As Long) As Long

Public Sub GetIconByExtension(ByVal Extension As String, _
ByVal PB As PictureBox)
Dim shinfo As SHFILEINFO
Dim dwFlags As Long
Dim x As Long
PB.Cls
dwFlags = SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICON Or SHGFI_ICON
x = SHGetFileInfo("C:\x." & Extension, FILE_ATTRIBUTE_NORMAL, _
shinfo, LenB(shinfo), dwFlags)
If shinfo.hIcon Then x = DrawIcon(PB.hdc, 0, 0, shinfo.hIcon)
x = DestroyIcon(shinfo.hIcon)
End Sub

Public Function FunImageExists(pList As ListImages, pKey As String) As Boolean
Dim lFakeKey As Long
On Error GoTo ErrHandler
lFakeKey = pList.Item(pKey).Index
FunImageExists = True
Exit Function
ErrHandler:
Err.Clear
FunImageExists = False
End Function

Public Function SetIcon(Sfile As String, Kis As String)
Dim StrAux As String, lFileExt As String
FrmManager.Picture1.AutoRedraw = True

FrmManager.ImageList1.ListImages.Add , , FrmMain.Icon
StrAux = Sfile 'Dir$("C:\")
Do While StrAux <> vbNullString
lFileExt = Split(StrAux, ".")(1)
If Not FunImageExists(FrmManager.ImageList1.ListImages, lFileExt) Then
GetIconByExtension lFileExt, FrmManager.Picture1
FrmManager.ImageList1.ListImages.Add , lFileExt, FrmManager.Picture1.Image
End If

With FrmManager.lstFiles.ListItems.Add(, , Sfile, , lFileExt)
.SubItems(1) = FormatKB(Kis)
.SubItems(2) = lFileExt
End With

StrAux = ""
Loop
End Function
