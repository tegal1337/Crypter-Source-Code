Attribute VB_Name = "Iconos"
Option Explicit

Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const ILD_TRANSPARENT = &H1
Public Const MAX_PATH = 260

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long



