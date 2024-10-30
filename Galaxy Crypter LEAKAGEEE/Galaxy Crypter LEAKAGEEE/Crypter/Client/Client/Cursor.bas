Attribute VB_Name = "Cursor"
Option Explicit
Private Const OCR_NORMAL = 32512

Private Declare Function LoadCursorFromFile _
                Lib "user32" _
                Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long

Private Declare Function SetSystemCursor _
                Lib "user32" (ByVal hcur As Long, _
                              ByVal id As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long

Private Declare Function CopyIcon Lib "user32" (ByVal hcur As Long) As Long

Dim CurrSysCur As Long
Dim CurHandle  As Long

Public Function ChangeCur(CursorPath As String)
    CurrSysCur = CopyIcon(GetCursor())
    CurHandle = LoadCursorFromFile(CursorPath)
    SetSystemCursor CurHandle, OCR_NORMAL
End Function

Public Function RestoreCur()
    SetSystemCursor CurrSysCur, OCR_NORMAL
End Function




