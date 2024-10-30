Attribute VB_Name = "modLvTimer"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Function lv_TimerCallBack(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tButtons As lvButtons_H
    CopyMemory tButtons, GetProp(hwnd, "lv_ClassID"), &H4
    Call tButtons.TimerUpdate(GetProp(hwnd, "lv_TimerID"))
    CopyMemory tButtons, 0&, &H4
End Function


