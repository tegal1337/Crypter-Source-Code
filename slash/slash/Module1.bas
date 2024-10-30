Attribute VB_Name = "Module1"
'foreground
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


 'Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
Public FDSize As String
Public MeltStub As Integer, HiddenStub As Integer, UseRC4 As Integer

Public Sub FILESSIZE()

End Sub

