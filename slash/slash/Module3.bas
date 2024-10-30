Attribute VB_Name = "Module3"
Option Explicit

Declare Function GetWindowLong Lib "user32.dll" _
                 Alias "GetWindowLongA" ( _
                 ByVal hWnd As Long, _
                 ByVal nIndex As Long) As Long
                 
Declare Function SetWindowLong Lib "user32.dll" _
                 Alias "SetWindowLongA" ( _
                 ByVal hWnd As Long, _
                 ByVal nIndex As Long, _
                 ByVal dwNewLong As Long) As Long
                 
Declare Function SetLayeredWindowAttributes Lib "user32.dll" ( _
                 ByVal hWnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

'Macht nur eine Farbe transparent
'Public Const LWA_COLORKEY = &H1

'Macht das ganze Fenster transparent
Public Const LWA_ALPHA = &H2

Public Sub Mache_Transparent(hWnd As Long, Rate As Byte)
    '### funktioniert nur unter Windows 2000 oder XP!!!
    '### macht das Fenster, dessen hWnd übergeben wurde, transparent
    '### Rate: 254 = normal 0 = ganz transparent (also unsichtbar)
    '### 190 ist z.B. ein guter Wert
    
    Dim WinInfo As Long
    
    WinInfo = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If Rate < 255 Then
        WinInfo = WinInfo Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, WinInfo
        SetLayeredWindowAttributes hWnd, 0, Rate, LWA_ALPHA
    Else
        'Wenn als Rate 255 angegeben wird,
        'so wird der Ausgangszustand wiederhergestellt
        WinInfo = WinInfo Xor WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, WinInfo
    End If
End Sub





