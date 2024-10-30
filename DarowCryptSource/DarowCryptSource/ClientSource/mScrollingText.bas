Attribute VB_Name = "mScrollingText"
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByValcrKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Const LWA_ALPHA = 2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Dim KaydirilacakMetin As String

Dim Mesaj(999) As String
Public Durdur As Boolean


Public Function RunMain(HakkindaResmi As PictureBox, RefreshForm As Form)
Dim GecenYaziZamani As Long
Const KaydirmaHizi As Long = 1
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long
Dim RectHeight As Long
Dim Sayac As Integer

KaydirilacakMetin = "Darow's Crypter 1.0.3" & vbCrLf & _
"Private Edition" & vbCrLf & _
"Coded by Darow in VB6" & vbCrLf & _
"[ Credits/Shoutouts ]" & vbCrLf
Mesaj(1) = "Yin my lover <3" & vbCrLf
Mesaj(2) = "Kyle for idea of song/scrolling text"
Mesaj(3) = "All my customers"
Mesaj(4) = "Pure epicness"

For Sayac = 0 To UBound(Mesaj)
KaydirilacakMetin = KaydirilacakMetin & Mesaj(Sayac) & vbCrLf
Next




RefreshForm.Refresh

rt = DrawText(HakkindaResmi.hdc, KaydirilacakMetin, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then
    Durdur = True
Else
    DrawingRect.Top = HakkindaResmi.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = HakkindaResmi.ScaleWidth
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + HakkindaResmi.ScaleHeight
End If

Do While Not Durdur
    If GetTickCount() - GecenYaziZamani > KaydirmaHizi Then
       
        HakkindaResmi.Cls
        
        DrawText HakkindaResmi.hdc, KaydirilacakMetin, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
    
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        

        If DrawingRect.Top < -(RectHeight) Then
            DrawingRect.Top = HakkindaResmi.ScaleHeight
            DrawingRect.Bottom = RectHeight + HakkindaResmi.ScaleHeight
        End If
        
        HakkindaResmi.Refresh
        
        GecenYaziZamani = GetTickCount()
        
    End If
    
    DoEvents
Loop

Durdur = True
Set RefreshForm = Nothing

End Function




