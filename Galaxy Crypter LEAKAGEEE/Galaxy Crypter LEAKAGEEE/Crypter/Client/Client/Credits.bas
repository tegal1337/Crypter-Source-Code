Attribute VB_Name = "mCredits"
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByValcrKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long

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

Dim CreditsText As String
Dim SndResData() As Byte
Dim Message(50) As String
Public Stop_Play As Boolean

Public Function RunMain(picture As PictureBox, RefreshForm As Form)
Dim PlayFile As String
Dim Text_On_Form As Long
Const ScrollSpeed As Long = 35
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long
Dim RectHeight As Long
Dim i As Integer

If CdPlay = True Then
MP3_Stop "MyAlias"
SndResData = LoadResData("CdzSound", "Custom")

Open Environ$("Temp") & "\CdzSound.mp3" For Binary As #1
Put #1, , SndResData
Close #1

PlayFile = Environ$("Temp") & "\CdzSound.mp3"
Call mciSendString("open " & PlayFile & " type MPEGVideo alias MyAlias", 0, 0, 0)
Call mciSendString("play myalias from 0", 0, 0, 0)

CreditsText = "Galaxy Crypt" & vbCrLf & _
"Private Edition" & vbCrLf & vbCrLf & _
"<<< Brought To You By... >>>" & vbCrLf & vbCrLf & _
"Programmed By Coder's Central" & vbCrLf & _
"Graphics By Spedunkle (Spedunkle@live.com)" & vbCrLf & _
"Compiled at : 31 January, 2010" & vbCrLf & vbCrLf & _
"<<< A Big Thanks To >>>" & vbCrLf & vbCrLf
Message(0) = "0P3R4T0R"
Message(1) = "RTFLOL" & vbNewLine & "Marienjz" & vbCrLf & "Steven" & vbCrLf & "Radek"
Message(2) = "Zero - Icon Changer"
Message(3) = "Xvisceral"
Message(4) = "Sir Cobein"
Message(5) = "Verbal" & vbCrLf & vbCrLf & _
"<<< Contact Information >>>" & vbCrLf & vbCrLf & _
"Coder: Coder's Central" & vbCrLf
Message(6) = "Msn: Hackforumslogs@live.com"
Message(7) = "Hackforums.net: Coder's Central"
Message(8) = "Email: FactualBusiness@gmail.com" & vbCrLf & vbCrLf & _
"<<< All Graphics By: Spedunkle >>>" & vbCrLf
Message(9) = "Email: Spedunkle@img.img"
Message(10) = "Msn: Spedunkle@live.com"
Message(11) = "Hackforums.net: Spedunkle" & vbCrLf & vbCrLf & _
"<<< Pricing >>>" & vbCrLf
Message(12) = "FUD Stub Cost: $5.00 (USD)" & vbCrLf & vbCrLf & _
"<<< Crypter Packages >>>"
Message(13) = "Original Package: 55 USD"
Message(14) = "Three FUD Stubs"
Message(15) = "Lifetime Updates"
Message(16) = "Update Requests (One Per Customer)"
Message(17) = "Full Support Through MSN" & vbCrLf & vbCrLf & _
"Original Package: 55 USD"
Message(18) = "Two Fud Stubs"
Message(19) = "Updates For First Six Months"
Message(20) = "Full Support" & vbCrLf & _
"Basic Package: 40 USD"
Message(21) = "One Fud Stub"
Message(22) = "Updates For First Three Weeks"
Message(23) = "Minimal Support" & vbCrLf
Message(24) = "If You Wish To Donate..."
Message(25) = "PayPal: FactualBusiness@gmail.com"

Credits.Show
For i = 0 To UBound(Message)
CreditsText = CreditsText & Message(i) & vbCrLf
Next i

RefreshForm.Refresh

rt = DrawText(picture.hdc, CreditsText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then
    Stop_Play = True
Else
    DrawingRect.Top = picture.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picture.ScaleWidth
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picture.ScaleHeight
End If

Do While CdPlay = True
    If GetTickCount() - Text_On_Form > ScrollSpeed Then
       
        picture.Cls
        
           
        DrawText picture.hdc, CreditsText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
    
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        

        If DrawingRect.Top < -(RectHeight) Then
            DrawingRect.Top = picture.ScaleHeight
            DrawingRect.Bottom = RectHeight + picture.ScaleHeight
        End If
        
        picture.Refresh
        
        Text_On_Form = GetTickCount()
        
    End If
    
       
    Do While m_cancel = True
    
    DoEvents
        If m_cancel = False Then
    Exit Do
    End If
    
      Loop
       
    DoEvents
  If DrawingRect.Bottom = 470 Then
  Call RunMain(Credits.Picture1, Credits)
  Exit Do
  End If
  
Loop

Else
CdPlay = False
Credits.Visible = False

End If

Set RefreshForm = Nothing
Stop_Play = True
End Function

Private Sub MP3_Stop(ByVal sAlias As String)
mciSendString "stop " & sAlias, 0, 0, 0
mciSendString "close " & sAlias, 0, 0, 0
End Sub
