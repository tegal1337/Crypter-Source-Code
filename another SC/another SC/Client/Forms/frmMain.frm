VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Runtime Crypter"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4890
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdProtect 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "By Natsu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProtect_Click()

Dim Stub As String



Open App.Path & "\Stub.exe" For Binary As #1
Stub = Space(LOF(1))
Get #1, , Stub
Close #1

With CD

        .DialogTitle = "Select Where you want to Save Crypted File"
        .Filter = "EXE Files |*.exe"
        .ShowSave

End With


Dim File As String

Open txtfile.Text For Binary As #1
File = Space(LOF(1))
Get #1, , File
Close #1




File = RC4(File, "EMOROCK")

Open CD.FileName For Binary As #1
Put #1, , Stub & "=DELIMITER=" & File
Close #1

MsgBox "Encrypt file successfully", vbInformation, Me.Caption
End Sub

Private Sub cmdSearch_Click()

With CD
        .DialogTitle = "Select The file you Want to Protect"
        .Filter = "EXE Files |*.exe"
        .ShowOpen
End With

If Not CD.FileName = vbNullString Then

txtfile.Text = CD.FileName

End If



End Sub

Public Function RC4(ByVal Data As String, ByVal Password As String) As String
On Error Resume Next
Dim F(0 To 255) As Integer, X, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For X = 0 To 255
    Y = (Y + F(X) + Key(X Mod Len(Password))) Mod 256
    F(X) = X
Next X
Key() = StrConv(Data, vbFromUnicode)
For X = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(X) = Key(X) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next X
RC4 = StrConv(Key, vbUnicode)
End Function

