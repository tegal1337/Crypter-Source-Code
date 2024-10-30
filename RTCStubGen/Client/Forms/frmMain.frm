VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GraphicCrypter"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Generate"
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Delimiter:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdProtect 
      Caption         =   "Build"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProtect_Click()

Dim Stub As String

With CD
        .DialogTitle = "Select stub"
        .Filter = "EXE Files |*.exe"
        .ShowSave

End With

Open CD.FileName For Binary As #1
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


File = RC4(File, Text1.Text)

Open CD.FileName For Binary As #1
Put #1, , Stub & Text2.Text & File
Close #1

MsgBox "Hey All is Done ;)", vbInformation

End Sub

Private Sub cmdSearch_Click()

'Now we are on the Sub cmdSearch if for Select The file that we want to Protect

'Lets code

With CD 'With The reference of Command Dialog
        .DialogTitle = "Select The file you Want to Protect" ' Is for The Command Dialog Title
        .Filter = "EXE Files |*.exe" ' We Filter for EXE Files only
        .ShowOpen ' Show Dialog
End With ' Close reference

If Not CD.FileName = vbNullString Then  ' If The Client Select a File

txtfile.Text = CD.FileName ' TXTFILENAME = The Path of The file that we want Protect

End If


'Ok Lets see
'ok All is Fine now go to Sub cmdProtect

End Sub

Public Function RC4(ByVal Data As String, ByVal Password As String) As String ' This is a Modified RC4 Function ^^
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


Public Function RandomLetter() As String
  RandomLetter = ""
  Dim Keyset As String
  Keyset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Anfang:
  Randomize
  var1 = Int(26 * Rnd)
  If var1 = 0 Then GoTo Anfang
  RandomLetter = Mid(Keyset, var1, 1)
End Function
Public Function RandomNumber() As String
  RandomNumber = ""
als:
  Randomize
  var1 = Int(9 * Rnd)
  RandomNumber = var1
If RandomNumber = "0" Then GoTo als
End Function

Private Sub Command1_Click()
Text1.Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber
Text2.Text = "###" & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & "###"

End Sub
