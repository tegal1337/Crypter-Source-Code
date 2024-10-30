VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "iCrypt"
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   5040
      Picture         =   "Form1.frx":2810A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   4800
      Width           =   510
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5E7C5&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2895
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00936337&
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   4680
      Width           =   210
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00936337&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4680
      Width           =   210
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F5E7C5&
      Height          =   1695
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "Form1.frx":50214
      Top             =   -3960
      Width           =   5295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00936337&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   210
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00936337&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   5100
      Width           =   210
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5E7C5&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Chg Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded By omc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4440
      TabIndex        =   18
      Top             =   6000
      Width           =   1185
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Encryption Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   255
      TabIndex        =   16
      Top             =   3540
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Anti Methods"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NT Compression"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Preserve EOF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   5100
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Align PE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   5100
      Width           =   1215
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3840
      TabIndex        =   0
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Image I1 
      Height          =   345
      Left            =   3840
      Picture         =   "Form1.frx":50311
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "iCrypt File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1215
      TabIndex        =   5
      Top             =   5700
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   1215
      Picture         =   "Form1.frx":52407
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label titlu 
      BackStyle       =   0  'Transparent
      Caption         =   "iCrypt v1  -- Public for HackHound.org--                "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   4185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   5220
      TabIndex        =   2
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   195
   End
   Begin VB.Image bd 
      Height          =   345
      Left            =   6090
      Picture         =   "Form1.frx":544FD
      Top             =   4620
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image bo 
      Height          =   345
      Left            =   4560
      Picture         =   "Form1.frx":565F3
      Top             =   -4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image bn 
      Height          =   345
      Left            =   6090
      Picture         =   "Form1.frx":586E9
      Top             =   3750
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   6435
      Left            =   0
      Picture         =   "Form1.frx":5A7DF
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Dim MoveMe As Boolean
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


Private Sub Form_Load()
Me.Show
Invizibile_Form fMain
Text3.Text = RandString(8)
End Sub



Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbLeftButton Then FormDrag Me
 I1.Picture = bn.Picture
  Image2.Picture = bn.Picture
 
End Sub



Private Sub L1_Click()
Dim b1() As Byte, s1 As String
Me.Enabled = False
s1 = GetFileName(, "Executable Files (*.exe)" & Chr(0) & "*.exe" & Chr(0), "Open file to iCrypt!")
Me.Enabled = True
If s1 <> "" Then Text1.Text = s1
If Text1.Text <> "" Then
    Label3.Enabled = True
    
End If
Me.SetFocus
End Sub

Private Sub L1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
I1.Picture = bd.Picture

End Sub

Private Sub L1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
I1.Picture = bo.Picture

End Sub

Private Sub Label1_Click()
End

End Sub

Private Sub Label2_Click()
Me.WindowState = 1

End Sub


Private Sub Label3_Click()
Dim b1() As Byte, b2() As Byte, I1 As Long
Dim s1 As String, s2 As String * 8, s3 As String, s4 As String * 8
Me.Enabled = False
s1 = GetFileName("crypted.exe", "Executable Files *.exe" & Chr(0) & "*.exe", "Save Crypted File As", False)
Me.Enabled = True
If s1 = "" Then Exit Sub
If Not Right(LCase(s1), 4) = ".exe" Then s1 = s1 & ".exe"


s3 = s3 & IIf(Check3.Value = vbChecked, "1", "0") & ","
s3 = s3 & IIf(Check4.Value = vbChecked, "1", "0") & ","
s3 = s3 & IIf(Check1.Value = vbChecked, "1", "0") & ","
s3 = s3 & IIf(Check2.Value = vbChecked, "1", "0") & ","

                vbWriteByteFile s1, LoadResData(101, "CUSTOM")
                b1 = LoadFile(Text1)
                
                If Check3.Value = vbChecked Then
                    b2 = b1
                    's2 = CompressedSize(b2)
                    b2 = CompressData(b1)
                    b1 = b2
                Else
                    's2 = UBound(b1) + 1
                End If
                RC4 b1, Text3.Text
                SetResourceBytes 1000, 1000, b1, s1
                SetResource 1000, 1001, Text3.Text, s1
                SetResource 1000, 1002, UBound(b1) + 1, s1
                SetResource 1000, 1003, IIf(Check1.Value = vbChecked, 1, 0), s1
                SetResource 1000, 1004, IIf(Check2.Value = vbChecked, 1, 0), s1
                SetResource 1000, 1005, IIf(Check3.Value = vbChecked, 1, 0), s1
                SetResource 1000, 1006, IIf(Check4.Value = vbChecked, 1, 0), s1
                
                
                If Check1.Value = vbChecked Then
                    b1 = LoadFile(s1)
                    b2 = EndOfFileByte(Text1.Text)
                    DeleteFile s1
                    vbWriteEOF s1, b1, b2
                    If Check2.Value = vbChecked Then mPE_Realign.RealignPEFromFile s1
                End If
                
                'Open s1 For Binary Access Write As #2
                'b1 = LoadResData(101, "CUSTOM")
                'Put #2, , b1
                'b1 = LoadFile(Text1)
                'S4 = UBound(b1) + 1

                'RC4 b1, Text3.Text
                'Put #2, , b1
                'Put #2, , s4
                'Put #2, , s3
                'Put #2, , s2
                'Put #2, , Text3.Text
                
                'Close #2
                
                '
                
                'Erase b1
                'Erase b2
                'b1 = LoadFile(Text1.Text)
                'I1 = UBound(b1) + 1
                'b2 = CompressData(b1)
                'vbWriteByteFile "compress1.txt", b2
                'B1 = LoadFile("compress1.txt")
                'b2 = DecompressData(b1, I1)
                'vbWriteByteFile "decompress1.txt", b2
                
                
                'b1 = LoadFile(s1)
                'DeleteFile s1
                'vbWriteEOF s1, b1, EndOfFileByte(Text1.Text)
                'mPE_Realign.RealignPEFromFile s1
                Me.SetFocus
                MsgBox "Saved as " & s1, vbInformation, "iCrypt"

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Picture = bd.Picture
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Picture = bo.Picture
End Sub

Private Sub Label4_Click()
Check2.Value = IIf(Check2.Value = vbChecked, vbUnchecked, vbChecked)
End Sub

Private Sub Label5_Click()
Check1.Value = IIf(Check1.Value = vbChecked, vbUnchecked, vbChecked)
End Sub

Private Sub Label6_Click()
Check3.Value = IIf(Check3.Value = vbChecked, vbUnchecked, vbChecked)
End Sub

Private Sub Label7_Click()
Check4.Value = IIf(Check4.Value = vbChecked, vbUnchecked, vbChecked)
End Sub

Private Sub Label9_Click()
Form1.Show
End Sub

Private Sub Text3_Click()
Text3.Text = RandString(Rand(8, 16))
End Sub

Private Sub titlu_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  MoveMe = True
    OldX = x
    OldY = Y
End Sub

Private Sub titlu_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If MoveMe = True Then
       Me.Left = Me.Left + (x - OldX)
      Me.Top = Me.Top + (Y - OldY)
 End If
 
 I1.Picture = bn.Picture
 
End Sub

Private Sub titlu_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Me.Left = Me.Left + (x - OldX)
    Me.Top = Me.Top + (Y - OldY)
    MoveMe = False
End Sub

