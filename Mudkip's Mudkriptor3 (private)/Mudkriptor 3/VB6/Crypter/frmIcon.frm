VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIcon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Icon"
   ClientHeight    =   2415
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   4650
   Icon            =   "frmIcon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Use custom icon"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog icodiag 
      Left            =   4080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Ilykriptor.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   2
      TX              =   ""
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIcon.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox iconprev 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Warning - Many common icons can get detected. If your crypted file is detected, try a new icon or no icon at all."
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   3855
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
    On Error Resume Next
    icodiag.Filter = "Icon | *.ico"
    icodiag.InitDir = "/icons"
    icodiag.ShowOpen
    iconprev.Picture = LoadPicture(icodiag.FileName)
    frmMain.iconpath = icodiag.FileName
End Sub

Private Sub Check1_Click()
    If frmMain.iconpath = "" Then Call chameleonButton1_Click
    
    If Check1.Value = 1 Then
        frmMain.changeico = True
    Else
        frmMain.changeico = False
    End If
    
    
    
End Sub
