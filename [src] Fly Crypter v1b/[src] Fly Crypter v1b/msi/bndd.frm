VERSION 5.00
Begin VB.Form bndd 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1590
   ClientLeft      =   5745
   ClientTop       =   5340
   ClientWidth     =   2640
   LinkTopic       =   "Form2"
   ScaleHeight     =   1590
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
      Begin MSI.cmd cmd1 
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   14
         TX              =   "Ok"
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
         MICON           =   "bndd.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "bndd.frx":001C
         Left            =   1320
         List            =   "bndd.frx":0029
         TabIndex        =   5
         Text            =   "Explorer"
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Inject File to"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H8000000F&
         Height          =   315
         ItemData        =   "bndd.frx":004E
         Left            =   1320
         List            =   "bndd.frx":005B
         TabIndex        =   3
         Text            =   "Temp"
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Drop file to"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSI.UsrSkin UsrSkin1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3201
   End
End
Attribute VB_Name = "bndd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
  Me.Hide
End Sub
Private Sub Combo1_Change()
  Combo1.Text = drt("Rckn")
End Sub
Private Sub Combo2_Change()
  Combo2.Text = drt("Cvnjmpcp")
End Sub
Private Sub Form_Load()
  UsrSkin1.SkinCaption = drt("@glbMnrgml%q")
  Option1.Value = True
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Drag Me
End Sub
Private Sub UsrSkin1_SkinMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then Drag Me
End Sub
Private Sub UsrSkin1_SkinUnload()
  Unload Me
End Sub
Private Sub Form_Resize()
  UsrSkin1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
