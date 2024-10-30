VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Fake Error"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Ilykriptor.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmError.frx":0000
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
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2600
      Left            =   0
      Picture         =   "frmError.frx":001C
      ScaleHeight     =   2565
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CheckBox chkError 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox fakeError 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   960
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
    If chkError.Value = 1 Then
        frmMain.fakeError = True
        frmMain.errormsg = fakeError.Text
    Else
        frmMain.fakeError = False
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    If frmMain.fakeError = True Then
        fakeError.Text = frmMain.errormsg
        chkError.Value = 1
    End If
End Sub
