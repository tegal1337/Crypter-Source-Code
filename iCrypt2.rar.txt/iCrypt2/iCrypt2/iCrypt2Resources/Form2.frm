VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3270
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5E7C5&
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form2.frx":38E7A
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Update"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   3480
      Picture         =   "Form2.frx":3903A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2520
      TabIndex        =   0
      Top             =   2880
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbLeftButton Then FormDrag Me
End Sub

Private Sub Label1_Click()
Form1.Hide
End Sub

