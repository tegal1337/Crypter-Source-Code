VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLoginStore 
      Caption         =   "Save Info"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin DarowsCrypter.jcbutton cmdLogin 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      Caption         =   "Login"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
