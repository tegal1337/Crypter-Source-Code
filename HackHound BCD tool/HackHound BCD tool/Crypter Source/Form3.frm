VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   5340
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by Legssmit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1320
      TabIndex        =   16
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "- Noble"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "- Drizzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "- Hakuno"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "- :D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   8610
      Left            =   0
      Picture         =   "Form3.frx":D5B7
      Top             =   360
      Width           =   60
   End
   Begin VB.Image Image6 
      Height          =   60
      Left            =   -360
      Picture         =   "Form3.frx":1070D
      Top             =   5280
      Width           =   11985
   End
   Begin VB.Image Image4 
      Height          =   8610
      Left            =   4920
      Picture         =   "Form3.frx":13924
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   1320
      Picture         =   "Form3.frx":16A31
      Top             =   4680
      Width           =   1965
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "www.hackhound.org"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Dedicated To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "- SWiM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "- Anyone else who helped me :)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "- Slayer616"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "- Steve10120"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "- Cm2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "- Cobein"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Image Image15 
      DragMode        =   1  'Automatic
      Height          =   12000
      Left            =   0
      Picture         =   "Form3.frx":1A6A5
      Top             =   360
      Width           =   12000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "                                 About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   4320
      Picture         =   "Form3.frx":27C5C
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   4320
      Picture         =   "Form3.frx":2AE81
      Top             =   60
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   0
      Picture         =   "Form3.frx":2E09A
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Crypter based off Cobeins Cryptosy
'Edited by legssmit
' Use  : At your own risk
' ' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission FROM COBEIN AND ME (Legssmit).
Private Dragging As Boolean
Private SettedX As Integer, SettedY As Integer
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = True
End Sub

Private Sub Image3_Click()
Unload Form3
End Sub


Private Sub Image8_Click()
Unload Form3
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label12.Visible = True
End Sub

Private Sub Label12_Click()
Unload Form3
End Sub

Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label12.Visible = False
Image3.Visible = False
End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SettedX = x
    SettedY = y
    Dragging = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Dragging Then
        Me.Left = Me.Left + (x - SettedX)
        Me.Top = Me.Top + (y - SettedY)
    End If
Image3.Visible = False
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dragging = False
End Sub

