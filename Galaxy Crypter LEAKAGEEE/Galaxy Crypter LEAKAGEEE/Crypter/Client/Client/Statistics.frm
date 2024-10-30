VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "CODEJO~3.OCX"
Begin VB.Form Statistics 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Statistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   3495
      Left            =   6960
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   6165
      _StockProps     =   77
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Build Statistics:"
      BackColor       =   -2147483641
      MaxLength       =   500
      MultiLine       =   -1  'True
      Appearance      =   1
      UseVisualStyle  =   0   'False
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit Text5 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   6165
      _StockProps     =   77
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Original Statistics:"
      BackColor       =   -2147483641
      MaxLength       =   500
      MultiLine       =   -1  'True
      Appearance      =   1
      UseVisualStyle  =   0   'False
      ShowBorder      =   0   'False
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   10845
      Picture         =   "Statistics.frx":F172
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   10395
      Picture         =   "Statistics.frx":1257D
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   10845
      Picture         =   "Statistics.frx":15600
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image11 
      Height          =   285
      Left            =   10400
      Picture         =   "Statistics.frx":189F8
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Statistics.frx":1BB6D
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SettedX As Integer, SettedY As Integer, Dragging As Boolean

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Image3.Visible = True
Image2.Visible = True
Image11.Visible = False

If Dragging Then
        Me.Left = Me.Left + (X - SettedX)
        Me.Top = Me.Top + (Y - SettedY)
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub Image11_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image11.Visible = True
End Sub

Private Sub Image5_Click()
Me.Hide
End Sub


Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image5.Visible = True
End Sub
