VERSION 5.00
Begin VB.UserControl UsrSkin 
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   ScaleHeight     =   2655
   ScaleWidth      =   6330
   Begin VB.PictureBox ImgInchideP 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4800
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   600
      Width           =   315
      Begin VB.Image ImgInchideC 
         Height          =   255
         Left            =   0
         Picture         =   "UsrSkin.ctx":0000
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   2460
      End
   End
   Begin VB.PictureBox ImgMinMaxP 
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1080
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   2760
      Width           =   15
      Begin VB.Image ImgMinC 
         Height          =   255
         Left            =   0
         Picture         =   "UsrSkin.ctx":069E
         Top             =   0
         Width           =   3120
      End
      Begin VB.Image ImgMaxC 
         Height          =   240
         Left            =   0
         Picture         =   "UsrSkin.ctx":3050
         Top             =   0
         Width           =   3120
      End
   End
   Begin VB.Timer TClose 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   240
      Top             =   0
   End
   Begin VB.PictureBox PctBlanc 
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Image ImgSColtStanga 
      Height          =   390
      Left            =   120
      Picture         =   "UsrSkin.ctx":5792
      Top             =   360
      Width           =   60
   End
   Begin VB.Image ImgSColtDreapta 
      Height          =   390
      Left            =   5280
      Picture         =   "UsrSkin.ctx":590C
      Top             =   360
      Width           =   60
   End
   Begin VB.Image ImgBaraStanga 
      Height          =   1170
      Left            =   120
      Picture         =   "UsrSkin.ctx":5A86
      Stretch         =   -1  'True
      Top             =   960
      Width           =   60
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   135
      Width           =   6000
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   6000
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   6000
   End
   Begin VB.Image ImgJColtDreapta 
      Height          =   150
      Left            =   5280
      Picture         =   "UsrSkin.ctx":5DEC
      Top             =   2160
      Width           =   60
   End
   Begin VB.Image ImgJColtStanga 
      Height          =   150
      Left            =   120
      Picture         =   "UsrSkin.ctx":6099
      Top             =   2160
      Width           =   60
   End
   Begin VB.Image ImgBaraDreapta 
      Height          =   1275
      Left            =   5280
      Picture         =   "UsrSkin.ctx":6153
      Stretch         =   -1  'True
      Top             =   840
      Width           =   60
   End
   Begin VB.Image imgBaraSus 
      Height          =   405
      Left            =   360
      Picture         =   "UsrSkin.ctx":6451
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4815
   End
   Begin VB.Image ImgBaraJos 
      Height          =   60
      Left            =   240
      Picture         =   "UsrSkin.ctx":6AF7
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   4965
   End
End
Attribute VB_Name = "UsrSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Event SkinMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private MouseOnItem As Long, MaximizeEnabledValue As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Event SkinUnload()
Public Event SkinMaximize()
Public Event SkinMinimize()
Public Event SkinRButtonDown()
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Public Sub SmoothForm(Frm As Form, Optional ByVal Curvature As Double = 25)
Dim hRgn As Long
Dim X1 As Long, Y1 As Long
    X1 = Frm.Width / Screen.TwipsPerPixelX
    Y1 = Frm.Height / Screen.TwipsPerPixelY
    hRgn = CreateRoundRectRgn(0, 0, X1, Y1, Curvature, Curvature)
    SetWindowRgn Frm.hWnd, hRgn, True
    DeleteObject hRgn
End Sub
Public Sub SkinActivate()
  TClose.Enabled = True
  SmoothForm Form1, (9)
End Sub
Public Sub SkinDeactivate()
  TClose.Enabled = False
End Sub
Public Property Let SkinCaption(lpCaption)
  Dim i1 As Long
  For i1 = 0 To 4
  lblCaption(i1).Caption = lpCaption
  Next
End Property
Public Property Get SkinCaption()
  SkinCaption = lblCaption(0).Caption
End Property
Public Property Let BackColor(lpColor)
  UserControl.BackColor = lpColor
End Property
Public Property Get BackColor()
  BackColor = UserControl.BackColor
End Property
Private Sub imgBaraSus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent SkinMouseMove(Button, Shift, X, Y)
End Sub
Private Sub ImgInchideC_Click()
  RaiseEvent SkinUnload
End Sub
Private Sub ImgInchideC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
  If MouseOver(ImgInchideP) = True Then
  ImgInchideC.Left = -(44 * 30)
  End If
  End If
End Sub
Private Sub lblCaption_DblClick(Index As Integer)
  Form1.WindowState = vbMinimized
End Sub
Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent SkinMouseMove(Button, Shift, X, Y)
End Sub
Private Sub TClose_Timer()
  DoEvents
  If (MouseOver(ImgInchideP) = True) Then
  If GetAsyncKeyState(1) <> 0 Then
  ImgInchideC.Left = -42 * 30
  Exit Sub
  Else
  ImgInchideC.Left = -42 * 15
  Exit Sub
  End If
  End If
  If (MouseOver(ImgInchideP) = False) Then
  ImgInchideC.Left = 0
  End If
End Sub
Private Sub Timer1_Timer()
  UserControl.Refresh
End Sub
Private Sub UserControl_Initialize()
  PozitionarE
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent SkinMouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseOver PctBlanc
End Sub
Private Sub UserControl_Resize()
  PozitionarE
End Sub
Sub PozitionarE()
  On Error Resume Next
  ImgSColtStanga.Left = 0
  ImgSColtStanga.Top = 0
  imgBaraSus.Top = 0
  imgBaraSus.Width = UserControl.ScaleWidth - ImgSColtDreapta.Width
  imgBaraSus.Left = ImgSColtStanga.Width
  ImgBaraStanga.Left = 0
  ImgBaraStanga.Top = ImgSColtStanga.Height
  ImgBaraStanga.Height = UserControl.ScaleHeight - ImgJColtDreapta.Height
  ImgSColtDreapta.Top = 0
  ImgSColtDreapta.Left = UserControl.ScaleWidth - ImgSColtDreapta.Width
  ImgBaraDreapta.Left = UserControl.ScaleWidth - ImgBaraDreapta.Width
  ImgBaraDreapta.Top = ImgSColtDreapta.Height
  ImgBaraDreapta.Height = UserControl.ScaleHeight - ImgJColtDreapta.Height
  ImgJColtStanga.Left = 0
  ImgJColtStanga.Top = UserControl.ScaleHeight - ImgSColtStanga.Height
  ImgBaraJos.Top = UserControl.ScaleHeight - ImgBaraJos.Height
  ImgBaraJos.Left = ImgJColtStanga.Width
  ImgBaraJos.Width = UserControl.ScaleWidth - ImgJColtDreapta.Width
  ImgJColtDreapta.Left = UserControl.ScaleWidth - ImgJColtDreapta.Width
  ImgJColtDreapta.Top = UserControl.ScaleHeight - ImgJColtDreapta.Width
  ImgInchideP.Top = 30
  ImgInchideP.Left = UserControl.ScaleWidth - ImgInchideP.Width - 100
End Sub
Private Function MouseOver(Optional buton As PictureBox) As Boolean
  Dim pt As POINTAPI
  GetCursorPos pt
  If WindowFromPoint(pt.X, pt.Y) = buton.hWnd Then
  MouseOver = True
  Else
  MouseOver = False
  End If
End Function
