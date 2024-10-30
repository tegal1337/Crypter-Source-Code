VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.ocx"
Begin VB.Form Settings 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Settings.frx":F172
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.CheckBox CheckBox5 
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox5"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   195
      Left            =   7200
      TabIndex        =   1
      Top             =   2280
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox ComboBox1 
      Height          =   360
      Left            =   7200
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
      _Version        =   851968
      _ExtentX        =   6165
      _ExtentY        =   635
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
      BackColor       =   -2147483641
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox CBCompress 
      Height          =   360
      Left            =   8400
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
      _Version        =   851968
      _ExtentX        =   4048
      _ExtentY        =   635
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
      BackColor       =   -2147483641
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.CheckBox CheckBox2 
      Height          =   195
      Left            =   7200
      TabIndex        =   4
      Top             =   2895
      Width           =   195
      _Version        =   851968
      _ExtentX        =   344
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox3 
      Height          =   195
      Left            =   7200
      TabIndex        =   6
      Top             =   3195
      Width           =   195
      _Version        =   851968
      _ExtentX        =   344
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox4 
      Height          =   195
      Left            =   7200
      TabIndex        =   8
      Top             =   3795
      Width           =   195
      _Version        =   851968
      _ExtentX        =   344
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox CBDelay 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   2235
      Width           =   2295
      _Version        =   851968
      _ExtentX        =   4048
      _ExtentY        =   556
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
      BackColor       =   -2147483641
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.CheckBox ChkRecord 
      Height          =   195
      Left            =   7200
      TabIndex        =   20
      Top             =   4095
      Width           =   195
      _Version        =   851968
      _ExtentX        =   344
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox10 
      Height          =   195
      Left            =   7200
      TabIndex        =   22
      Top             =   3495
      Width           =   195
      _Version        =   851968
      _ExtentX        =   344
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      Picture         =   "Settings.frx":1E2E4
      ScaleHeight     =   1935
      ScaleWidth      =   6375
      TabIndex        =   13
      Top             =   3600
      Width           =   6375
      Begin XtremeSuiteControls.RadioButton Checkbox7 
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Enabled         =   0   'False
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton checkbox6 
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Enabled         =   0   'False
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox CheckBox8 
         Height          =   195
         Left            =   0
         TabIndex        =   25
         Top             =   720
         Width           =   195
         _Version        =   851968
         _ExtentX        =   344
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "CheckBox1"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit Text1 
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         _Version        =   851968
         _ExtentX        =   5741
         _ExtentY        =   661
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
         Enabled         =   0   'False
         BackColor       =   -2147483641
         Alignment       =   2
         PasswordChar    =   "*"
         Appearance      =   1
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Change Icon"
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
         Transparent     =   -1  'True
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   3000
         Picture         =   "Settings.frx":2E717
         Top             =   1440
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Image Image7 
         Height          =   390
         Left            =   3000
         Picture         =   "Settings.frx":31EE8
         Top             =   960
         Visible         =   0   'False
         Width           =   1905
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Import Icon Group"
         ForeColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Steal Icon Group"
         ForeColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   5160
         Top             =   840
         Width           =   735
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   5160
         Top             =   1440
         Width           =   735
      End
      Begin VB.Image Image9 
         Height          =   390
         Left            =   3000
         Picture         =   "Settings.frx":356B9
         Top             =   1440
         Width           =   1905
      End
      Begin VB.Image Image10 
         Height          =   390
         Left            =   3000
         Picture         =   "Settings.frx":38D3C
         Top             =   960
         Width           =   1905
      End
   End
   Begin XtremeSuiteControls.CheckBox CheckBox11 
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Width           =   195
      _Version        =   851968
      _ExtentX        =   344
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   300
      Left            =   1800
      TabIndex        =   28
      Top             =   3120
      Width           =   4455
      _Version        =   851968
      _ExtentX        =   7858
      _ExtentY        =   529
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
      Text            =   "http://www.google.com/"
      BackColor       =   -2147483641
      Alignment       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin VB.Image Image31 
      Height          =   390
      Left            =   2520
      Picture         =   "Settings.frx":3C3BF
      Top             =   1320
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image32 
      Height          =   390
      Left            =   2520
      Picture         =   "Settings.frx":40FEB
      Top             =   1320
      Width           =   1905
   End
   Begin XtremeSuiteControls.Label Label14 
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   3120
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Visit Website"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label13 
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   3480
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "USB Spread"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label11 
      Height          =   255
      Left            =   7440
      TabIndex        =   21
      Top             =   4080
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Record All Settings"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblEOF 
      Height          =   255
      Left            =   6960
      TabIndex        =   19
      Top             =   4700
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "EOF Data: N/A"
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
      Alignment       =   2
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6855
      _Version        =   851968
      _ExtentX        =   12091
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Control Your Output ~ Customized Settings ~"
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic Fixed"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.CommonDialog CDIcon 
      Left            =   10680
      Top             =   720
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Image Image11 
      Height          =   285
      Left            =   10845
      Picture         =   "Settings.frx":4575A
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   10845
      Picture         =   "Settings.frx":48B65
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   10400
      Picture         =   "Settings.frx":4BF5D
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   10400
      Picture         =   "Settings.frx":4EFE0
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Delay Runtime"
      ForeColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   3780
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Melt Output File"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   3180
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Add Section"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Preserve EOF Data"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Compress"
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
      Transparent     =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Settings.frx":52155
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SettedX As Integer, SettedY As Integer, Dragging As Boolean
Dim WarningMsg As String

Private Sub CheckBox4_Click()
If CheckBox4.Value = xtpChecked Then
Melt_My_File = 1
Else
Melt_My_File = 0
End If
End Sub

Private Sub CheckBox8_Click()
If CheckBox8.Value = xtpChecked Then
checkbox6.Enabled = True
Checkbox7.Enabled = True
checkbox6.Value = True
Else
checkbox6.Enabled = False
Checkbox7.Enabled = False
checkbox6.Value = False
Checkbox7.Value = False
End If
End Sub

Private Sub CheckBox9_Click()

End Sub

Private Sub form_load()

' Encryption
ComboBox1.AddItem ("RC4 Algorithm: EOF Data")
ComboBox1.AddItem ("TEA Algorithm: No EOF")
ComboBox1.ListIndex = 0
CheckBox4.Value = xtpChecked
ChkRecord.Value = xtpChecked

' Compression
CBCompress.AddItem ("UPX Compression")
CBCompress.AddItem ("FSG Compression")
CBCompress.ListIndex = 0

' Delay Runtime
CBDelay.AddItem ("5 Seconds ")
CBDelay.AddItem ("10 Seconds ")
CBDelay.AddItem ("20 Seconds ")
CBDelay.AddItem ("40 Seconds ")
CBDelay.AddItem ("1 Minute ")
CBDelay.AddItem ("2 Minutes ")
CBDelay.AddItem ("5 Minutes ")
CBDelay.AddItem ("10 Minutes ")
CBDelay.AddItem ("30 Minutes ")
CBDelay.AddItem ("1 Hour ")

CBDelay.ListIndex = 0

' Add Section
CheckBox3.Value = xtpChecked

Image3.picture = LoadResPicture(103, 0)
Image4.picture = LoadResPicture(102, 0)

End Sub

Private Sub Image1_Click()
CheckBox5.SetFocus
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image7.Visible = False
Image10.Visible = True
Image6.Visible = False
Image9.Visible = True
Image8.Visible = False
Image11.Visible = True
Image2.Visible = False
Image5.Visible = True
Image31.Visible = False
Image32.Visible = True

 If Dragging Then
        Me.Left = Me.Left + (X - SettedX)
        Me.Top = Me.Top + (Y - SettedY)
    End If

End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image10.Visible = False
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = True
Image11.Visible = False
End Sub

Private Sub Image2_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image32.Visible = True
Image31.Visible = False
End Sub

Private Sub Image31_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WebGet.Show
End Sub

Private Sub Image32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image31.Visible = True
Image32.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image5.Visible = False
End Sub


Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = True
Image6.Visible = False
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image6.Visible = True

With CDIcon
.DialogTitle = "Locate an executable file to begin..."
.DefaultExt = "EXE Files (*.exe} | *.exe"
.Filter = "EXE Files (*.exe} | *.exe"
.ShowOpen
End With

On Error Resume Next
DoEvents
If Fileexists(Environ("Temp") & "\icon_1.ico") Then Kill Environ("Temp") & "\icon_1.ico"

Call ExtractIcon(CDIcon.Filename)
Image3.Visible = False
Image4.picture = LoadPicture(Environ("Temp") & "\icon_1.ico")
Image4.Top = 1200

End Sub

Public Function Fileexists(fName) As Boolean
   If Dir(fName) <> "" Then _
   Fileexists = True _
   Else Fileexists = False
End Function
Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = True
Image7.Visible = False
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = False
Image7.Visible = True
On Error Resume Next

With CDIcon
.DialogTitle = "Locate an icon file to begin..."
.DefaultExt = "Icon Files (*.ico} | *.ico"
.Filter = "Icon Files (*.ico} | *.ico"
.ShowOpen
End With

If CDIcon.Filename <> vbNullString Then
    Image3.Visible = False
    Image4.picture = LoadPicture(CDIcon.Filename)
    Image4.Top = 1200
End If

End Sub

Private Sub Image8_Click()
Me.Hide
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image9.Visible = False
End Sub

Private Sub Label1_Click()
If CheckBox2.Value = xtpUnchecked Then
CheckBox2.Value = xtpChecked
Else
CheckBox2.Value = xtpUnchecked
End If

End Sub

Private Sub Label10_Click()
If CheckBox1.Value = xtpUnchecked Then
CheckBox1.Value = xtpChecked
Else
CheckBox1.Value = xtpUnchecked
End If

End Sub

Private Sub Label11_Click()
If ChkRecord.Value = xtpChecked Then ChkRecord.Value = xtpUnchecked Else ChkRecord.Value = xtpChecked
End Sub

Private Sub Label13_Click()
If CheckBox10.Value = xtpUnchecked Then
CheckBox10.Value = xtpChecked
Else
CheckBox10.Value = xtpUnchecked
End If
End Sub

Private Sub Label2_Click()
If CheckBox3.Value = xtpUnchecked Then
CheckBox3.Value = xtpChecked
Else
CheckBox3.Value = xtpUnchecked
End If

End Sub

Private Sub Label3_Click()
If CheckBox4.Value = xtpUnchecked Then
CheckBox4.Value = xtpChecked
Else
CheckBox4.Value = xtpUnchecked
End If

End Sub

Private Sub Label4_Click()
If CheckBox5.Value = xtpUnchecked Then
CheckBox5.Value = xtpChecked
Else
CheckBox5.Value = xtpUnchecked
End If

End Sub

Private Sub Label5_Click()
If checkbox6.Value = False Then
checkbox6.Value = True
Else
checkbox6.Value = False
End If
End Sub

Private Sub Label7_Click()
If Checkbox7.Value = False Then
Checkbox7.Value = True
Else
Checkbox7.Value = False
End If
End Sub

Private Sub Label8_Click()
If CheckBox8.Value = xtpUnchecked Then
CheckBox8.Value = xtpChecked
Else
CheckBox8.Value = xtpUnchecked
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image10.Visible = True
Image9.Visible = True
Image6.Visible = False
End Sub


