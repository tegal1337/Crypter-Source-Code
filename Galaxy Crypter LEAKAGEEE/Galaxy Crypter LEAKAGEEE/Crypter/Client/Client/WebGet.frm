VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.ocx"
Begin VB.Form WebGet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   5640
      Picture         =   "WebGet.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   5295
      TabIndex        =   28
      Top             =   2880
      Width           =   5295
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   300
         Left            =   1800
         TabIndex        =   29
         Top             =   1080
         Width           =   3375
         _Version        =   851968
         _ExtentX        =   5953
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
         Text            =   "C:\Windows\System32\Winlogon.exe"
         BackColor       =   -2147483641
         Alignment       =   2
         Appearance      =   1
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton Rb1 
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton Rb2 
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   360
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton Rb3 
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   720
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton Rb4 
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   1080
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton RadioButton5 
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   0
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton RadioButton6 
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   360
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.Label Label26 
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   0
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notepad.exe"
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
      Begin XtremeSuiteControls.Label Label25 
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Taskmgr.exe"
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
      Begin XtremeSuiteControls.Label Label24 
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Default Browser"
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
      Begin XtremeSuiteControls.Label Label23 
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Custom Process"
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
      Begin XtremeSuiteControls.Label Label22 
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   0
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Execute Normally"
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
      Begin XtremeSuiteControls.Label Label21 
         Height          =   255
         Left            =   2760
         TabIndex        =   36
         Top             =   360
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Execute Hidden"
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
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit3 
      Height          =   300
      Left            =   1560
      TabIndex        =   21
      Top             =   4440
      Width           =   615
      _Version        =   851968
      _ExtentX        =   1085
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
      BackColor       =   -2147483641
      Alignment       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4680
      Picture         =   "WebGet.frx":10433
      ScaleHeight     =   375
      ScaleWidth      =   6375
      TabIndex        =   6
      Top             =   2280
      Width           =   6375
      Begin XtremeSuiteControls.RadioButton Checkbox13 
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   120
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton checkbox12 
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   120
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   120
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Execute Downloaded File"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   120
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Inject Downloaded File"
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
   End
   Begin XtremeSuiteControls.FlatEdit Text5 
      Height          =   300
      Left            =   8640
      TabIndex        =   4
      Top             =   5160
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
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
      Text            =   ".exe"
      BackColor       =   -2147483641
      Alignment       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit TxtURL 
      Height          =   300
      Left            =   6480
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
      _Version        =   851968
      _ExtentX        =   7646
      _ExtentY        =   520
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
      Text            =   "http://rs498tl.rapidshare.com/cgi-bin/upload.cgi?rsuploadid=124610267242291886"
      BackColor       =   -2147483641
      Alignment       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   300
      Left            =   5760
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
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
      Text            =   "Testing"
      BackColor       =   -2147483641
      Alignment       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox CheckBox9 
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Top             =   1500
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox5"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.RadioButton o1 
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   2520
      Width           =   255
      _Version        =   851968
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "RadioButton1"
      BackColor       =   -2147483641
      Appearance      =   2
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton o3 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   3000
      Width           =   255
      _Version        =   851968
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "RadioButton1"
      BackColor       =   -2147483641
      Appearance      =   2
   End
   Begin XtremeSuiteControls.RadioButton o4 
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   3240
      Width           =   255
      _Version        =   851968
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "RadioButton1"
      BackColor       =   -2147483641
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox Com1 
      Height          =   360
      Left            =   2280
      TabIndex        =   23
      Top             =   4440
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
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
      Appearance      =   1
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
      DropDownItemCount=   1
   End
   Begin XtremeSuiteControls.RadioButton o2 
      Height          =   255
      Left            =   960
      TabIndex        =   24
      Top             =   2760
      Width           =   255
      _Version        =   851968
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "RadioButton1"
      BackColor       =   -2147483641
      Appearance      =   2
   End
   Begin XtremeSuiteControls.RadioButton o5 
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   3480
      Width           =   255
      _Version        =   851968
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "RadioButton1"
      BackColor       =   -2147483641
      Appearance      =   2
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   10400
      Picture         =   "WebGet.frx":20866
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image11 
      Height          =   285
      Left            =   10845
      Picture         =   "WebGet.frx":238E9
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   10400
      Picture         =   "WebGet.frx":26CF4
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   10845
      Picture         =   "WebGet.frx":29E69
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin XtremeSuiteControls.Label Label20 
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   3480
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "App Data"
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
   Begin XtremeSuiteControls.Label Label19 
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   2760
      Width           =   1695
      _Version        =   851968
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Program Files"
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
   Begin XtremeSuiteControls.Label Label18 
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   4440
      Width           =   495
      _Version        =   851968
      _ExtentX        =   873
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Time:"
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
   Begin XtremeSuiteControls.Label Label17 
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   3960
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Delay Run Time"
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
   Begin XtremeSuiteControls.Label Label16 
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   2160
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Download Directory:"
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
   Begin XtremeSuiteControls.Label Label11 
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   1560
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Additional WebGet Settings"
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
   Begin XtremeSuiteControls.Label Label15 
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "System32"
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
   Begin XtremeSuiteControls.Label Label14 
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   2520
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Temp Directory"
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
   Begin XtremeSuiteControls.Label Label13 
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "System Root"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   6105
      TabIndex        =   5
      Top             =   4875
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Filename After Download"
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
   Begin XtremeSuiteControls.Label Label9 
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   1485
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Download File"
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
   Begin XtremeSuiteControls.Label Label12 
      Height          =   255
      Left            =   8685
      TabIndex        =   1
      Top             =   4875
      Width           =   855
      _Version        =   851968
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Extension"
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
      Picture         =   "WebGet.frx":2D261
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "WebGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SettedX As Integer, SettedY As Integer, Dragging As Boolean


Private Sub checkbox12_Click()
    
    Rb1.Enabled = True
    Rb2.Enabled = True
    Rb3.Enabled = True
    Rb4.Enabled = True
    FlatEdit2.Enabled = True
    
    RadioButton5.Enabled = False
    RadioButton6.Enabled = False
    Rb1.Value = True

End Sub

Private Sub Checkbox13_Click()


    RadioButton5.Enabled = True
    RadioButton6.Enabled = True
    
    checkbox12.Value = False
    Rb1.Enabled = False
    Rb2.Enabled = False
    Rb3.Enabled = False
    Rb4.Enabled = False
    FlatEdit2.Enabled = False
    
    RadioButton5.Value = True
    
End Sub

Private Sub CheckBox9_Click()
If CheckBox9.Value = xtpChecked Then
    DownLoadFile = 1
Else
    DownLoadFile = 0
End If
End Sub

Private Sub form_load()

    checkbox12.Value = True

    With Com1
        .AddItem "Seconds"
        .AddItem "Minutes"
        .AddItem "Hours"
        .ListIndex = 0
        .Height = 300
    End With

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image11.Visible = True
Image6.Visible = False
Image5.Visible = True

If Dragging Then
Me.Left = Me.Left + (X - SettedX)
Me.Top = Me.Top + (Y - SettedY)
End If

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = True
Image11.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image5.Visible = False
End Sub

Private Sub Image6_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image9_Click()
Me.Hide
End Sub
