VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "CODEJO~3.OCX"
Begin VB.Form Message 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Message.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Message.frx":F172
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   615
      Left            =   2520
      TabIndex        =   31
      Top             =   2160
      Width           =   8415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   2430
      Left            =   2520
      TabIndex        =   30
      Top             =   3000
      Width           =   8415
   End
   Begin XtremeSuiteControls.CheckBox CheckBox7 
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   1800
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   120
      Picture         =   "Message.frx":1E2E4
      ScaleHeight     =   3255
      ScaleWidth      =   2415
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
      Begin XtremeSuiteControls.RadioButton CheckBox1 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton Checkbox6 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton CheckBox2 
         Height          =   255
         Left            =   120
         TabIndex        =   20
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
      Begin XtremeSuiteControls.RadioButton CheckBox3 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton Checkbox4 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton Checkbox5 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RadioButton1"
         BackColor       =   -2147483641
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox CheckBox8 
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   195
         _Version        =   851968
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         Caption         =   "Play On Installation"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.Label Label14 
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   0
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Play on installation"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   720
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Displayed Buttons"
         ForeColor       =   16777215
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
      Begin XtremeSuiteControls.Label Label12 
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2280
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Yes, No, Cancle"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Abort, Retry, Ignore"
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
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "OK, Cancle"
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
         Left            =   360
         TabIndex        =   14
         Top             =   2640
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Yes, No"
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
         Left            =   360
         TabIndex        =   13
         Top             =   3000
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Retry, Cancle"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "OK Only"
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
   End
   Begin XtremeSuiteControls.ComboBox ComboBox2 
      Height          =   360
      Left            =   4440
      TabIndex        =   24
      Top             =   1560
      Width           =   4455
      _Version        =   851968
      _ExtentX        =   7858
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
      Text            =   "ComboBox2"
      DropDownItemCount=   6
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   480
      Picture         =   "Message.frx":2E717
      ScaleHeight     =   735
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   840
      Width           =   10935
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   255
         Left            =   7080
         TabIndex        =   6
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
      Begin XtremeSuiteControls.RadioButton RadioButton2 
         Height          =   255
         Left            =   9000
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
      Begin XtremeSuiteControls.RadioButton RadioButton3 
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
      Begin XtremeSuiteControls.RadioButton RadioButton4 
         Height          =   255
         Left            =   3000
         TabIndex        =   9
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
      Begin XtremeSuiteControls.RadioButton RadioButton5 
         Height          =   255
         Left            =   5160
         TabIndex        =   10
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
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Left            =   9240
         TabIndex        =   5
         Top             =   120
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No Icon"
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
         Index           =   7980
         Left            =   7400
         TabIndex        =   4
         Top             =   120
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Warning"
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
         Left            =   5400
         TabIndex        =   3
         Top             =   120
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Question"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Information"
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
         Left            =   840
         TabIndex        =   1
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Critical Error"
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
      Begin VB.Image Image5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image4 
         Enabled         =   0   'False
         Height          =   510
         Left            =   0
         Picture         =   "Message.frx":3EB4A
         Top             =   0
         Width           =   510
      End
      Begin VB.Image Image3 
         Enabled         =   0   'False
         Height          =   510
         Left            =   4560
         Picture         =   "Message.frx":3EF32
         Top             =   0
         Width           =   510
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   510
         Left            =   2400
         Picture         =   "Message.frx":3F36C
         Top             =   0
         Width           =   510
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   510
         Left            =   6480
         Picture         =   "Message.frx":3F751
         Top             =   0
         Width           =   510
      End
   End
   Begin XtremeSuiteControls.Label Label15 
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   3615
      _Version        =   851968
      _ExtentX        =   6376
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Fake Runtime Message"
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
   Begin VB.Image Image10 
      Height          =   285
      Left            =   10845
      Picture         =   "Message.frx":3FBA7
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   10845
      Picture         =   "Message.frx":42FB2
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   10400
      Picture         =   "Message.frx":463AA
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   10400
      Picture         =   "Message.frx":4942D
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   9120
      Picture         =   "Message.frx":4C5A2
      Top             =   1560
      Width           =   1905
   End
   Begin XtremeSuiteControls.Label Label13 
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   1800
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Enable Message"
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
   Begin VB.Image Image11 
      Height          =   390
      Left            =   9120
      Picture         =   "Message.frx":50181
      Top             =   1560
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Background 
      Height          =   6000
      Left            =   0
      Picture         =   "Message.frx":53FA2
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SettedX As Integer, SettedY As Integer, Dragging As Boolean

Private Sub Background_Click()
CheckBox7.SetFocus
End Sub

Private Sub Background_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 SettedX = X
    SettedY = Y
    Dragging = True
End Sub


Private Sub Background_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image11.Visible = False
Image9.Visible = False
Image10.Visible = True
Image7.Visible = False
Image8.Visible = True
 If Dragging Then
        Me.Left = Me.Left + (X - SettedX)
        Me.Top = Me.Top + (Y - SettedY)
    End If

End Sub

Private Sub Background_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub CheckBox7_Click()

If CheckBox8.Value = xtpChecked Then
CheckBox8.Value = xtpChecked
Message_Play = 1
Else
Message_Play = 0
CheckBox8.Value = xtpUnchecked
End If

End Sub

Private Sub form_load()
 
' Populate Msg CB
Message_Play = 0
List1.AddItem "Keep message body null"
List1.Text = List1.Text & vbCrLf
GetErrorList List1
List1.ListIndex = 0
' Populate Title CB
ComboBox2.AddItem (" Error")
ComboBox2.AddItem (" System Error")
ComboBox2.AddItem (" Missing File")
ComboBox2.AddItem (" Critical Error")
ComboBox2.AddItem (" Unknown Failure")
ComboBox2.AddItem (" Memory Error")
ComboBox2.AddItem (" Windows Error")
ComboBox2.AddItem (" Success")
ComboBox2.AddItem (" Microsoft Visual Basic ")
ComboBox2.AddItem (" Norton Internet Security")
ComboBox2.AddItem (" LimeWire Pro")
ComboBox2.AddItem (" Mozilla Firefox")
ComboBox2.AddItem (" Microsoft")
ComboBox2.AddItem (" Restart Prompt")
ComboBox2.AddItem (" Use Null Message Title")

Message_Body = vbNullString
Message_Title = vbNullString
Message_Icon = 0
Message_Options = 0

ComboBox2.ListIndex = 0

RadioButton3.Value = True
CheckBox1.Value = True

End Sub
Function GetErrorList(List As ListBox)
    Dim ErCode As Long
    Dim ErDesc As String
    Dim FatalEr As String
    FatalEr = error(1)
            For ErCode = 3 To 750
            If error(ErCode) <> FatalEr Then
                List.AddItem error(ErCode)
             End If
         Next
End Function
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then Message_Options = 0
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then Message_Options = 5
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then Message_Options = 4
End Sub

Private Sub CheckBox4_Click()
If Checkbox4.Value = True Then Message_Options = 3
End Sub

Private Sub CheckBox5_Click()
If Checkbox5.Value = True Then Message_Options = 2
End Sub

Private Sub CheckBox6_Click()
If Checkbox6.Value = True Then Message_Options = 1
End Sub

Private Sub Image10_Click()
Me.Hide
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = True
Image10.Visible = False
End Sub

Private Sub Image11_Click()
Message_Body = Text1.Text
Message_Title = ComboBox2.Text
If ComboBox2.Text = " Use Null Message Title" Then Message_Title = vbNullString
If List1.Text = "Keep message body null" Then Message_Body = vbNullString
If CheckBox7.Value = xtpChecked Then

MsgBox Message_Body, Message_Options + Message_Icon, Message_Title
Else

MsgBox "The message must first be enabled", vbCritical, "Error"
End If

End Sub

Private Sub Image7_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image8.Visible = False
End Sub

Private Sub Image9_Click()
Me.Hide
End Sub

Private Sub Label1_Click()
RadioButton3.Value = True
End Sub

Private Sub Label10_Click()
Checkbox6.Value = True
End Sub

Private Sub Label11_Click()
Checkbox5.Value = True
End Sub

Private Sub Label12_Click()
Checkbox4.Value = True
End Sub

Private Sub Label13_Click()
CheckBox7.Value = xtpChecked
End Sub

Private Sub Label14_Click()
CheckBox8.Value = xtpChecked
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
Image6.Visible = False
End Sub

Private Sub Label2_Click()
CheckBox1.Value = True
End Sub

Private Sub Label3_Click()
RadioButton4.Value = True
End Sub

Private Sub Label4_Click()
RadioButton5.Value = True
End Sub

Private Sub Label5_Click(Index As Integer)
RadioButton1.Value = True
End Sub

Private Sub Label6_Click()
RadioButton2.Value = True
End Sub

Private Sub Label8_Click()
CheckBox2.Value = True
End Sub

Private Sub Label9_Click()
CheckBox3.Value = True
End Sub
Public Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = ""
Text1.Text = List1.Text
End Sub

Private Sub RadioButton1_Click()
If RadioButton1.Value = True Then Message_Icon = 48
End Sub

Private Sub RadioButton2_Click()
If RadioButton2.Value = True Then Message_Icon = 0
End Sub

Private Sub RadioButton3_Click()
If RadioButton3.Value = True Then Message_Icon = 16
End Sub

Private Sub RadioButton4_Click()
If RadioButton4.Value = True Then Message_Icon = 64
End Sub

Private Sub RadioButton5_Click()
If RadioButton5.Value = True Then Message_Icon = 32
End Sub


