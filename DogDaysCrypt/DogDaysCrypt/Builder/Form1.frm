VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "CODEJO~3.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form2"
   ScaleHeight     =   3315
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
   Begin XtremeSuiteControls.FlatEdit txtfile 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
      _Version        =   851972
      _ExtentX        =   8705
      _ExtentY        =   450
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Path..."
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdfilebrowse 
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   1320
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Add"
      ForeColor       =   0
      Appearance      =   6
   End
   Begin XtremeSuiteControls.FlatEdit txtRandom 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
      _Version        =   851972
      _ExtentX        =   8705
      _ExtentY        =   450
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdrnd 
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   1680
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Gen"
      BackColor       =   -2147483633
      Appearance      =   6
   End
   Begin XtremeSuiteControls.CheckBox packfile 
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   2400
      Width           =   975
      _Version        =   851972
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Pack File"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox eofsupport 
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
      _Version        =   851972
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Eof Support"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox checkicon 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Enable Icon Changer"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txticon 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   4935
      _Version        =   851972
      _ExtentX        =   8705
      _ExtentY        =   450
      _StockProps     =   77
      BackColor       =   -2147483643
      Enabled         =   0   'False
      Text            =   "Icon Path..."
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdiconbrowse 
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   2040
      Width           =   615
      _Version        =   851972
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Add"
      BackColor       =   -2147483633
      Enabled         =   0   'False
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdabout 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
      _Version        =   851972
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "About"
      BackColor       =   -2147483633
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdcrypt 
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   2880
      Width           =   2895
      _Version        =   851972
      _ExtentX        =   5106
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Crypt"
      BackColor       =   -2147483633
      Appearance      =   6
   End
   Begin XtremeSuiteControls.CommonDialog wow 
      Left            =   120
      Top             =   4560
      _Version        =   851972
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DialogStyle     =   1
   End
   Begin XtremeSuiteControls.CommonDialog iconwow 
      Left            =   480
      Top             =   4560
      _Version        =   851972
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DialogStyle     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RT_MESSAGETABLE      As Long = &H11

Private Sub checkicon_Click()
If checkicon.Value = xtpChecked Then
txticon.Enabled = True
cmdiconbrowse.Enabled = True
Else
txticon.Enabled = False
cmdiconbrowse.Enabled = False
End If
End Sub

Private Sub cmdabout_Click()
MsgBox "By Limited for Dogdays", vbInformation, "About"
End Sub

Private Sub cmdcrypt_Click()
Dim xd As String
Dim bfile() As Byte
Dim EofLen As String
Dim aFile As String
Dim tFile() As Byte
Dim mKey As String

mKey = "ZGUzkfgZTDFtzdIUtgugZzf"

If FileExists(App.Path + "\" + "doggy.exe") = False Then
MsgBox "No Stub!", vbInformation, "Info"
Exit Sub
End If

With wow
.CancelError = False
.FileName = ""
.DialogTitle = "Save Pe File"
.Filter = "Executable Files (*.exe) | *.exe"
.ShowSave
End With

Open txtfile.Text For Binary As #1
xd = Space$(LOF(1))
Get #1, , xd
Close #1

If eofsupport.Value = 1 Then
EofLen = GetEOFData(txtfile.Text)
End If

Call FileCopy(App.Path & "\doggy.exe", wow.FileName)

bfile() = StrConv(xd, vbFromUnicode)

If packfile.Value = 1 Then
bfile() = Compress(bfile())
End If

Call RC4(bfile(), txtRandom.Text)

Call SetResourceBytes(RT_MESSAGETABLE, 6000, AddJpgHeader(bfile()), wow.FileName)

aFile = mKey & packfile.Value & mKey & txtRandom.Text & mKey

tFile() = StrConv(aFile, vbFromUnicode)

Call SetResourceBytes(RT_MESSAGETABLE, 7000, tFile(), wow.FileName)

If checkicon.Value = xtpChecked Then
Call ChangeIcon(wow.FileName, txticon.Text)
End If

If eofsupport.Value = 1 Then
Call WriteEOFData(wow.FileName, EofLen)
End If

MsgBox "Done !" & vbNewLine & "Final Size: " & FileLen(wow.FileName) & " Bytes", vbInformation, "Info"

End Sub

Private Sub cmdfilebrowse_Click()
With wow
.CancelError = False
.DialogTitle = "Open Pe File"
.Filter = "Executable Files (*.exe) | *.exe"
.ShowOpen
txtfile.Text = .FileName
End With
End Sub

Private Sub cmdiconbrowse_Click()
With iconwow
.CancelError = False
.DialogTitle = "Open Icon"
.Filter = "Icons (*.ico)|*.ico"
.ShowOpen
txticon.Text = .FileName
End With
End Sub

Private Sub cmdrnd_Click()
txtRandom.Text = RndNames
End Sub

Private Sub Form_Load()
cmdrnd.Value = True
End Sub

