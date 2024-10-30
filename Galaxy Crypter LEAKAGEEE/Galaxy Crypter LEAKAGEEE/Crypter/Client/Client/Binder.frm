VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "CODEJO~3.OCX"
Begin VB.Form Binder 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Binder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView ListView1 
      Height          =   2775
      Left            =   45
      TabIndex        =   0
      Top             =   840
      Width           =   11175
      _Version        =   851968
      _ExtentX        =   19711
      _ExtentY        =   4895
      _StockProps     =   77
      ForeColor       =   16777088
      BackColor       =   -2147483647
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      FullRowSelect   =   -1  'True
      BackColor       =   -2147483647
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Sorted          =   -1  'True
      Arrange         =   2
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   360
      Left            =   360
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
      _Version        =   851968
      _ExtentX        =   4260
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
      Text            =   "Filename"
      BackColor       =   -2147483641
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   5295
      _Version        =   851968
      _ExtentX        =   9340
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
      BackColor       =   -2147483641
      Appearance      =   1
      UseVisualStyle  =   0   'False
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox ComboBox2 
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   4680
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
   Begin XtremeSuiteControls.ComboBox ComboBox3 
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   4200
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
   Begin XtremeSuiteControls.CheckBox CheckBox6 
      Height          =   195
      Left            =   6480
      TabIndex        =   10
      Top             =   4965
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
      Left            =   6480
      TabIndex        =   14
      Top             =   5325
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
      Left            =   6480
      TabIndex        =   16
      Top             =   4590
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ComboBox ComboBox1 
      Height          =   360
      Left            =   2880
      TabIndex        =   18
      Top             =   5160
      Width           =   855
      _Version        =   851968
      _ExtentX        =   1508
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
      Text            =   "Extension"
   End
   Begin XtremeSuiteControls.Label Label9 
      Height          =   255
      Left            =   6720
      TabIndex        =   17
      Top             =   4560
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Encrypt File"
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
      Left            =   6720
      TabIndex        =   15
      Top             =   5280
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inject File Into Memory"
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
      Left            =   6000
      TabIndex        =   13
      Top             =   4080
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
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
      Alignment       =   2
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   3840
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
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
      Alignment       =   2
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Compress File"
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
      TabIndex        =   9
      Top             =   3720
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   450
      _StockProps     =   79
      ForeColor       =   255
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   5280
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "<<< Filename After"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "<<< Extract File To "
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "File Binder"
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
   Begin VB.Image Image12 
      Enabled         =   0   'False
      Height          =   390
      Left            =   9000
      Picture         =   "Binder.frx":F172
      Top             =   4080
      Width           =   1905
   End
   Begin VB.Image Image11 
      Height          =   390
      Left            =   9000
      Picture         =   "Binder.frx":13730
      Top             =   4080
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image10 
      Height          =   285
      Left            =   10845
      Picture         =   "Binder.frx":1812E
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   10845
      Picture         =   "Binder.frx":1B539
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   10395
      Picture         =   "Binder.frx":1E931
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   10400
      Picture         =   "Binder.frx":21AA6
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   10400
      Picture         =   "Binder.frx":24B29
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   9000
      Picture         =   "Binder.frx":27BAC
      Top             =   4560
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   9000
      Picture         =   "Binder.frx":2C46C
      Top             =   5040
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image4 
      Enabled         =   0   'False
      Height          =   390
      Left            =   9000
      Picture         =   "Binder.frx":30633
      Top             =   5040
      Width           =   1905
   End
   Begin XtremeSuiteControls.Label Label14 
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   4320
      Width           =   2055
      _Version        =   851968
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "<<< Execution Mode"
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
   Begin XtremeSuiteControls.CommonDialog CDBind 
      Left            =   3240
      Top             =   120
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   9000
      Picture         =   "Binder.frx":34623
      Top             =   4560
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Binder.frx":38AC0
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Binder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long

Private SettedX As Integer, SettedY As Integer, Dragging As Boolean

Private Sub CheckBox1_Click()

    If CheckBox1.Value = xtpChecked Then
        ComboBox3.ListIndex = 2
    End If
    
End Sub

Private Sub CheckBox2_Click()

    If CheckBox2.Value = xtpChecked Then EncryptBound = "1" Else EncryptBound = "0"
    
    End Sub

Private Sub form_load()

ListView1.ColumnHeaders.Add , , "File Path", 3500
ListView1.ColumnHeaders.Add , , "Extract To", 1800
ListView1.ColumnHeaders.Add , , "Execute", 1800
ListView1.ColumnHeaders.Add , , "Size", 1000
ListView1.ColumnHeaders.Add , , "Filename", 1200
ListView1.ColumnHeaders.Add , , "UPX", 800
ListView1.ColumnHeaders.Add , , "Encrypt", 800
ListView1.ColumnHeaders.Add , , "Inject", 800

' Extract to
ComboBox2.AddItem ("Application Directory ")
ComboBox2.AddItem ("Windows Directory ")
ComboBox2.AddItem ("System Directory ")
ComboBox2.AddItem ("Temp Directory ")
ComboBox2.AddItem ("System 32")
ComboBox2.AddItem ("Program Files")
ComboBox2.AddItem ("AppData (Documents & Settings) ")
ComboBox2.ListIndex = 0

' Run as hidden
ComboBox3.AddItem "Shell Execute As Normal"
ComboBox3.AddItem "Shell Execute As Hidden"
ComboBox3.AddItem "Do Not Execute File"
ComboBox3.ListIndex = 0

With ComboBox1
    .AddItem ".exe"
    .AddItem ".bat"
    .AddItem ".com"
    .AddItem ".bin"
    .AddItem ".bin"
    .AddItem ".txt"
    .AddItem ".ini"
    .AddItem ".rar"
    .AddItem ".jpg"
    .AddItem ".ico"
    .AddItem ".bmp"
    .AddItem ".png"
    .AddItem ".psd"
    .AddItem ".mp3"
    .AddItem ".wav"
    .AddItem ".wmp"
    .AddItem ".avi"
    .AddItem ".mp4"
    .AddItem ".zip"
    .AddItem ".vbp"
    .AddItem ".bas"
    .AddItem ".cls"
    .AddItem ".cur"
    .ListIndex = 0
End With


BoundSize = 0
Label3.Caption = "Total Size Of Bound Files: "
Label5.Caption = FormatKB(BoundSize)


End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image11.Visible = False
Image12.Visible = True
Image2.Visible = True
Image3.Visible = False
Image5.Visible = False
Image4.Visible = True
Image9.Visible = False
Image10.Visible = True
Image8.Visible = False
Image7.Visible = True
 
 If Dragging Then
        Me.Left = Me.Left + (X - SettedX)
        Me.Top = Me.Top + (Y - SettedY)
    End If

End Sub

Private Sub Image10_Click()
Me.Hide
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = False
Image9.Visible = True
Image10.Visible = False
Image7.Visible = True
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Visible = True
Image11.Visible = False
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Visible = False
Image11.Visible = True

Dim i As Integer
Dim z As Integer
Dim J As Integer

If ListView1.SelectedCount = 0 Then ListView1.ListItems.Remove (ListView1.ListItems.Count)


For i = 1 To ListView1.SelectedCount
ListView1.ListItems.Remove (ListView1.SelectedCount)
BoundSize = 0
For z = 1 To ListView1.ListItems.Count
BoundSize = BoundSize + FileLen(ListView1.ListItems(z).Text)
Next z
Label3.Caption = "Total Size Of Bound Files: "
Label5.Caption = FormatKB(BoundSize)
Next i

If ListView1.ListItems.Count = 0 Then
Image11.Enabled = True
Image12.Enabled = False
End If
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
Image12.Visible = False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image2.Visible = False
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image3.Visible = False
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Image3.Visible = True

With CDBind
.DialogTitle = "Select any file to bind"
.Filter = "All Files (*.*)*.* "
.ShowOpen
End With

If Not CDBind.Filename = vbNullString Then
    Text1.Text = CDBind.Filename
    Image4.Enabled = True
Dim Filename() As String
    Filename = Split(CDBind.Filename, "\")
FlatEdit1.Text = Filename(UBound(Filename))
Else
    Text1.Text = "Locate a file to begin..."
    Image4.Enabled = False
    Exit Sub
End If
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = True
Image4.Visible = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image5.Visible = False
End Sub

Public Sub Terminate_Mci(ByVal sAlias As String)
On Local Error Resume Next
mciSendString "Stop " & sAlias, 0, 0, 0
mciSendString "Close " & sAlias, 0, 0, 0
End Sub

Public Sub Initialize_Mci(ByVal FilePath As String, sAlias As String)
mciSendString "Stop " & sAlias, 0, 0, 0
mciSendString "Close " & sAlias, 0, 0, 0
mciSendString "open " & FilePath & " Type MPEGVIDEO alias " & sAlias, 0, 0, 0
mciSendString "Play " & sAlias & " from 0", 0, 0, 0
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image5.Visible = True
On Error Resume Next

Dim i As Integer
Dim mystring() As Byte
mystring = LoadResData("Critical", "Custom")
If Build.Fileexists(Environ("Temp") & "\TmpSnd.wav") Then Kill Environ("Temp") & "\tmpsnd.wav"

Open Environ$("Temp") & "\TmpSnd.wav" For Binary As #1
Put #1, , mystring
Close #1

    For i = 1 To ListView1.ListItems.Count
        If InStr(Binder.ListView1.ListItems(i).SubItems(4), FlatEdit1.Text) <> 0 Then

            Call Initialize_Mci(Environ$("Temp") & "\tmpsnd.wav", "myalias")
            Label2.Caption = "Error: Filename already exists!"
            Text1.Visible = False
            Delay 1
            
            Call Terminate_Mci("myalias")
            Label2.Caption = ""
            Text1.Visible = True
            Exit Sub
        End If
    Next i

If Text1.Text = "Locate a file to begin..." Or Text1.Text = "" Then Exit Sub

Dim FinalStr As String
Dim Ext_Split() As String

Ext_Split = Split(FlatEdit1.Text, ".")

If Ext_Split(1) <> "" Then FinalStr = Replace(FlatEdit1, Ext_Split(1), "")
 FinalStr = Replace(FinalStr, ".", "")
 FlatEdit1.Text = FinalStr

ListView1.ListItems.Add ListView1.ListItems.Count + 1, , CDBind.Filename
ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(1) = ComboBox2.Text
ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(2) = ComboBox3.Text
ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(3) = FormatKB(FileLen(Text1.Text))
ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(4) = FlatEdit1.Text & ComboBox1.Text
If CheckBox6.Value = xtpChecked Then ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(5) = "Yes" Else: ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(5) = "No"
If CheckBox2.Value = xtpChecked Then ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(6) = "Yes" Else: ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(6) = "No"
If CheckBox1.Value = xtpChecked Then ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(7) = "Yes" Else: ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(7) = "No"

BoundSize = 0
For i = 1 To ListView1.ListItems.Count
BoundSize = BoundSize + FileLen(ListView1.ListItems(i).Text)
Next i

Label3.Caption = "Total Size Of Bound Files: "
Label5.Caption = FormatKB(BoundSize)
If ListView1.ListItems.Count > 0 Then
Image12.Enabled = True
End If

FlatEdit1.Text = FlatEdit1.Text & "." & Ext_Split(1)

End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image8.Visible = True
Image7.Visible = False
Image10.Visible = True
End Sub

Private Sub Image8_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image9_Click()
Me.Hide
End Sub



Private Sub Label6_Click()
If CheckBox6.Value = xtpChecked Then CheckBox6.Value = xtpUnchecked: Exit Sub
If CheckBox6.Value = xtpUnchecked Then CheckBox6.Value = xtpChecked: Exit Sub
End Sub

Private Sub Text1_Click()
Image3_MouseUp 1, 1, 1, 1
End Sub
