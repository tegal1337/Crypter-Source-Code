VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Lilith"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   2895
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Lilith.chameleonButton chameleonButton3 
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Crypt"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Lilith.chameleonButton chameleonButton2 
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Random"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Lilith.chameleonButton chameleonButton1 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Select"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub chameleonButton1_Click()
On Error GoTo Ende
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Select the Exe you wish to crypt..."
        .DefaultExt = ".exe"
        .Filter = "Executables|*.exe"
    End With
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
    Text3.Text = Text3.Text & "Exe changed: " & Text1.Text & vbCrLf
Ende:
End Sub

Private Sub chameleonButton2_Click()
    Text2.Text = ""
    For i = 1 To 6
        If i = 2 Or i = 4 Or i = 6 Then
            Text2.Text = Text2.Text & RandomNumber
        Else
            Text2.Text = Text2.Text & RandomLetter
        End If
    Next i
    Text3.Text = Text3.Text & "New Key generated: " & Text2.Text & vbCrLf
End Sub

Private Sub chameleonButton3_Click()
    Dim Buffer() As Byte
    Dim ResBuffer() As Byte
    Dim EofData As String
    Dim Buffer2 As String
    Dim Stubpath As String
    
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Select where to save the crypted file.."
        .DefaultExt = ".exe"
        .Filter = "Executables|*.exe"
        .FileName = "Crypted.exe"
    End With
    CommonDialog1.ShowSave
    Stubpath = CommonDialog1.FileName
    
    ResBuffer() = LoadResData(101, "STUB")
    Open Stubpath For Binary As #2
    Put #2, , ResBuffer()
    Close #2

    Text3.Text = Text3.Text & "File read.." & vbCrLf
    Text3.Text = Text3.Text & "Crypting.." & vbCrLf
    
    EncodeArrayB LoadFile(Text1.Text), Text2.Text
    Buffer() = encoded()
    
    Buffer2 = StrConv(LoadFile(Text1.Text), vbUnicode)
    EofData = Mid(Buffer2, GetEOF(Text1.Text), FileLen(Text1.Text))
    
    Open Stubpath For Binary As #1
    Put #1, LOF(1) + 1, "<F1l3>"
    Put #1, LOF(1) + 1, Buffer()
    Put #1, LOF(1) + 1, "<F1l3>"
    Put #1, LOF(1) + 1, Text2.Text
    Put #1, LOF(1) + 1, "<F1l3>"
    Put #1, LOF(1) + 1, EofData
    Close #1
    
    'PatchEOF Stubpath 'removed cause it crashes the eof data
    
    Open Stubpath For Binary As #1
    Put #1, LOF(1) + 1, EofData
    Close #1
    
    Text3.Text = Text3.Text & "Successfull!" & vbCrLf
    MsgBox "The file has been successfully crypted", 64, "Lilith"
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

Private Sub Form_Load()
    Randomize
    Me.Show
    Text3.Text = Text3.Text & "Lilith loaded" & Text2.Text & vbCrLf
    Text2.Text = ""
    For i = 1 To 6
        If i = 2 Or i = 4 Or i = 6 Then
            Text2.Text = Text2.Text & RandomNumber
        Else
            Text2.Text = Text2.Text & RandomLetter
        End If
    Next i
    Text3.Text = Text3.Text & "New Key generated: " & Text2.Text & vbCrLf
End Sub

Private Sub Text1_Change()
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right(Data.Files.Item(1), 4) = ".exe" Then
        Text1.Text = Data.Files.Item(1)
        Text3.Text = Text3.Text & "Exe changed: " & Text1.Text & vbCrLf
    Else
        MsgBox "This is not an Executable", 16, "Lilith"
    End If
End Sub

Private Sub Text3_Change()
    Text3.SelStart = Len(Text3.Text)
End Sub

Private Function LoadFile(ByVal sName As String) As Byte()
   Dim nFile As Integer
   Dim arrFile() As Byte
   nFile = FreeFile
   Open sName For Binary As #nFile
      ReDim arrFile(LOF(nFile) - 1)
      Get #nFile, , arrFile
   Close #nFile
   LoadFile = arrFile
End Function
