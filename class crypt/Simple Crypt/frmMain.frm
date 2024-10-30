VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Simple Crypt by mar"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4785
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdCrypt 
      Caption         =   "Verschl¸sseln!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame grpOpt 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      Begin VB.CheckBox chkReAlign 
         Caption         =   "ReAlign PE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkAddSec 
         Caption         =   "Add Section"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   1335
      End
   End
   Begin VB.TextBox txtFile 
      Alignment       =   2  'Zentriert
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      Text            =   "Drag n Drop"
      Top             =   120
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1080
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CryptKey As String = "AZJ|ƒ…FLt⁄à(kˆ3Ö®{#ÓÈˇæ†øb’ó‹"
Const SplitKey As String = "ﬁ63xàîp€>˝πkuÕi|∏‰_ôb'ú¸˜êAãO`"

Dim sBuffer As String
Dim sBuffer2 As String
Dim sStub As String
Dim crypt As New clsTwofish
Dim sOutput As String
Dim resdata() As Byte

Private Sub cmdCrypt_Click()
On Error GoTo errhandle
If txtFile.Text = "Drag n Drop" Then
    MsgBox "Keine File ausgew‰hlt", vbCritical
    Exit Sub
End If

With CD
    .DialogTitle = "Speicher"
    .Filter = ".exe|*.exe"
    .CancelError = True
    .ShowSave
    sOutput = .Filename
End With

'Stub laden
resdata() = LoadResData(101, "CUSTOM")
Open Environ("TEMP") & "\stub.exe" For Binary As #1
    Put #1, , resdata
Close #1

sStub = Environ("TEMP") & "\stub.exe"
   
Open sStub For Binary As #1
    sBuffer = Space(LOF(1))
    Get #1, , sBuffer
Close #1

Open txtFile.Text For Binary As #1
    sBuffer2 = Space(LOF(1))
    Get #1, , sBuffer2
    sBuffer2 = crypt.EncryptString(sBuffer2, CryptKey, False)
Close #1

Open sOutput For Binary As #1
    Put #1, , sBuffer & SplitKey & sBuffer2 & SplitKey
Close #1



'Optionen
'AddSec
Call AddSection(sOutput, ".reloc", Len(sBuffer2), &H8)

'ReAlign
If chkReAlign.Value = 1 Then
    If RealignPEFromFile(sOutput) = True Then
        MsgBox "PE ReAligned!", vbInformation
    Else
        MsgBox "ReAlign failed!", vbCritical
    End If
End If

Kill sStub
MsgBox "Fertig!", vbInformation

Exit Sub

errhandle:

If Err.Number = 32755 Then
    Exit Sub
End If

End Sub



Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
txtFile.Text = Data.Files(1)
End Sub
