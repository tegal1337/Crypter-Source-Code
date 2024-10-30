VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Roxxer's PowerBasic Crypter Example"
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCrypt 
      Caption         =   "Crypt File"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4680
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFileIn 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
MsgBox "PowerBasic Crypter Example" & vbNewLine _
& "by Roxxer." & vbNewLine & vbNewLine _
& "Thanks goes to:" & vbNewLine _
& "steve10120" & vbNewLine _
& "Slayer616" & vbNewLine _
& "cobein" & vbNewLine & vbNewLine _
& "Visit: www.HackHound.org", vbInformation, "About"
End Sub

Private Sub cmdCrypt_Click()
Dim sFileData As String
Dim sStubData As String

If txtFileIn.Text = "" Then
    MsgBox "Please select a File."
    Exit Sub
End If

Open App.Path & "\stub.exe" For Binary As #1
    sStubData = Space(LOF(1))
    Get #1, , sStubData
Close #1

Open txtFileIn.Text For Binary As #1
    sFileData = Space(LOF(1))
    Get #1, , sFileData
Close #1

sFileData = sStubData & "SplitItHere" & XOREncryption("lol", sFileData) & "SplitItHere"

Open App.Path & "\Crypted.exe" For Binary As #1
    Put #1, , sFileData
Close #1

MsgBox "File crypted sucessfuly."
End Sub

Private Sub txtFileIn_Click()
CD1.Filter = "Exe|*.exe"
CD1.ShowOpen

If CD1.FileName <> "" Then
    txtFileIn.Text = CD1.FileName
End If
End Sub
