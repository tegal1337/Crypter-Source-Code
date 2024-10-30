VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuild 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "/slash v 0.9"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin Builder.ccXPButton ccXPButton1 
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   2640
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "A&bout"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H80000005&
      Caption         =   "Use RC4"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H80000005&
      Caption         =   "Melt Stub"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":06D7
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CheckBox chk 
      BackColor       =   &H80000005&
      Caption         =   "Hidden"
      DisabledPicture =   "Form1.frx":0AA4
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":19802
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox txtstatus 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   8535
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   0
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   5880
   End
   Begin MSComctlLib.ListView lstv 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   720
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Builder.ccXPButton cmdclear 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "&Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Builder.ccXPButton cmdremove 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1400
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "&Lˆschen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Builder.ccXPButton cmdadd 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "&Hinzuf¸gen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Builder.ccXPButton cmdedit 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   2640
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "&Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Builder.ccXPButton cmdaaaa 
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   5
      Top             =   2640
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "&Build"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   -720
      Picture         =   "Form1.frx":19BCF
      Top             =   -2040
      Width           =   19200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -120
      TabIndex        =   6
      Top             =   6720
      Width           =   10335
   End
End
Attribute VB_Name = "frmBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ccXPButton1_Click()
frmAbout.Show
End Sub

Private Sub chk_Click(Index As Integer)

HiddenStub = chk(0).Value
MeltStub = chk(1).Value
UseRC4 = chk(2).Value

End Sub

Private Sub cmdabout_Click(Index As Integer)

End Sub

Private Sub cmdaaaa_Click(Index As Integer)
Dim Output As String
    Dim SEF As String
    SEF = "£⁄±}G±}G√÷M/·≈±}GÀ"
    Dim SPF As String
    SPF = "~ˇIÒu∞S·˚Ì8∞DO”Æoy"

    Dim bStub() As Byte
    Dim intStub As Integer
    Dim exePath As String
    Dim stubPath As String
    Dim StubSettings As String
    
    If lstv.ListItems.Count <= 0 Then Exit Sub
    
 
CHOFILE:
    With CommonDialog1
        
        .InitDir = App.Path
        .Filter = ".exe|*.exe"
        .DialogTitle = "Save As"
        .ShowSave
        exePath = .FileName
    End With
    
    If exePath = "" Then Exit Sub


    
    With CommonDialog2
        .InitDir = App.Path & "\Stub"
        .Filter = "*.x30|*.x30"
        .DialogTitle = "Chose Stub"
        .ShowOpen
        stubPath = .FileName
    End With
    DoEvents
    If stubPath = "" Then Exit Sub

        Dim F As Integer
        F = FreeFile

    Open stubPath For Binary As #F
        bStub() = Space(LOF(F))
        Get #F, , bStub
    Close #F






    Dim e As Integer
    Dim FFile As String
    Dim AFFile As String
    
    Dim i As Integer
    e = FreeFile
    
    Dim STC As String
    STC = Time
    Me.Enabled = False
  For xy = 1 To lstv.ListItems.Count

    Open lstv.ListItems.Item(xy).SubItems(2) For Binary As #e
        FFile = Space(LOF(F))
        Get #e, , FFile
    Close #e
    FFile = lstv.ListItems.Item(xy).SubItems(3) & SEF & lstv.ListItems.Item(xy).SubItems(4) & SEF & FFile & SEF & lstv.ListItems.Item(xy).SubItems(5)
    If UseRC4 = 0 Then
    Output = Output & FFile & SPF
    ElseIf UseRC4 = 1 Then
    Output = Output & RC4(FFile, "S·”·≈±") & SPF
    End If
    
    Next xy

    
    
    StubSettings = HiddenStub & SEF & MeltStub & SEF & UseRC4
    
    StubSettings = RC4(StubSettings, "G±}G√÷M/·≈±}")
    
    Output = "ﬂ§A:m c•¯/◊˜" & StubSettings & SPF & Output & "ﬂ§A:m c•¯/◊˜"
    
    
    Open exePath For Binary As #1
        Put #1, 1, bStub
        Put #1, , Output
    Close #1
    
    Me.Enabled = True
    MsgBox "Fertig :)"


End Sub

Private Sub cmdadd_Click()
frmadd.Show
End Sub








Private Sub cmdchangeDad_Click()
frmdad.Show
End Sub

Private Sub cmdclear_Click()
    If lstv.ListItems.Count <= 0 Then Exit Sub
    lstv.ListItems.Clear
    Call FILESSIZE
End Sub

Private Sub cmdedit_Click()
    If lstv.ListItems.Count <= 0 Then Exit Sub
    frmedit.Show
End Sub

Private Sub cmdload_Click()
    Dim dfiles As String
    With CommonDialog3
        .InitDir = App.Path
        .Filter = ".dat|*.dat"
        .DialogTitle = "Open FileList"
        .ShowOpen
        dfiles = .FileName
    End With
     If dfiles = "" Then Exit Sub
    lvw_ReadData lstv, dfiles
    Call FILESSIZE
End Sub

Private Sub cmdremove_Click()
    If lstv.ListItems.Count <= 0 Then Exit Sub
    lstv.ListItems.Remove (lstv.SelectedItem.Index)
    Call FILESSIZE
End Sub

Private Sub cmdsave_Click()
    Dim dfiles As String
    With CommonDialog3
        .InitDir = App.Path
        .Filter = ".dat|*.dat"
        .DialogTitle = "Save FileList"
        .ShowSave
        dfiles = .FileName
    End With
    If dfiles = "" Then Exit Sub
    lvw_WriteData lstv, dfiles
End Sub

Private Sub Command1_Click()
    MsgBox lstv.SelectedItem
End Sub

Private Sub cmdstubopt_Click()
    frmstubopt.Show
End Sub

Private Sub Form_Load()
    lstv.ColumnHeaders.Add , , "File-Name", 1700
    lstv.ColumnHeaders.Add , , "Size", 1200
    lstv.ColumnHeaders.Add , , "Path", 1700
    lstv.ColumnHeaders.Add , , "Path to unpack", 1750
    lstv.ColumnHeaders.Add , , "Name", 1500
    lstv.ColumnHeaders.Add , , "Start", 750
    

    txtstatus.Text = "File-Name:  " & vbTab & vbNewLine & _
                 "File-Size:  " & vbTab & vbNewLine & _
                 "File-Path:  " & vbTab & vbNewLine & _
                 "Unpack-Path:" & vbTab & vbNewLine & _
                 "File-Name:  " & vbTab & vbNewLine & _
                 "Start??:    " & vbTab
                 
                 
Call Mache_Transparent(Me.hWnd, 190)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
End Sub



Private Sub lbl_Click(Index As Integer)

End Sub

Private Sub lstv_Click()
    If lstv.ListItems.Count = 1 Then Call status
    If lstv.ListItems.Count > 1 Then Call status
End Sub

Private Sub lstv_DblClick()
    If lstv.ListItems.Count = 0 Then Exit Sub
    frmedit.Show
End Sub

Private Sub lstvinfo_DblClick()
    If lstv.ListItems.Count = 0 Then Exit Sub
    frmedit.Show
End Sub

Private Sub Timer1_Timer()


    
      Label1.Caption = "Files: " & lstv.ListItems.Count & "        " & "Date: " & Format(Now(), "dd:mm:yyyy") & "        " & "Time: " & Format(Now(), "hh:mm:ss") & "        " & RC4("$¶W.m¨çóªﬁp", "yeçsﬂz-`Sı")

                    

End Sub

Private Sub status()


txtstatus.Text = "File-Name:  " & vbTab & lstv.SelectedItem & vbNewLine & _
                 "File-Size:  " & vbTab & lstv.SelectedItem.SubItems(1) & vbNewLine & _
                 "File-Path:  " & vbTab & lstv.SelectedItem.SubItems(2) & vbNewLine & _
                 "Unpack-Path:" & vbTab & lstv.SelectedItem.SubItems(3) & vbNewLine & _
                 "File-Name:  " & vbTab & lstv.SelectedItem.SubItems(4) & vbNewLine & _
                 "Start??:    " & vbTab & lstv.SelectedItem.SubItems(5)
    
End Sub

