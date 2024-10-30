VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit"
   ClientHeight    =   5265
   ClientLeft      =   1215
   ClientTop       =   6690
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtadd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtadd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pfad"
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   4815
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2520
            Width           =   4575
            Begin VB.OptionButton opt 
               Caption         =   "visible"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   22
               Top             =   0
               Width           =   855
            End
            Begin VB.OptionButton opt 
               Caption         =   "no (only unpack)"
               Height          =   255
               Index           =   0
               Left            =   840
               TabIndex        =   21
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton opt 
               Caption         =   "hidden"
               Height          =   255
               Index           =   2
               Left            =   3600
               TabIndex        =   20
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "exe-cute:"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.TextBox txtadd 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   2040
            Width           =   2055
         End
         Begin VB.TextBox txtadd 
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   10
            Top             =   2040
            Width           =   2415
         End
         Begin VB.OptionButton optpath 
            Caption         =   "windir"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton optpath 
            Caption         =   "SystemDrive"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton optpath 
            Caption         =   "ProgramFiles"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
         End
         Begin VB.OptionButton optpath 
            Caption         =   "Appdata"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   1680
            Width           =   1335
         End
         Begin VB.OptionButton optpath 
            Caption         =   "UserProfile"
            Height          =   255
            Index           =   4
            Left            =   1800
            TabIndex        =   5
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optpath 
            Caption         =   "AppPath (Path from Stub)"
            Height          =   255
            Index           =   5
            Left            =   1800
            TabIndex        =   4
            Top             =   960
            Width           =   2415
         End
         Begin VB.OptionButton optpath 
            Caption         =   "TempDir"
            Height          =   255
            Index           =   6
            Left            =   1800
            TabIndex        =   3
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optpath 
            Caption         =   "Custom"
            Height          =   255
            Index           =   7
            Left            =   1800
            TabIndex        =   2
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File -Name && Path to unpack"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2025
         End
      End
      Begin Builder.ccXPButton cmdchose 
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Caption         =   "........"
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
      Begin MSComDlg.CommonDialog cdl 
         Left            =   4560
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   285
      End
   End
   Begin Builder.ccXPButton cmdedit 
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   4680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Caption         =   "&Edit (ALT + E)"
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
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSize As String
Private Sub Form_Load1()
Call Mache_Transparent(Me.hwnd, 190) ' Macht das fenster Transperent, 190
'ein guter wert
End Sub
Private Sub cmdchose_Click()
With cdl
    .InitDir = App.Path
    .DialogTitle = "Add a File"
    .ShowOpen
    txtadd(0).Text = .FileName
    txtadd(1).Text = .FileTitle
    FSize = FileLen(.FileName)
    txtadd(3).Text = .FileTitle ' "\" & .FileTitle
End With
Me.Caption = "Add File: " & txtadd(1).Text
FSize = FSize / 1024
FSize = Format(FSize, "00.00")
End Sub

Private Sub cmdedit_Click()
   
    Dim ExeCute As String

    If opt(0).Value = True Then
        ExeCute = "no"
    ElseIf opt(1).Value = True Then
        ExeCute = "visible"
    ElseIf opt(2).Value = True Then
        ExeCute = "hidden"
    End If


    If txtadd(0).Text = "" Then Exit Sub
    If txtadd(1).Text = "" Then Exit Sub
    If txtadd(2).Text = "" Then txtadd(2).Text = "C:"
    If txtadd(3).Text = "" Then Exit Sub
    
    
    If Right(txtadd(2).Text, 1) = "\" Then txtadd(2) = Left(txtadd(2).Text, Len(txtadd(2).Text) - 1)



    With frmBuild.lstv
        .SelectedItem.SubItems(1) = FSize
        .SelectedItem.SubItems(2) = txtadd(0).Text
        .SelectedItem = txtadd(1).Text
        .SelectedItem.SubItems(3) = txtadd(2).Text
        .SelectedItem.SubItems(4) = txtadd(3).Text
        .SelectedItem.SubItems(5) = ExeCute
    End With
    
    
    Call FILESSIZE
    Unload Me

    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
    With frmBuild.lstv
        FSize = .SelectedItem.SubItems(1)
        txtadd(0).Text = .SelectedItem.SubItems(2)
        txtadd(1).Text = .SelectedItem
        txtadd(2).Text = .SelectedItem.SubItems(3)
        txtadd(3).Text = .SelectedItem.SubItems(4)
    
        If .SelectedItem.SubItems(5) = "no" Then
            opt(0).Value = True
        ElseIf .SelectedItem.SubItems(5) = "visible" Then
            opt(1).Value = True
        ElseIf .SelectedItem.SubItems(4) = "hidden" Then
            opt(2).Value = True
        End If
    End With
    
    Dim pathsinstall As String
    pathinstall = frmBuild.lstv.SelectedItem.SubItems(3)
    
    If pathinstall = "%windir%" Then
        optpath(0).Value = True
    ElseIf pathinstall = "%systemdrive%" Then
        optpath(1).Value = True
    ElseIf pathinstall = "%programfiles%" Then
        optpath(2).Value = True
    ElseIf pathinstall = "%appdata%" Then
        optpath(3).Value = True
    ElseIf pathinstall = "%userprofile%" Then
        optpath(4).Value = True
    ElseIf pathinstall = "%apppath%" Then
        optpath(5).Value = True
    ElseIf pathinstall = "%tempdir%" Then
        optpath(6).Value = True
    Else
        optpath(7).Value = True
        txtadd(2).Text = frmBuild.lstv.SelectedItem.SubItems(3)
    End If
        
    
    Me.Caption = "Edit File: " & frmBuild.lstv.SelectedItem
    frmBuild.Enabled = False
    Me.Show
    txtadd(0).SetFocus
    
    
    Call Mache_Transparent(Me.hwnd, 190)
End Sub




Private Sub Form_Unload(Cancel As Integer)
    frmBuild.Enabled = True
End Sub




Private Sub optpath_Click(Index As Integer)
    If optpath(0).Value = True Then
        txtadd(2).Text = "%windir%"

    ElseIf optpath(1).Value = True Then
        txtadd(2).Text = "%systemdrive%"

    ElseIf optpath(2).Value = True Then
        txtadd(2).Text = "%programfiles%"

    ElseIf optpath(3).Value = True Then
        txtadd(2).Text = "%appdata%"

    ElseIf optpath(4).Value = True Then
        txtadd(2).Text = "%userprofile%"

    ElseIf optpath(5).Value = True Then
        txtadd(2).Text = "%apppath%"

    ElseIf optpath(6).Value = True Then
        txtadd(2).Text = "%tempdir%"

    End If

    If optpath(7).Value = True Then

        txtadd(2).Text = "C:"

    End If
End Sub

