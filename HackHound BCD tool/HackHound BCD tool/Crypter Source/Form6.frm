VERSION 5.00
Begin VB.Form Form6 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   LinkTopic       =   "Form3"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      Caption         =   "Change File Information"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4695
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image7 
         Height          =   435
         Left            =   1920
         Picture         =   "Form6.frx":D5B7
         Top             =   840
         Width           =   1965
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         Caption         =   "EXE to get file info from :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Caption         =   "Icon Source"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4695
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000014&
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   915
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000016&
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   1920
         Picture         =   "Form6.frx":1122B
         Top             =   840
         Width           =   1965
      End
   End
   Begin VB.Image Image5 
      Height          =   8610
      Left            =   0
      Picture         =   "Form6.frx":14E9F
      Top             =   360
      Width           =   60
   End
   Begin VB.Image Image6 
      Height          =   60
      Left            =   -360
      Picture         =   "Form6.frx":17FF5
      Top             =   4200
      Width           =   11985
   End
   Begin VB.Image Image4 
      Height          =   8610
      Left            =   4920
      Picture         =   "Form6.frx":1B20C
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   1440
      Picture         =   "Form6.frx":1E319
      Top             =   3600
      Width           =   1965
   End
   Begin VB.Image Image15 
      DragMode        =   1  'Automatic
      Height          =   12000
      Left            =   0
      Picture         =   "Form6.frx":21F8D
      Top             =   360
      Width           =   12000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "                             Icon Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   4320
      Picture         =   "Form6.frx":2F544
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   4320
      Picture         =   "Form6.frx":32769
      Top             =   60
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   0
      Picture         =   "Form6.frx":35982
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Crypter based off Cobeins Cryptosy
'Edited by legssmit
' Use  : At your own risk
' ' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission FROM COBEIN AND ME (Legssmit).
Private Dragging As Boolean
Private SettedX As Integer, SettedY As Integer
Private Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Sub Form_Load()
Text1 = Form1.Text2
End Sub

Private Sub Image1_Click()
Call BrowseForIcon
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = True
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = False
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.Visible = False
End Sub

Private Sub Image3_Click()
InfoPath = Text2
Unload Form6
End Sub

Private Sub Image7_Click()
Call GetFileInformation
End Sub

Private Sub Image8_Click()
InfoPath = Text2
Unload Form6
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label12.Visible = True
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = True
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.Visible = True
End Sub
Private Sub Label12_Click()
InfoPath = Text2
Unload Form6
End Sub

Private Sub Label6_Click()
Call GetFileInformation
End Sub
Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label12.Visible = False
Label3.Visible = False
Image3.Visible = False
End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SettedX = x
    SettedY = y
    Dragging = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Dragging Then
        Me.Left = Me.Left + (x - SettedX)
        Me.Top = Me.Top + (y - SettedY)
    End If
Image3.Visible = False
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dragging = False
End Sub

Private Sub BrowseForIcon()
Dim sIcon As String
Dim SERVER_RESOURCE() As Byte
            sIcon = Space(255)
            GetFileNameFromBrowseW Me.hWnd, StrPtr(sIcon), 255, StrPtr("c:\"), StrPtr("txt"), StrPtr("Icons (*.ICO)" + Chr$(0) + "*.ICO" + Chr$(0) + "Apps (*.EXE)" + Chr$(0) + "*.EXE" + Chr$(0)), StrPtr("Select File")
            Form1.Text2 = Left$(sIcon, lstrlen(sIcon))
            Text1 = Left$(sIcon, lstrlen(sIcon))

            If Right(Text1, 3) = "exe" Then
            Open App.Path + "\ResHacker.exe" For Binary As #1
            SERVER_RESOURCE() = LoadResData(102, "CUSTOM")
            Put #1, , SERVER_RESOURCE()
            Close #1

            If PathFileExists(App.Path & "\Icon_1.ico") Then
            Kill App.Path & "\Icon_1.ico"
            End If

            Call Shell("ResHacker.exe -extract " & Chr(34) & Text1 & Chr(34) & "," & App.Path & "\myprogicons.rc" & "," & "icongroup,,""")
            Call UpdatePicture(True)
            Else
            Call UpdatePicture(False)
            End If
End Sub

Private Sub Label1_Click()
Call BrowseForIcon
End Sub

Private Sub UpdatePicture(ByVal IfExe As Boolean)
If IfExe = False Then Picture1.Picture = LoadPicture(Text1) ' change to your picture path
If IfExe = True Then Picture1.Picture = LoadPicture(App.Path & "\Icon_1.ico")
Dim x, y As Single
x = Picture1.Width
y = Picture1.Height
Do While x > Picture1.Width Or y > Picture1.Height
x = x / 1.01
y = y / 1.01
Loop
Picture1.PaintPicture Picture1.Picture, 0, 0, x, y
If IfExe = True Then
Form1.Text2 = App.Path & "\Icon_1.ico"
Text1 = App.Path & "\Icon_1.ico"
End If
            If PathFileExists(App.Path & "\ResHacker.log") Then
            Kill App.Path & "\ResHacker.log"
            Kill App.Path & "\ResHacker.ini"
            End If
End Sub

Private Sub GetFileInformation()
Dim SERVER_RESOURCE() As Byte
Dim sSave As String
            sSave = Space(255)
            GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("txt"), StrPtr("Apps (*.EXE)" + Chr$(0) + "*.EXE" + Chr$(0)), StrPtr("Select File")
            Text2 = Left$(sSave, lstrlen(sSave))
            If PathFileExists(App.Path & "\ResHacker.exe") Then
            Kill App.Path & "\ResHacker.exe"
            End If
                
            InfoPath = Text2
                
            Open App.Path + "\ResHacker.exe" For Binary As #1
            SERVER_RESOURCE() = LoadResData(102, "CUSTOM")
            Put #1, , SERVER_RESOURCE()
            Close #1
            Call Shell("ResHacker.exe -extract " & Chr(34) & Text2 & Chr(34) & "," & Chr(34) & App.Path & "\test.res" & Chr(34) & ", VERSIONINFO, 1, 1033")
End Sub

