VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.ocx"
Begin VB.Form Login 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Galaxy Crypter [Login]"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":F172
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.FlatEdit text3 
      Height          =   3495
      Left            =   5040
      TabIndex        =   11
      Top             =   1460
      Width           =   5655
      _Version        =   851968
      _ExtentX        =   9975
      _ExtentY        =   6165
      _StockProps     =   77
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483641
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      Appearance      =   1
      UseVisualStyle  =   0   'False
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox CheckBox3 
      Height          =   195
      Left            =   5400
      TabIndex        =   7
      Top             =   5190
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox3"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox4 
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   4440
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      BackColor       =   -2147483633
      Enabled         =   0   'False
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   4080
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      BackColor       =   -2147483633
      Enabled         =   0   'False
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox CheckBox2 
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox7"
      BackColor       =   -2147483633
      Enabled         =   0   'False
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BackColor       =   -2147483641
      Alignment       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit2 
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
      _Version        =   851968
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BackColor       =   -2147483641
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   1
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3975
      _Version        =   851968
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Welcome To Galaxy Crypt"
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic Fixed"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
      _Version        =   851968
      _ExtentX        =   7223
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Welcome to Galaxy Crypt"
      ForeColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image Image10 
      Height          =   285
      Left            =   10845
      Picture         =   "Login.frx":13B67
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   10845
      Picture         =   "Login.frx":16F72
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   10400
      Picture         =   "Login.frx":1A36A
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   10400
      Picture         =   "Login.frx":1D4DF
      Top             =   20
      Width           =   420
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   5160
      Width           =   5175
      _Version        =   851968
      _ExtentX        =   9128
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "I have read and fully understand the terms and conditions stated above"
      ForeColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   390
      Left            =   1440
      Picture         =   "Login.frx":20562
      Top             =   3000
      Width           =   1905
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   1440
      Picture         =   "Login.frx":246C0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1905
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Remember Me"
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
      _Version        =   851968
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Remember My Password"
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label8 
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
      _Version        =   851968
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Sign Me In Automatically"
      ForeColor       =   16777088
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Login.frx":28B20
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Const SND_SYNC = &H0        ' Play synchronously (default).
Const SND_NODEFAULT = &H2   ' Do not use default sound.
Const SND_MEMORY = &H4      ' lpszSoundName points to a
Const SND_ASYNC = &H1


 Private SettedX As Integer, SettedY As Integer, Dragging As Boolean
 Dim sAll As String
 Dim m_Access As Boolean
 Dim EndApp As String
 Dim resData() As Byte
 Dim uLower As String
 Dim uUpper As String
 Dim Md5 As New MD5Login
 ' memory file.
Private Sub PlayWaveRes(vntResourceID As String, Optional vntFlags)
Dim bytSound() As Byte
bytSound = LoadResData(vntResourceID, "Custom")
If IsMissing(vntFlags) Then
    vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
End If
If (vntFlags And SND_MEMORY) = 0 Then
    vntFlags = vntFlags Or SND_MEMORY
End If
sndPlaySound bytSound(0), vntFlags
End Sub

Private Sub CheckBox3_Click()

Dim Rmu As String
Dim RMP As String
Dim ReadAs As String
Dim txtuser As String
Dim txtpass As String

WriteIniValue App.Path & "\settings.ini", "Default", "EULA", "Agreed"

If CheckBox3.Value = xtpChecked Then
Label4.Visible = True
CheckBox1.Enabled = True
CheckBox2.Enabled = True
CheckBox4.Enabled = True
Image2.Enabled = True
FlatEdit1.Enabled = True
FlatEdit2.Enabled = True

Rmu = ReadIniValue(App.Path & "\settings.ini", "Default", "RememberMe")
RMP = ReadIniValue(App.Path & "\settings.ini", "Default", "SavePass")
ReadAs = ReadIniValue(App.Path & "\settings.ini", "Default", "Auto-Signin")

If Rmu = "1" Then
CheckBox2.Value = ReadIniValue(App.Path & "\settings.ini", "Default", "RememberMe")
txtuser = ReadIniValue(App.Path & "\settings.ini", "Default", "User")
FlatEdit1.Text = txtuser
End If

If RMP = "1" Then
CheckBox1.Value = ReadIniValue(App.Path & "\settings.ini", "Default", "SavePass")
txtpass = ReadIniValue(App.Path & "\settings.ini", "Default", "Pass")
Debug.Print txtpass
FlatEdit2.Text = txtpass
End If

If ReadAs = "1" Then
CheckBox4.Value = xtpChecked
End If

Else
WriteIniValue App.Path & "\settings.ini", "Default", "EULA", "Declined"
Label4.Visible = False
CheckBox1.Value = xtpUnchecked
CheckBox2.Value = xtpUnchecked
CheckBox4.Value = xtpUnchecked
CheckBox1.Enabled = False
CheckBox2.Enabled = False
CheckBox4.Enabled = False
Image2.Enabled = False
FlatEdit1.Enabled = False
FlatEdit2.Enabled = False
FlatEdit1.Text = ""
FlatEdit2.Text = ""
End If

End Sub

Private Sub CheckBox4_Click()

    If CheckBox4.Value = xtpChecked Then
    
        CheckBox1.Value = xtpChecked And CheckBox2.Value = xtpChecked
        WriteIniValue App.Path & "\settings.ini", "Default", "Auto-SignIn", CheckBox4.Value
    Else
        WriteIniValue App.Path & "\settings.ini", "Default", "Auto-Signin", CheckBox4.Value
    End If

End Sub

Private Sub form_load()

On Error Resume Next
Set ShowTsk = New ShwTskBar
Set ShowTsk.Client = Login
ShowTsk.ShowInTaskbar = True


Dim sPathUser       As String
Dim strCodeKey      As String
Dim enchwid         As String
Dim salt1           As String
Dim XorString       As String
Dim WebSite         As String
Dim StrTemp         As String
Dim lret            As Long

    ' Check if HWID matches:
DoEvents
   lret = URLDownloadToFile(0, "http://host3266.net/coderscentral/HWID1.txt", Environ("Temp") & "\Hwid1.txt", 0, 0)

    hwid1 = CREATEID()
    hwid1 = dbvgbwdiz(hwid1)
    hwid1 = Md5.DigestStrToHexStr(hwid1)

    Open Environ("temp") & "\HWID1.txt" For Input As #1

        Do
            Line Input #1, StrTemp
             If InStr(StrTemp, hwid1) Then
                Close #1
                m_Access = True
                GoTo ResumeLogin
             End If

        Loop Until EOF(1)

        If m_Access = True Then
        Else

        MsgBox "This computer is not authorized to run Galaxy Crypter!", vbCritical + vbOKOnly, "Access Denied!"
        End

    Close #1

        End If

' -----------------------------------------------------------------------------------------------------------------
' -------------------------------------------HWID CHECK IS A SUCCESS----------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------------
' -------------------------------------------GRANT ACCESS TO GALAXY --> --------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------------


ResumeLogin:

' Declarations
Dim Rmu As String
Dim RMP As String
Dim Reula As String

FlatEdit1.Text = vbNullString
FlatEdit2.Text = vbNullString
Label4.Caption = "Welcome to Galaxy Crypt"
Label4.ForeColor = &HFFFF80
'EndApp = 0
'Check for ini
MakeSureDirectoryPathExists App.Path & "\settings.ini"
Reula = ReadIniValue(App.Path & "\settings.ini", "Default", "EULA")
' End user agreement
If Reula = "Agreed" Then
CheckBox3.Value = xtpChecked
Else
CheckBox3.Value = xtpUnchecked
End If
  
   ' Eula Text
   text3.Text = "Galaxy Crypt V6.0" & vbCrLf & _
"User Licence Agreement & Terms Of Use" & vbCrLf & _
"April 1, 2010" & vbCrLf & vbCrLf & _
"By clicking accept you state that you are obligated to follow the rules, statements, and requests implemented in this End User License Agreement. Not reading, understanding, or any variation of the words stated prior does not void this End User License Agreement. Violation of this End User License Agreement either partial or as a whole will result in a federal investigation and legal actions will be implemented." & vbCrLf & _
"1. No source code, either general or specific or any variation of this word is to be distributed via any form of communication. This includes but is not limited to the internet, written, or spoken." & vbCrLf & _
"2. This program/software is intended for educational use, or securing your home files. Misuse of this product (anything other than the above stated) will result in a voided warranty. This includes but is not limited to securing your files with intent to bypass security or anti-virus software/programs, harm one's computer, or any action leading to the harm of another person(s) computer or hardware." & vbCrLf & _
"3. All sales are final and refunding is at the discretion of the program designer(s). No refunds are required. By purchasing this product and accepting this End User License Agreement you state that you fully understand and have read this End User License Agreement."
text3.MaxLength = Len(text3.Text)
text3.Locked = True

 ' Auto sign-in
   If CheckBox4.Value = xtpChecked Then
   Me.Show
   Delay (0.2)
   Call Image3_MouseUp(1, 1, 0, 0)
   Else
   End If
   
End Sub

Private Sub CheckBox1_Click()
WriteIniValue App.Path & "\" & "settings.ini", "Default", "SavePass", CheckBox1.Value

End Sub

Private Sub CheckBox2_Click()
WriteIniValue App.Path & "\" & "settings.ini", "Default", "RememberMe", CheckBox2.Value
If CheckBox2.Value = xtpUnchecked Then
CheckBox1.Value = xtpUnchecked
End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SettedX = X
SettedY = Y
Dragging = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image2.Visible = True
Image8.Visible = False
Image9.Visible = False
Image7.Visible = True
Image10.Visible = True

 If Dragging Then
        Me.Left = Me.Left + (X - SettedX)
        Me.Top = Me.Top + (Y - SettedY)
    End If

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image2.Visible = False
End Sub
Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Visible = False
Image9.Visible = True
Image10.Visible = False
Image7.Visible = True
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image3.Visible = False
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo error
Image2.Visible = False
Image3.Visible = True
Dim Cuerda$
Dim C%, Count%, Found%

EndApp = 0

If FlatEdit1.Text = vbNullString Or FlatEdit2.Text = vbNullString Then
Call Found_Error
Exit Sub
End If

Call DisableAll

Label4.Caption = "Recording selected settings..."
WriteIniValue App.Path & "\settings.ini", "Default", "User", FlatEdit1.Text
WriteIniValue App.Path & "\settings.ini", "Default", "Pass", FlatEdit2.Text

uLower = Md5.DigestStrToHexStr(LCase(FlatEdit1.Text))
Dim lret As Long
lret = URLDownloadToFile(0, "http://host3266.net/coderscentral/Usernames.txt", Environ("Temp") & "\Usernames.txt", 0, 0)
 

Open Environ("Temp") & "\Usernames.txt" For Input As #1

Do
    Count% = Count% + 1
    Line Input #1, sAll
    C% = InStr(sAll, uLower)
        If C% <> 0 Then
            Found% = True
        End If


Loop Until EOF(1) Or Found%

Close #1

If Found% Then
Found% = True
Label4.Caption = "Username confirmed..."
DelayLbl
Label4.Caption = "Confirming Password..."
StageTwo

If EndApp = 1 Then EndApp = 0: Exit Sub
GrantAccess
Me.WindowState = vbMinimized
Exit Sub
Else
Call Found_Error
Exit Sub
End If

error:

Debug.Print Err.Description
Debug.Print Err.Number

Dim IntResponse As Integer

If Err.Number = "53" Then
IntResponse = MsgBox("Cannot locate the database, please contact the vendor with this issue", vbCritical + vbRetryCancel, "No Database")

If IntResponse = vbRetry Then
Call Image3_MouseUp(1, 1, 0, 0)
Else
EnableAll
Exit Sub
End If
End If

End Sub


Private Sub Found_Error()

On Local Error Resume Next
DoEvents

Label4.Caption = "Error: Invalid username/password combination!"
Label4.ForeColor = vbRed
Call PlayWaveRes("ERROR", SND_ASYNC)
Sleep (1000)
Call EnableAll
Delay (2)
CheckBox2.Value = xtpUnchecked
CheckBox1.Value = xtpUnchecked
CheckBox4.Value = xtpUnchecked
Call form_load
End Sub

Private Sub EnableAll()
CheckBox1.Enabled = True
CheckBox2.Enabled = True
CheckBox3.Enabled = True
CheckBox4.Enabled = True
Image3.Enabled = True
Image2.Enabled = True
FlatEdit1.Locked = False
FlatEdit2.Locked = False
End Sub
Private Sub DisableAll()
CheckBox1.Enabled = False
CheckBox2.Enabled = False
CheckBox3.Enabled = False
CheckBox4.Enabled = False
Image3.Enabled = False
Image2.Enabled = False
FlatEdit1.Locked = True
FlatEdit2.Locked = True
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image8.Visible = True
Image7.Visible = False
Image10.Visible = True
End Sub
Private Sub Image8_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image9_Click()
SystemParametersInfo SPI_SETCURSORS, 0&, ByVal 0&, (SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

End
End Sub
Private Function StageTwo()

Dim Cuerda$
Dim C%, Count%, Found%

uLower = Md5.DigestStrToHexStr(LCase(FlatEdit2.Text))
Dim lret As Long
lret = URLDownloadToFile(0, "http://host3266.net/coderscentral/Passwords.txt", Environ("Temp") & "\Passwords.txt", 0, 0)
 

Open Environ("Temp") & "\Passwords.txt" For Input As #1

sAll = ""
Do
    Line Input #1, sAll
    C% = InStr(sAll, uLower)
        If C% <> 0 Then
            Found% = True
        End If
        
Loop Until EOF(1) Or Found%

Close #1

If Found% Then
DelayLbl
Label4.Caption = "Password confirmed..."
DelayLbl
Exit Function
Else
EndApp = 1
Call Found_Error
Exit Function
End If
End Function

Private Function GrantAccess()
On Error Resume Next
Dim Sound_Path As String

Label4.Caption = "Inputed information confirmed..."
DelayLbl
Label4.Caption = "Loading preselected runtime settings..."
DelayLbl
Label4.Caption = "Finalizing loading sequence..."
DelayLbl
Label4.Caption = "Exploring Blackholes..."
DelayLbl
Label4.Caption = "Welcome to Galaxy Crypt!"

resData = LoadResData("Startup", "Custom")
Open Environ$("Temp") & "\startup.wav" For Binary As #1
Put #1, , resData
Close #1

Sound_Path = Environ$("Temp") & "\Startup.wav"
Call Initialize_Mci(Sound_Path, "Startup")
Call Terminate_Mci("Startup")
Call EnableAll
Main.Visible = True
Main.Label7.Caption = "Galaxy Crypt Private Edition: " & FlatEdit1.Text
Me.Visible = False
Close #1
Me.Visible = False
End Function

' Random Delay Labels
Private Sub DelayLbl()
DoEvents
Delay (RndDecimal)
End Sub

Private Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String

    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)

    ByteArrayToString = sAns
End Function

Private Sub Label1_Click()
CheckBox1.Value = xtpChecked
End Sub

Private Sub Label2_Click()
CheckBox2.Value = xtpChecked
End Sub

Private Sub Label3_Click()
CheckBox3.Value = xtpChecked
End Sub

Private Sub Label8_Click()
CheckBox4.Value = xtpChecked
End Sub
Public Sub Terminate_Mci(ByVal sAlias As String)
On Local Error Resume Next
mciSendString "Stop " & sAlias, 0, 0, 0
mciSendString "Close " & sAlias, 0, 0, 0
Kill Environ$("Temp") & "\startup.wav"
End Sub

Public Sub Initialize_Mci(ByVal FilePath As String, sAlias As String)
mciSendString "Stop " & sAlias, 0, 0, 0
mciSendString "Close " & sAlias, 0, 0, 0
mciSendString "open " & FilePath & " Type MPEGVIDEO alias " & sAlias, 0, 0, 0
mciSendString "Play " & sAlias & " from 0", 0, 0, 0
Delay (2.5)
End Sub


