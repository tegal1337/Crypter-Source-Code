VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProCrypter v0.01"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "File"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdLoadstub 
         Caption         =   "..."
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtStub 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   4215
      End
      Begin VB.CommandButton cmdOutput 
         Caption         =   "..."
         Height          =   255
         Left            =   5280
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdInput 
         Caption         =   "..."
         Height          =   255
         Left            =   5280
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Stub file:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Output file:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Input file:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu mnuProcessMenu 
      Caption         =   "Processes"
      Visible         =   0   'False
      Begin VB.Menu mnuProcess 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuProcess 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuProcess 
         Caption         =   "Remove"
         Index           =   2
      End
   End
   Begin VB.Menu mnuMessagesMenu 
      Caption         =   "Messages"
      Visible         =   0   'False
      Begin VB.Menu mnuMessages 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuMessages 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuMessages 
         Caption         =   "Remove"
         Index           =   2
      End
   End
   Begin VB.Menu mnuP2PMenu 
      Caption         =   "P2P"
      Visible         =   0   'False
      Begin VB.Menu mnuP2P 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuP2P 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuP2P 
         Caption         =   "Remove"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long: hwndOwner As Long: hInstance As Long: lpstrFilter As String: lpstrCustomFilter As String: nMaxCustFilter As Long: nFilterIndex As Long: lpstrFile As String: nMaxFile As Long: lpstrFileTitle As String: nMaxFileTitle As Long: lpstrInitialDir As String: lpstrTitle As String: flags As Long: nFileOffset As Integer: nFileExtension As Integer: lpstrDefExt As String: lCustData As Long: lpfnHook As Long: lpTemplateName As String
End Type

Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String

Private Sub cmdInput_Click()
    txtInput = Browse("exe")
End Sub
Private Sub cmdLoadstub_Click()
    txtStub = Browse("exe")
End Sub

Private Sub cmdOutput_Click()
    txtOutput = ShowSaveFileDialog("All files *.*", "exe", App.Path, OFN_OVERWRITEPROMPT)
End Sub

Private Sub txtInput_Click()
    txtInput = Browse("exe")
End Sub

Private Sub txtOutput_Click()
    txtOutput = ShowSaveFileDialog("All files *.*", "exe", App.Path, OFN_OVERWRITEPROMPT)
End Sub

Private Sub txtStub_Click()
    txtStub = Browse("exe")
End Sub

Private Sub cmdBuild_Click()
    Dim bufFile() As Byte
    
    If txtInput = vbNullString Then MsgBox "No inputfile selected", vbCritical: Exit Sub
    If FileExists(txtInput) = False Then MsgBox "Inputfile doesn't exists", vbCritical: Exit Sub
    If txtOutput = vbNullString Then MsgBox "No ouputfile selected", vbCritical: Exit Sub
    If txtInput = txtOutput Then MsgBox "Choose another output path", vbCritical: Exit Sub
    If txtStub = vbNullString Then MsgBox "No stub selected", vbCritical: Exit Sub
            
    Call FileCopy(txtStub, txtOutput) 'Prepare the stub
    bufFile = ReadFile(txtInput) 'Pass the read file to bufFile
    Call bufXOR(bufFile) 'XOR file-buffer (bufFile), byte by byte

    Call UpdateRes(bufFile, txtOutput, "4", "40") 'Put the encrypted file as a resource in the stub
    Call CleanEOF(txtOutput) 'Overwrite "PADDINGX" at EOF
            
    MsgBox "File successfully encrypted", vbInformation
End Sub

Private Function ReadFile(ByVal strFile As String) As Byte()
    Dim ff As Long
    Dim bufFile() As Byte
    
    ff = FreeFile
    ReDim bufFile(0 To FileLen(strFile) - 1)
    Open strFile For Binary Access Read As ff
        Get ff, , bufFile
    Close ff
    
    ReadFile = bufFile
End Function

Private Function bufXOR(bufData() As Byte)
    For i = 0 To UBound(bufData)
        bufData(i) = bufData(i) Xor (i Mod 255)
    Next i
End Function

Private Sub UpdateRes(ByRef bufSrc() As Byte, ByVal strTarget As String, ByVal resType As String, ByVal resID As String)
    Dim lngSize As Long
    lngSize = UBound(bufSrc) + 1
        
    Dim lngHandle As Long, ret As Long
    lngHandle = BeginUpdateResource(strTarget, 0)
    Call UpdateResource(lngHandle, resType, resID, 1, bufSrc(0), lngSize)
    Call EndUpdateResource(lngHandle, 0)
End Sub
    
Private Sub CleanEOF(ByVal strFile As String)
    Dim lngPos As Long
        
    strFileBuf = StrConv(ReadFile(strFile), vbUnicode)
    lngPos = InStr(1, strFileBuf, "PADDINGX")

    ReDim bufNull(0 To FileLen(strFile) - lngPos) As Byte
    ff = FreeFile
        
    Open strFile For Binary Access Write As ff
        Seek ff, lngPos
        Put ff, , bufNull
    Close ff
End Sub
        
Function FileExists(ByVal strFile As String) As Boolean
    On Error GoTo Err
    FileExists = (GetAttr(strFile) And vbDirectory) = 0
    Exit Function
Err: FileExists = False
End Function

Public Function Browse(strFilter As String) As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = frmMain.hWnd
    OpenFile.hInstance = App.hInstance
    sFilter = "File (*." & strFilter & ")" & Chr(0) & "*." & UCase(strFilter) & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = App.Path & "\"
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        Exit Function
    Else
        Browse = Trim(OpenFile.lpstrFile)
    End If
End Function

Public Function ShowSaveFileDialog(ByVal sFilter As String, Optional ByVal sDefExt As String, Optional ByVal sInitDir As String, Optional ByVal lFlags As Long, Optional ByVal hParent As Long) As String
    Dim OFN As OPENFILENAME
    On Error Resume Next
    
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hParent
        .lpstrFilter = Replace(sFilter, "|", vbNullChar) & vbNullChar
        .lpstrFile = Space$(255) & vbNullChar & vbNullChar
        .nMaxFile = Len(.lpstrFile)
        .flags = lFlags
        .lpstrInitialDir = sInitDir
        .lpstrDefExt = sDefExt
    End With
    
    If GetSaveFileName(OFN) Then
        ShowSaveFileDialog = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, vbNullChar) - 1)
    End If
End Function
