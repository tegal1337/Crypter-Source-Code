VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form frmHaupt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Split file into parts and add it to Resource"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmHaupt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox boxHaupt 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _Version        =   851970
      _ExtentX        =   8281
      _ExtentY        =   1720
      _StockProps     =   79
      Appearance      =   1
      Begin XtremeSuiteControls.PushButton buttAbout 
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   600
         Width           =   1575
         _Version        =   851970
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "About"
         Appearance      =   1
      End
      Begin XtremeSuiteControls.PushButton buttPack 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   1575
         _Version        =   851970
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pack file..."
         Appearance      =   1
      End
      Begin XtremeSuiteControls.UpDown udParts 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Select size of the parts that file gets splittet into in bytes"
         Top             =   600
         Width           =   375
         _Version        =   851970
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   64
         Appearance      =   1
         UseVisualStyle  =   0   'False
         BuddyControl    =   ""
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtSelectFile 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         _Version        =   851970
         _ExtentX        =   7858
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Drop file or click me..."
         Appearance      =   1
         UseVisualStyle  =   0   'False
         OLEDropMode     =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtParts 
         Height          =   255
         Left            =   480
         TabIndex        =   3
         ToolTipText     =   "Select size of the parts that file gets splittet into in bytes"
         Top             =   600
         Width           =   615
         _Version        =   851970
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "1000"
         Alignment       =   1
         Appearance      =   1
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.CommonDialog cdHaupt 
      Left            =   120
      Top             =   3000
      _Version        =   851970
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
End
Attribute VB_Name = "frmHaupt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' for clean code. i hope its clean

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long ' to detect if a file exists

Private Function fSelectFile() As String ' function to select a file with commondialog
    
    With cdHaupt
     .FileName = vbNullString
     .InitDir = App.Path
     .CancelError = False
     .ShowOpen
     If .FileName = vbNullString Then Exit Function
     fSelectFile = .FileName
    End With
    
End Function
Private Function fSaveFile() As String ' function to save a file with commondialog
    
    With cdHaupt
     .FileName = "out"
     .DefaultExt = "exe"
     .InitDir = App.Path
     .CancelError = False
     .ShowSave
     If .FileName = vbNullString Then Exit Function
     fSaveFile = .FileName
    End With
    
End Function

Private Sub buttAbout_Click()
    
    MsgBox "Credits to: " & vbCrLf & _
           "ap0calypse" & vbCrLf & _
           "Karcrack" & vbCrLf & _
           "krizhiel"
           
End Sub

Private Sub buttPack_Click()
 Dim bStub() As Byte ' Variable for storing our stubbytearray
 Dim bFile() As Byte
 Dim sStubPath As String ' Variable to setup our stubpath
 Dim sFile As String ' Variable which gets cryptet and splittet into parts
 Dim sStub As String ' Variable for storing our stub
 Dim sSaveFile As String ' Variable to set our savefile
 Dim sTempText As String ' ' Variable to store the parts we splittet
 Dim sCount As String ' Variable to store splitcount
 Dim iFF As Integer
 Dim i As Integer
    
    sSaveFile = fSaveFile ' set our savefile
    sStubPath = App.Path & "\stub.exe" ' set our stubpath
    
    'if stub.exe already exist delet it
    If PathFileExists(sStubPath) Then
     Kill sStubPath
    End If
    
    'get our stub from resourcefile with the id 1 in resourcetype "STB" and put it into bStub
    iFF = FreeFile
    Open sStubPath For Binary Access Write As iFF
     bStub = LoadResData(1, "STB")
     Put iFF, , bStub
    Close iFF
    
    
    'open stub.exe for reading and get stubdata into sStub
    iFF = FreeFile
    Open sStubPath For Binary Access Read As iFF
     sStub = Space(LOF(iFF))
     Get iFF, , sStub
    Close iFF
    
    'open our savefile and put sStub data in it
    iFF = FreeFile
    Open sSaveFile For Binary Access Write As iFF
     Put iFF, , sStub
    Close iFF
    
    
    'Get our file that has to be splitted and crypted
    iFF = FreeFile ' free our iFF
    Open txtSelectFile.Text For Binary Access Read As iFF ' open our file for reading
     sFile = Space(LOF(iFF)) ' make sFile as big as our file that we are going to crypt
     Get iFF, , sFile ' put our file into sFile variable
    Close iFF ' no comment
        
  
    ' Add your encryption here
    'sFile = RC4(sFile, "QIfOT87133U") ' encrypt file
    bFile = StrConv(sFile, vbFromUnicode)
    'Call Compress_Huffman_Dynamic(bFile)
    sFile = StrConv(bFile, vbUnicode)
    
    'get splitcount
    sCount = UBound(fSplitParts(sFile, CInt(txtParts.Text)))
    
    'put splitcount in RC_PART id 0
    Call PatchResource(sSaveFile, "0", "RCPART", sCount, 1033)
    
    'write parts to "RC_DATA"
    For i = 0 To UBound(fSplitParts(sFile, CInt(txtParts.Text))) - 1
    sTempText = vbNullString
    sTempText = fSplitParts(sFile, CInt(txtParts.Text))(i)
     Call PatchResource(sSaveFile, CStr(i), "RCDATA", sTempText, 1033)
    
    Next i
    
    
    If PathFileExists(sStubPath) Then
     Kill sStubPath
    End If
    
MsgBox sCount & " Parts written to:" & vbCrLf & sSaveFile, , "Done"
End Sub


Private Sub txtSelectFile_Click()
    
    txtSelectFile.Text = fSelectFile
    txtSelectFile.ToolTipText = txtSelectFile
    
End Sub

Private Sub txtSelectFile_OLEDragDrop(ByVal Data As XtremeSuiteControls.DataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    
    txtSelectFile.Text = Data.Files(1)
    txtSelectFile.ToolTipText = txtSelectFile.Text
    
End Sub

Private Sub udParts_DownClick()
 
 Dim sTempCount As String
    
    sTempCount = txtParts.Text
    If sTempCount <= 0 Then Exit Sub
    sTempCount = sTempCount - 50
    txtParts.Text = sTempCount
    
End Sub

Private Sub udParts_UpClick()
 
 Dim sTempCount As String
 
    sTempCount = txtParts.Text
    If sTempCount <= 0 Then
     sTempCount = 50
    Else
     sTempCount = sTempCount + 50
    End If
    
    txtParts.Text = sTempCount
    
End Sub
