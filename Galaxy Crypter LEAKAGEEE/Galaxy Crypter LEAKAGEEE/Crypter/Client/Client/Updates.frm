VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "CODEJO~3.OCX"
Begin VB.Form Updates 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "Updates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   3720
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   3240
      Width           =   6735
      _Version        =   851968
      _ExtentX        =   11880
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483641
      Scrolling       =   2
      Appearance      =   1
      UseVisualStyle  =   0   'False
      BarColor        =   16777088
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin XtremeSuiteControls.CheckBox ChkOpen 
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   4080
      Width           =   195
      _Version        =   851968
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   4080
      Width           =   3135
      _Version        =   851968
      _ExtentX        =   5530
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Open Folder Following Download"
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
   Begin VB.Image Image5 
      Height          =   285
      Left            =   10400
      Picture         =   "Updates.frx":F172
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image10 
      Height          =   285
      Left            =   10845
      Picture         =   "Updates.frx":121F5
      Top             =   20
      Width           =   420
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   10845
      Picture         =   "Updates.frx":15600
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   10400
      Picture         =   "Updates.frx":189F8
      Top             =   20
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   390
      Left            =   9120
      Picture         =   "Updates.frx":1BB6D
      Top             =   5040
      Width           =   1905
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   9120
      Picture         =   "Updates.frx":2058D
      Top             =   5040
      Width           =   1905
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3720
      Width           =   5415
      _Version        =   851968
      _ExtentX        =   9551
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Update Status (Pending)"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   8175
      _Version        =   851968
      _ExtentX        =   14420
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Label1"
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
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Updates.frx":24D38
      Top             =   0
      Width           =   11250
   End
End
Attribute VB_Name = "Updates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private SettedX As Integer, SettedY As Integer, Dragging As Boolean

Dim l_update As Boolean
Dim Message_Text As String

Private Sub Command3_Click()
End
End Sub

Private Sub form_load()


Image6.Visible = False
Image7.Visible = False
   
  'load the cached file list
GetCacheURLList

SecondTry

On Error Resume Next

ProgressBar1.Max = 100
l_update = True
Message_Text = "Browsing The Web For The Latest Version of Galaxy Crypter"
Me.Show

    With Label1
 
        .Caption = Message_Text & "."

Delay (3)
' ---------------------------------------------------------------------------------------------------------------

Dim StrTemp         As String
Dim m_update        As Boolean
Dim IntResponse     As Long

  .Caption = Message_Text & "."
       
lret = URLDownloadToFile(0, "http://host3266.net/coderscentral/CurVersion.txt", Environ("Temp") & "\CurVersion.txt", 0, 0)
 .Caption = Message_Text & "."
                

If lret = 0 Then
DoEvents
Open Environ$("Temp") & "\CurVersion.txt" For Input As #1

    Do While Not EOF(1)
         Line Input #1, StrTemp
         StrTemp = Replace(StrTemp, "MajorVer =", "")
            If StrTemp <= App.Major Then m_update = False Else m_update = True
            .Caption = Message_Text & ".."
    Loop
    
    .Caption = Message_Text & "..."
    
Close #1

If Build.Fileexists(Environ$("Temp") & "\CurVersion.txt") = True Then Kill Environ$("Temp") & "\CurVersion.txt"

If m_update = True Then
    .Caption = Message_Text & "..."
    CheckBox1.Value = xtpChecked
    Label4.Caption = "Update Status (Available)"
    Message_Text = "A Program Update Is Available For Download"
    .Caption = Message_Text & "..."
    Image6.Visible = True
Else
.Caption = "Your Version of Galaxy Crypter Is Current"
Message_Text = "Your Version of Galaxy Crypter Is Current"
Label4.Caption = "Update Status (Up To Date)"
Image6.Visible = False
Image7.Visible = False
End If

End If

   End With
   
EndUpdate:
End Sub

Public Function DirExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
DirExists = fs.FolderExists(OrigFile)
End Function

Private Sub FirstTry()
   Dim cachefile As String

  'delete the selected file
   cachefile = List1.List(List1.ListIndex)
   Call DeleteUrlCacheEntry(cachefile)

  'reload the list
   GetCacheURLList
End Sub

Private Sub SecondTry()

   Dim cachefile As String
   Dim i As Long

  'delete all files except..
   For i = 0 To List1.ListCount - 1

      cachefile = List1.List(i)

     '..if the file is a cookie, don't screw
     'up saved passwords, so skip it
      If InStr(cachefile, "Cookie") = 0 Then

         Call DeleteUrlCacheEntry(cachefile)

      End If

   Next

  'reload the list
   GetCacheURLList
End Sub

Public Sub GetCacheURLList()

   Dim ICEI As INTERNET_CACHE_ENTRY_INFO
   Dim hFile As Long
   Dim cachefile As String
   Dim dwBuffer As Long
   Dim pntrICE As Long

   List1.Clear

  'Like other APIs, calling FindFirstUrlCacheEntry or
  'FindNextUrlCacheEntry with an insufficient buffer will
  'cause the API to fail, and the buffer pointing to the
  'correct size required for a successful call.
   dwBuffer = 0

  'Call to determine the required buffer size
   hFile = FindFirstUrlCacheEntry(0&, ByVal 0, dwBuffer)

  'both conditions hould be met by the first call
   If (hFile = ERROR_CACHE_FIND_FAIL) And _
      (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then

     'The INTERNET_CACHE_ENTRY_INFO data type is a
     'variable-length type. It is neccessary to allocate
     'memnory for the result of the call and pass the
     'pointer to this memory location to the API.
      pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)

     'allocation successful
      If pntrICE Then

        'set a Long pointer to the memory location
         CopyMemory ByVal pntrICE, dwBuffer, 4

        'and call the first find API again passing the
        'pointer to the allocated memory
         hFile = FindFirstUrlCacheEntry(vbNullString, ByVal pntrICE, dwBuffer)

        'hfile should = 1 (success)
         If hFile <> ERROR_CACHE_FIND_FAIL Then

           'loop through the cache
            Do

              'the pointer has ben filled, so move the
              'data back into a ICEI structure
               CopyMemory ICEI, ByVal pntrICE, Len(ICEI)

              'CacheEntryType is a long representing
              'the type of entry returned
               If (ICEI.CacheEntryType And _
                   NORMAL_CACHE_ENTRY) = NORMAL_CACHE_ENTRY Then

                 'extract the string from the memory location
                 'pointed to by the lpszSourceUrlName member
                 'and add to a list
                  cachefile = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                  List1.AddItem cachefile

               End If

              'free the pointer and memory associated
              'with the last-retrieved file
               Call LocalFree(pntrICE)

              'and again repeat the procedure, this time calling
              'FindNextUrlCacheEntry with a buffer size set to 0.
              'This will cause the call to once again fail,
              'returning the required size as dwBuffer
               dwBuffer = 0
               Call FindNextUrlCacheEntry(hFile, ByVal 0, dwBuffer)

              'allocate and assign the memory to the pointer
               pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
               CopyMemory ByVal pntrICE, dwBuffer, 4

           'and call again with the valid parameters.
           'If the call fails (no more data), the loop exits.
           'If the call is successful, the Do portion of the
           'loop is executed again, extracting the data from
           'the returned type
            Loop While FindNextUrlCacheEntry(hFile, ByVal pntrICE, dwBuffer)

         End If 'hFile

      End If 'pntrICE

   End If 'hFile

  'clean up by closing the find handle, as
  'well as calling LocalFree again to be safe
   Call LocalFree(pntrICE)
   Call FindCloseUrlCache(hFile)

End Sub

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)

End Function


Private Sub Form_Unload(Cancel As Integer)
Me.Hide
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 SettedX = X
    SettedY = Y
    Dragging = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image2.Visible = False
Image5.Visible = True

Image4.Visible = False
Image10.Visible = True

If CheckBox1.Value = xtpChecked Then
    Image7.Visible = False
    Image6.Visible = True
Else
    Image7.Visible = False
    Image6.Visible = False
End If

If Dragging Then
Me.Left = Me.Left + (X - SettedX)
Me.Top = Me.Top + (Y - SettedY)
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
Image10.Visible = False
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = True
Image2.Visible = False
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = vbMinimized
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Visible = True
Image4.Visible = False
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Hide
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
Image5.Visible = False
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If DirExists(App.Path & "\Galaxy Update") = False Then MkDir App.Path & "\Galaxy Update"

    
    Label1.Caption = "Downloading Most Recent Version of Galaxy Crypter From The Web..."

Dim ftp As New ChilkatFtp2

Dim success As Integer

success = ftp.UnlockComponent("Anything for 30-day trial")
If (success <> 1) Then GoTo DlError


ftp.HostName = "ftp://ftp.host3266.net/"
ftp.Username = "coderscentral@host3266.net"
ftp.Password = "At3safety"

ftp.Passive = 1

success = ftp.Connect()
If (success <> 1) Then GoTo DlError

success = ftp.ChangeRemoteDir("/")
If (success <> 1) Then GoTo DlError

Dim localFilename As String
localFilename = App.Path & "\Galaxy Update\Update.rar"
Dim remoteFilename As String
remoteFilename = hwid1 & ".rar"


success = ftp.GetFile(remoteFilename, localFilename)
If (success <> 1) Then GoTo DlError


ftp.Disconnect

        MsgBox "Download Successful!" & vbNewLine & "Your new update can be located in: " & _
        App.Path & "\Galaxy Update\Client.rar", vbInformation, "Download Successful"
        
    If ChkOpen.Value = xtpChecked Then
        Shell ("Explorer.exe " & App.Path & "\Galaxy Update"), vbNormalFocus
    End If
    
        ProgressBar1.Scrolling = xtpProgressBarSmooth
        ProgressBar1.Value = 100
        Call Form_Unload(1)
        Exit Sub
        
DlError:

        MsgBox "There was an unexpected error during download", vbOKOnly + vbCritical, "Error"
        ProgressBar1.Scrolling = xtpProgressBarSmooth
        ProgressBar1.Value = 100
        Me.Visible = False
   
     
    Me.Visible = False
    
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True
Image6.Visible = False
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Visible = True
Image7.Visible = False
End Sub
