VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6300
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   47
      Left            =   1440
      TabIndex        =   63
      Text            =   "sSplit2"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   46
      Left            =   1440
      TabIndex        =   62
      Text            =   "lLen"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   1440
      TabIndex        =   61
      Text            =   "sBuffer"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   1440
      TabIndex        =   60
      Text            =   "lhKey"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   1440
      TabIndex        =   59
      Text            =   "lAfter"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   1440
      TabIndex        =   58
      Text            =   "lBefore"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   1440
      TabIndex        =   57
      Text            =   "bFound"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   1440
      TabIndex        =   56
      Text            =   "sWinSerial"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   1440
      TabIndex        =   55
      Text            =   "sComputer"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   1440
      TabIndex        =   54
      Text            =   "sUser"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   1440
      TabIndex        =   53
      Text            =   "aSerials"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   1440
      TabIndex        =   52
      Text            =   "aHDDs"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   1440
      TabIndex        =   51
      Text            =   "aDlls"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   1440
      TabIndex        =   50
      Text            =   "aComputers"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   1440
      TabIndex        =   49
      Text            =   "aUsers"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   6
      Left            =   2880
      TabIndex        =   48
      Text            =   "sAnti"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   47
      Text            =   "mAnti"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   46
      Text            =   "Form1"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   1440
      TabIndex        =   45
      Text            =   "Params"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   1440
      TabIndex        =   44
      Text            =   "zDoNotCall"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   1440
      TabIndex        =   43
      Text            =   "sProc"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   1440
      TabIndex        =   42
      Text            =   "GetMod"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   1440
      TabIndex        =   41
      Text            =   "bvBuff"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   1440
      TabIndex        =   40
      Text            =   "hProc"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   1440
      TabIndex        =   39
      Text            =   "sHost"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   1440
      TabIndex        =   38
      Text            =   "sLib"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   5
      Left            =   2880
      TabIndex        =   37
      Text            =   "LoadLibrary"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   0
      TabIndex        =   36
      Text            =   "KeyEnc"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   4
      Left            =   2880
      TabIndex        =   35
      Text            =   "ResolveForward"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   34
      Text            =   "GetProcAddress"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   33
      Text            =   "PatchCall"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   32
      Text            =   "Invoke"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   31
      Text            =   "RunPE"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   29
      Text            =   "ProjectName"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   28
      Text            =   "Exename"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   0
      TabIndex        =   27
      Text            =   "Key"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   0
      TabIndex        =   26
      Text            =   "F"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   0
      TabIndex        =   25
      Text            =   "Password"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   0
      TabIndex        =   24
      Text            =   "Data"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   0
      TabIndex        =   23
      Text            =   "RC4"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   22
      Text            =   "cNtPEL"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   21
      Text            =   "mMain"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   0
      TabIndex        =   19
      Text            =   "tCONTEXT"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   0
      TabIndex        =   18
      Text            =   "tPROCESS_INFORMATION"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   0
      TabIndex        =   17
      Text            =   "tSTARTUPINFO"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   0
      TabIndex        =   16
      Text            =   "tIMAGE_SECTION_HEADER"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   0
      TabIndex        =   15
      Text            =   "tIMAGE_NT_HEADERS"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   0
      TabIndex        =   14
      Text            =   "tIMAGE_DOS_HEADER"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   0
      TabIndex        =   13
      Text            =   "lMod"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   0
      TabIndex        =   12
      Text            =   "lNTDll"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   0
      TabIndex        =   11
      Text            =   "lKernel"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Text            =   "i"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Text            =   "c_bvASM"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Text            =   "c_lOldVTE"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   0
      TabIndex        =   7
      Text            =   "c_lVTE"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Text            =   "c_bInit"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Text            =   "c_lLoadLib"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Text            =   "c_lKrnl"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Text            =   "sDelim"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Text            =   "sData "
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Text            =   "cRPE"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Functions"
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Filenames"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Strings:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuRandomize 
         Caption         =   "Randomize"
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Create"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const X = """"
Private Sub mnuCreate_Click()
Open App.Path & "\Generated\" & frmMain.Text2(0).Text & ".bas" For Binary As #1
Put #1, , mMain & Enc
Close #1

Open App.Path & "\Generated\" & frmMain.Text2(4).Text & ".frm" For Binary As #1
Put #1, , frm
Close #1

Open App.Path & "\Generated\" & frmMain.Text2(5).Text & ".bas" For Binary As #1
Put #1, , sAnti
Close #1

Open App.Path & "\Generated\" & frmMain.Text2(1).Text & ".cls" For Binary As #1
Put #1, , ntPE
Close #1

Open App.Path & "\Generated\" & frmMain.Text2(3).Text & ".vbp" For Binary As #1
Put #1, , ProjectSettings
Close #1

End Sub

Private Sub mnuRandomize_Click()
Dim i As Integer
For i = 0 To 47
Text1(i).Text = lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber)
Next i
For i = 0 To 5
Text2(i).Text = lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber)
Next i
For i = 0 To 6
Text3(i).Text = lRan(RandomNumber) & lRan(RandomNumber) & lRan(RandomNumber)
Next i
End Sub

