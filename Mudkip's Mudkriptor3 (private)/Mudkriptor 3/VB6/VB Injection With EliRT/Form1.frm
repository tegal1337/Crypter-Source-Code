VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Load executable"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "Load executable"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Aggressor:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Victim:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" () 'For XP style

Private Function LoadFile(ByVal sName As String) As Byte()
   Dim nFile As Integer
   Dim arrFile() As Byte
   nFile = FreeFile
   Open sName For Binary As #nFile
      ReDim arrFile(LOF(nFile) - 1)
      Get #nFile, , arrFile
   Close #nFile
   LoadFile = arrFile
End Function

Private Sub Command1_Click()
   RunExe Text1(0).Text, LoadFile(Text1(1).Text)
End Sub

Private Sub Command2_Click(Index As Integer)
  Dim sExe As String
  sExe = GetFileName(Text1(Index).Text, "Executables|*.exe")
  If sExe <> "" Then Text1(Index) = sExe
End Sub

Private Sub Form_Initialize()
   InitCommonControls
End Sub

Private Sub Form_Load()
   Text1(0) = Environ$("COMSPEC")
   Text1(1) = Environ$("WINDIR") & "\system32\calc.exe"
   Command1.Caption = "Run exe from byte array!"
   Caption = "RunPE Demo"
End Sub

Private Sub Text1_Change(Index As Integer)
   Dim bEnable As Boolean
   bEnable = Trim(Text1(0).Text) <> ""
   bEnable = bEnable And (Dir(Text1(1).Text) <> "")
   Command1.Enabled = bEnable
End Sub
