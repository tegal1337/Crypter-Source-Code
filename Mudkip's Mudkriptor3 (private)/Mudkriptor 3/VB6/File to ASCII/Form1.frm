VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "File to VB6 ASCII"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox output 
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse ..."
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox filepath 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog cdiag 
      Left            =   360
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Load File"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    cdiag.ShowOpen
    filepath.Text = cdiag.FileName
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim f As String
    Dim code As Integer
    Dim chrCount As Integer
    chrCount = 0

    Open filepath.Text For Binary As #1
    f = Space(LOF(1))
    Get #1, , f
    Close #1
    
    For i = 1 To Len(f)
        code = Asc(Mid(f, i, 1))
        output.Text = output.Text & " & chr(" & code & ")"
        chrCount = chrCount + 1
        If chrCount > 10 Then
            output.Text = output.Text & " _" & vbCrLf
            chrCount = 0
        End If
        
    Next i
End Sub
