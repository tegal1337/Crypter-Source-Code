VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   14370
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "String Encrypt"
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   4815
      Begin VB.TextBox rotf 
         Height          =   1695
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Form1.frx":0000
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox rots 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Generate"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox rotn 
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Text            =   "10"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Rot"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Typevar"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin MSComDlg.CommonDialog opendia 
      Left            =   6840
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   9375
      Left            =   5160
      TabIndex        =   6
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton Command3 
         Caption         =   "Open file"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   8880
         Width           =   1335
      End
      Begin VB.TextBox fbox 
         Height          =   8535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Clear"
      Height          =   375
      Left            =   0
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   5160
      Width           =   4935
   End
   Begin VB.CommandButton fakeFuncs 
      Caption         =   "Fake Functions"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dim"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox num 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Text            =   "5"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox output 
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton fakeConsts 
      Caption         =   "Const"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function randConsts() As String
    On Error Resume Next
    Dim varname     As String
    Dim varval      As String
    Dim vartype     As String
    Dim rand        As Integer
    Dim i           As Integer
    
    For i = 0 To Random(3, 15)
            varname = varname & Chr(Random(97, 122))
    Next i
    
    
    rand = Random(1, 3)
    Select Case (rand)
    'String
    Case 1
        vartype = "String"
        varval = Chr(34)
        For i = 0 To Random(3, 15)
                varval = varval & Chr(Random(97, 122))
        Next i
        varval = varval & Chr(34)
    'Number
    Case 2
        vartype = "Long"
        varval = Str(Random(10, 30000))
    'String array
    Case 3
        vartype = "String"
        varval = "array("
        For i = 0 To Random(2, 25)
            varval = varval & Chr(34)
            For x = 0 To Random(3, 15)
                    varval = varval & Chr(Random(97, 122))
            Next x
            varval = varval & Chr(34) & ","
        Next i
        varval = varval & Chr(34) & "a" & Chr(34) & ")"
        
    'Integer array
    Case 4
        vartype = "Integer"
        varval = "array("
        For i = 0 To Random(3, 15)
            varval = varval & Random(10, 200) & ","
        Next i
        varval = varval & "0)"
        
        
    Case Default
    End Select
    
    randConsts = "const " & varname & " as " & vartype & " = " & varval
End Function
Private Function fakeVars() As String
    On Error Resume Next
    Dim types()     As Variant
    Dim varname     As String
    Dim vartype     As String
    For i = 0 To Random(3, 15)
    
            varname = varname & Chr(Random(97, 122))
    Next i
    
    types = Array("integer", "string", "byte", "long", "object", "OLE_COLOR", "OLE_CANCELBOOL", "OLE_HANDLE", "OLE_OPTEXCLUSIVE", _
    "variant", "boolean", "currency", "date", "double")
    i = Random(0, 14)
    vartype = types(i)
    
    fakeVars = "dim " & varname & " as " & vartype
End Function
Private Function fakeTypevar()
    On Error Resume Next
    Dim types()     As Variant
    Dim varname     As String
    Dim vartype     As String
    
    For i = 0 To Random(3, 15)
            varname = varname & Chr(Random(97, 122))
    Next i
    
    types = Array("integer", "string", "byte", "long", "object", "OLE_COLOR", "OLE_CANCELBOOL", "OLE_HANDLE", "OLE_OPTEXCLUSIVE", _
    "variant", "boolean", "currency", "date", "double")
    i = Random(0, 14)
    vartype = types(i)
    
    fakeTypevar = varname & " as " & vartype
End Function


Private Sub Command5_Click()
    Dim t As String
    t = rots.Text
    
    rot t, rotn.Text
    rotf.Text = "rot(" & Chr(34) & t & Chr(34) & ")"
End Sub

Private Sub rot(ByRef x As String, n As Integer)
    Dim i       As Long
   
    For i = 1 To Len(x)
        Mid$(x, i, 1) = Chr$(Asc(Mid$(x, i, 1)) - n)
    Next i
End Sub

Private Sub fakeConsts_Click()
    For i = 1 To num.Text
        output.Text = output.Text & randConsts & vbCrLf
    Next i
End Sub
Private Sub Command1_Click()
    For i = 1 To num.Text
        output.Text = output.Text & fakeVars & vbCrLf
    Next i
End Sub
Private Sub Command4_Click()
    For i = 1 To num.Text
        output.Text = output.Text & fakeTypevar & vbCrLf
    Next i
End Sub
Private Sub Command2_Click()
    output.Text = ""
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    Dim file As String
    opendia.ShowOpen
    Open opendia.FileName For Binary As #1
    file = Space(LOF(1))
    Get #1, , file
    Close #1
    file = Replace(file, "  ", vbTab)
    
    fbox.Text = file
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim sel As Integer
    Select Case KeyCode
        Case 113
            For i = 1 To num.Text
                output.Text = output.Text & randConsts & vbCrLf
            Next i
            sel = fbox.SelStart
            fbox.Text = Mid(fbox.Text, 1, fbox.SelStart) & vbCrLf & output.Text & vbCrLf & Mid(fbox.Text, fbox.SelStart + 1)
            fbox.SelStart = sel + Len(output.Text)
            output.Text = ""
            
        Case 114
                    For i = 1 To num.Text
                output.Text = output.Text & fakeVars & vbCrLf
            Next i
            sel = fbox.SelStart
            fbox.Text = Mid(fbox.Text, 1, fbox.SelStart) & vbCrLf & output.Text & vbCrLf & Mid(fbox.Text, fbox.SelStart + 1)
            fbox.SelStart = sel + Len(output.Text)
            output.Text = ""
            
        Case 115
    End Select
End Sub
Function Random(Lowerbound As Long, Upperbound As Long)
    Randomize
    Random = Int((Upperbound - Lowerbound) * Rnd + Lowerbound)
End Function
