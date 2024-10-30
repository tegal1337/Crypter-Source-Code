VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIcon 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CommandButton cmdIco 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdExe 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtIco 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txtExe 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExe_Click()
    CD.Filename = ""
    
    CD.DialogTitle = "Seleccione el archivo ejecutable deseado"
    CD.Filter = "Programas Ejecutables (*.exe)|*.exe"
    CD.ShowOpen
    
    txtExe.Text = CD.Filename
End Sub

Private Sub cmdIco_Click()
    CD.Filename = ""
    
    CD.DialogTitle = "Select icon to set..."
    CD.Filter = "Iconos (*.ico)|*.ico"
    CD.ShowOpen
    
    txtIco.Text = CD.Filename
End Sub

Private Sub cmdChange_Click()
    Dim fallo As String
    
    If txtIco.Text = "" Or txtExe.Text = "" Then
        MsgBox "Set File & Icon", vbInformation, "Leaf´s iChanger"
        Exit Sub
    End If
    
    ReplaceIcons txtIco.Text, txtExe.Text, fallo
End Sub


