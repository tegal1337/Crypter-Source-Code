VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Changer"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeIcon 
      Caption         =   "Change Icon"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3600
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseExe 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdBrowseIco 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtExe 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "ExeFile"
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtIco 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "IcoFile"
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowseExe_Click()
    With dlg
        .DialogTitle = "Select Exe File..."
        .Filter = "Executable Files (*.exe)|*.exe"
        .ShowOpen
    End With
    
    txtExe.Text = dlg.FileName
End Sub

Private Sub cmdBrowseIco_Click()
    With dlg
        .DialogTitle = "Select Icon File..."
        .Filter = "Icons (*.ico)|*.ico"
        .ShowOpen
    End With
    
    txtIco.Text = dlg.FileName
End Sub

Private Sub cmdChangeIcon_Click()
    If ChangeIcon(txtExe.Text, txtIco.Text) Then
        MsgBox "Done"
    Else
        MsgBox "Error Occurred."
    End If
End Sub
