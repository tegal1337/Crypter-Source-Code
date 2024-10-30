VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   5160
      Width           =   735
   End
   Begin VB.CheckBox check4 
      Caption         =   "qojtokqvn"
      Height          =   255
      Left            =   9840
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "dbvgbwdiz"
      Height          =   255
      Left            =   8400
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "omgqlhnsk"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "dbvgbvdiz"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt | Decrypt XOR"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   4215
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Check1.Value = 1 Then Text2.Text = dbvgbvdiz(Text1.Text)
If Check2.Value = 1 Then Text2.Text = omgqlhnsk(Text1.Text)
If Check3.Value = 1 Then Text2.Text = dbvgbwdiz(Text1.Text)
If check4.Value = 1 Then Text2.Text = qojtokqvn(Text1.Text)
If Check5.Value = 1 Then Text2.Text = sqlvqmsxp(Text1.Text)

End Sub
