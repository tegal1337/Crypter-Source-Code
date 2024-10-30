VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Leaf´s Crypter [www.forestmalware.blogspot.com]"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Harlow Solid Italic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAbout.frx":08CA
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1575
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Image Image2 
      Height          =   1500
      Left            =   120
      Picture         =   "frmAbout.frx":0973
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code by St0k                                 MSN: st0k@hotmail.es              WEB: www.forestmalwa.blogspor.com Program on: VB6"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   1695
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   1500
      Index           =   1
      Left            =   120
      Picture         =   "frmAbout.frx":1AA2
      Top             =   2160
      Width           =   1500
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
