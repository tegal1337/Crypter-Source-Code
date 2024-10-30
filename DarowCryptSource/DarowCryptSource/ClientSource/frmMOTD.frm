VERSION 5.00
Begin VB.Form frmMOTD 
   Caption         =   "Message of the Day"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMOTD 
      Enabled         =   0   'False
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long)
Private Declare Function InternetOpenA Lib "wininet.dll" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrlA Lib "wininet.dll" (ByVal hOpen As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Sub InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Const UserAgent = "drizzle@ymail.com"
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Sub Form_Load()
Call MOTD
End Sub

Public Function MOTD() As String
On Error Resume Next
  Dim hUrl As Long
  Dim hOpen As Long
  Dim szData As String
  Dim lNull As Long

 hOpen = InternetOpenA("USER_CHECK", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If hOpen <> 0 Then
  hUrl = InternetOpenUrlA(hOpen, "http://www.xblhackers.com/Infinity/test.txt", 0, 0, INTERNET_FLAG_EXISTING_CONNECT, 0)
If hUrl <> 0 Then
      szData = Space(1000)
      Call InternetReadFile(hUrl, szData, 1000, lNull)
      If InStr(szData, "<newline>") Then
      szData = Replace(szData, "<newline>", vbNewLine)
      End If
      txtMOTD.Text = szData
      End If
     Call InternetCloseHandle(hUrl)
    End If
   Call InternetCloseHandle(hOpen)
End Function

