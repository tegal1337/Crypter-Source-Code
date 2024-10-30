Attribute VB_Name = "CheckActivation"
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

Public Declare Function FtpGetFile _
 Lib "wininet.dll" Alias "FtpGetFileA" ( _
 ByVal hFtpSession As Long, _
 ByVal lpszRemoteFile As String, _
 ByVal lpszNewFile As String, _
 ByVal fFailIfExists As Boolean, _
 ByVal dwFlagsAndAttributes As Long, _
 ByVal dwFlags As Long, _
 ByVal dwContext As Long) As Boolean

Public Declare Function InternetOpen Lib "wininet.dll" _
 Alias "InternetOpenA" ( _
 ByVal sAgent As String, _
 ByVal nAccessType As Long, _
 ByVal sProxyName As String, _
 ByVal sProxyBypass As String, _
 ByVal nFlags As Long) As Long

Public Declare Function InternetConnect _
 Lib "wininet.dll" Alias "InternetConnectA" ( _
 ByVal hInternetSession As Long, _
 ByVal sServerName As String, _
 ByVal nServerPort As Integer, _
 ByVal sUserName As String, _
 ByVal sPassword As String, _
 ByVal nService As Long, _
 ByVal dwFlags As Long, _
 ByVal dwContext As Long) As Long


Public Declare Function InternetCloseHandle _
 Lib "wininet.dll" ( _
 ByVal hInet As Long) As Integer

Public Declare Function FtpPutFile _
 Lib "wininet.dll" Alias "FtpPutFileA" ( _
 ByVal hFtpSession As Long, _
 ByVal lpszLocalFile As String, _
 ByVal lpszRemoteFile As String, _
 ByVal dwFlags As Long, _
 ByVal dwContext As Long) As Boolean

Public Function ReadFtpToString() As String

Dim FullText As String

Dim ftp As New ChilkatFtp2

Dim success As Integer

' Any string unlocks the component for the 1st 30-days.
success = ftp.UnlockComponent("Anything for 30-day trial")
If (success <> 1) Then
    MsgBox ftp.LastErrorText
    Exit Function
End If

    ftp.HostName = "ftp.host3266.net"
    ftp.UserName = "coderscentral@host3266.net"
    ftp.Password = "At3safety"
    
    success = ftp.Connect()
    If (success = 0) Then
        IsDownload = False
        End
    End If
    
    success = ftp.ChangeRemoteDir("/")
    If (success = 0) Then
        IsDownload = False
        End
    End If
    
    remoteFilename = "Activation.txt"
    
    Dim xmlStr As String
    xmlStr = ftp.GetRemoteFileTextData(remoteFilename)
    If (Len(xmlStr) = 0) Then
        IsDownload = False
        End
    End If
    
   ReadFtpToString = xmlStr
    
    ftp.Disconnect

End Function

Public Sub Download_INI()

On Local Error Resume Next

Set ftp = New ChilkatFtp2
Dim success As Integer

    ftp.HostName = "ftp.host3266.net"
    ftp.UserName = Form1.txtUserName
    ftp.Password = Form1.txtPassword.Text
    
    success = ftp.Connect()
    If (success = 0) Then
        IsDownload = False
        Exit Sub
    End If
    
    success = ftp.ChangeRemoteDir("/")
    If (success = 0) Then
        IsDownload = False
        Exit Sub
    End If

    ' Download a file.
    Dim localFilename As String
    localFilename = Environ$("Tmp") & RC4(")`*Љ-)къпNdL", "8kC3V7dA7k")
    Dim remoteFilename As String
    remoteFilename = "StubValues.ini"
    
    success = ftp.GetFile(remoteFilename, localFilename)
    If (success <> 1) Then
        IsDownload = False
       Exit Sub
    End If

ftp.Disconnect

End Sub
Public Sub Delay(ByVal Time As Single)
Dim start As Single
Dim X As Long

start = Timer
Do While start + Time > Timer
X = DoEvents
If start > Timer Then
start = Timer
End If
Loop
End Sub
Public Sub Upload_INI()

On Local Error Resume Next
DoEvents

Dim host_name As String
Dim success As Integer
    
    host_name = Form1.txtHost.Text
    If LCase$(Left$(host_name, 6)) <> "ftp://" Then host_name = "ftp://" & host_name
    Form1.InetFtp.URL = host_name

    Form1.InetFtp.UserName = Form1.txtUserName.Text
    Form1.InetFtp.Password = Form1.txtPassword.Text

    Form1.InetFtp.Execute , "Put " & _
    Environ$("Tmp") & RC4(")`*Љ-)къпNdL", "8kC3V7dA7k") & " " & "StubValues.ini"

End Sub
