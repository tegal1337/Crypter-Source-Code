VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmManager 
   Caption         =   "File Manager"
   ClientHeight    =   6360
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10695
   Icon            =   "FrmManager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView Tvfolders 
      Height          =   6015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   10610
      _Version        =   327682
      Indentation     =   647
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImgEx"
      Appearance      =   1
   End
   Begin ComctlLib.ListView lstFiles 
      Height          =   6015
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File"
         Object.Width           =   5293
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   1305
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   6480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ComctlLib.StatusBar sbFiles 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6105
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11800
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3246
            MinWidth        =   3246
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3246
            MinWidth        =   3246
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImgEx 
      Left            =   0
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManager.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManager.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManager.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManager.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManager.frx":0D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmManager.frx":10A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnufolders 
      Caption         =   "<folders>"
      Begin VB.Menu Mnufol 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu Mnufol 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Mnufol 
         Caption         =   "Expand Tree"
         Index           =   2
      End
      Begin VB.Menu Mnufol 
         Caption         =   "Collapse Tree"
         Index           =   3
      End
      Begin VB.Menu Mnufol 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Mnufol 
         Caption         =   "Download folder"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu Mnufol 
         Caption         =   "New folder"
         Index           =   6
      End
      Begin VB.Menu Mnufol 
         Caption         =   "Delete"
         Index           =   7
      End
   End
   Begin VB.Menu MnuFiles 
      Caption         =   "<files>"
      Begin VB.Menu MnuBestanden 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu MnuBestanden 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuBestanden 
         Caption         =   "Download"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu MnuBestanden 
         Caption         =   "Upload"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu MnuBestanden 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuBestanden 
         Caption         =   "Execute"
         Index           =   6
         Begin VB.Menu MnuCute 
            Caption         =   "Hiden"
            Index           =   0
         End
         Begin VB.Menu MnuCute 
            Caption         =   "Normal"
            Index           =   1
         End
      End
      Begin VB.Menu MnuBestanden 
         Caption         =   "Rename"
         Index           =   7
      End
      Begin VB.Menu MnuBestanden 
         Caption         =   "Delete"
         Index           =   8
      End
   End
End
Attribute VB_Name = "FrmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************
'* Coded By: Aka.Ma
'* Non c? dio solo allah e Mohammed ? il messaggero di Allah
'* There is no God but allah ,and Mohammed is the messenger of Allah
'* Use this Source to learn from and not To Copy/Paste!!!!
'* E-mail: Akama.security@gmail.com
'*************************************
Dim Sother As String
Dim sPath As String
Dim sMap As String

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
FrmMain.Show
Me.Hide
End Sub

Private Sub lstFiles_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
 lstFiles.Sorted = True
 lstFiles.SortKey = ColumnHeader.Index - 1
 lstFiles.SortOrder = (lstFiles.SortOrder + 1) Mod 2
End Sub

Private Sub Form_Load()
   SendMessageA lstFiles.hWnd, &H1000 + 54, &H1, ByVal 1
   SendMessageA lstFiles.hWnd, &H1000 + 54, &H20, ByVal 1
  Set Fso = CreateObject("Scripting.FileSystemObject")
  FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "300100"
  '*****************************************************************************
With Picture1
.BackColor = vbWhite
.ScaleMode = vbPixels
.BorderStyle = 0
.Width = ScaleX(32, vbPixels, Me.ScaleMode)
.Height = ScaleY(32, vbPixels, Me.ScaleMode)
End With
End Sub

Private Sub BrowseFolder(sPath As String)
lstFiles.ListItems.Clear
lstFiles.Sorted = False
lstFiles.Enabled = False
FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "STRFLS" & sPath
sbFiles.Panels(1).Text = Left(Tvfolders.SelectedItem.FullPath, 2) & Sother
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MnuFiles
End Sub

Private Sub MnuBestanden_Click(Index As Integer)
Select Case Index
Case 0:
lstFiles.ListItems.Clear
sbFiles.Panels(2).Text = "Files Receiving"
FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "STRFLS" & sPath

Case 2: 'Download
Case 3: 'Upload

Case 7:
Dim Tp As String
Tp = InputBox("Wijzigen naar de naam je wilt!", , lstFiles.SelectedItem.Text)
FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "RNMFIL" & _
sPath & "\" & "\?/" & lstFiles.SelectedItem.Text & "\?/" & Tp


Case 8: FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "DELDEL" & _
sPath & "\" & lstFiles.SelectedItem.Text
End Select
End Sub

Private Sub MnuCute_Click(Index As Integer)
Select Case Index
Case 0:
FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "SHLEXE" & _
sPath & "\" & lstFiles.SelectedItem.Text & "\?/" & 0

Case 1:
FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "SHLEXE" & _
sPath & "\" & lstFiles.SelectedItem.Text & "\?/" & 1

End Select
End Sub

Private Sub Tvfolders_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Tvfolders.Nodes.Count <> 0 Then If Button = 2 Then PopupMenu Mnufolders
'***
End Sub

Private Sub Mnufol_Click(Index As Integer)
Select Case Index
Case 0: FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "300100"
Case 2: FrmManager.Tvfolders.SelectedItem.Expanded = True
Case 3: FrmManager.Tvfolders.SelectedItem.Expanded = False

Case 5: 'Download folder
Case 6: 'New Folder
'*************************************
Dim Name As String
Dim Person As Node
Name = InputBox("New Folder", "Create folder", "New"): If Name = "" Then Exit Sub
Set Group = Tvfolders.SelectedItem

If CheckenAll(Tvfolders, Tvfolders.SelectedItem.Key, Name) = False Then
Set Person = Tvfolders.Nodes.Add(Tvfolders.SelectedItem.Key, tvwChild, Name & Time, Name, 6)
Person.EnsureVisible
FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "DIRMAP" & sPath & "\" & Name
Else
MsgBox "This folder exist!!", vbInformation
End If

'*******************
Case 7:
FrmMain.Client(FrmMain.LstSIN.SelectedItem.SubItems(1)).SendData "DELMAP" & sPath & "\" & Tvfolders.SelectedItem.Text
Tvfolders.Nodes.Remove (Tvfolders.SelectedItem.Index)

End Select
End Sub


Private Sub Tvfolders_DblClick()
Dim Ssplit() As String

Ssplit = Split(SKeys, Chr(1))
For x = 0 To UBound(Ssplit) - 1
If InStr(Tvfolders.SelectedItem.Key, Ssplit(x)) <> 0 = True Then
Splaats = Ssplit(x)
Sother = Replace(Tvfolders.SelectedItem.Text, Splaats, "")
GoTo Contin
End If
Next x
Contin:

Select Case Tvfolders.SelectedItem.Key
Case Splaats
tEst = True
BrowseFolder Left(Tvfolders.SelectedItem.FullPath, 3)
sbFiles.Panels(1).Text = Splaats

Case Else
tEst = False
BrowseFolder Left(Tvfolders.SelectedItem.FullPath, 3) & Sother

''Tvfolders.Nodes(Index).LastSibling

End Select
End Sub

Private Sub Tvfolders_Click()
''On Error Resume Next
Dim Table As String
Dim Ssplit() As String
Ssplit = Split(Alles, Chr(1))
For x = 0 To UBound(Ssplit) - 1
If InStr(Tvfolders.SelectedItem.FullPath, Ssplit(x)) <> 0 = True Then
Table = Ssplit(x)
End If
Next x
Sother = Replace(Tvfolders.SelectedItem.FullPath, Table, "")
FrmManager.sbFiles.Panels(2).Text = "Files Receiving..."

sPath = Left(Tvfolders.SelectedItem.FullPath, 2) & Sother
sMap = Tvfolders.SelectedItem.Key
End Sub

Private Function CheckenAll(Tree As TreeView, NodeKey As String, Zoek As String) As Boolean
Dim myNode As Node
Dim childCount As Integer
  
Set myNode = Tree.Nodes(NodeKey)
childCount = myNode.Children

Set myNode = myNode.Child.FirstSibling
 
For i = 1 To childCount
If InStr(myNode.Text, Zoek) <> 0 Then
CheckenAll = True
Set myNode = myNode.Next
Else
CheckenAll = False
Set myNode = myNode.Next
End If
Next
''End If
End Function


