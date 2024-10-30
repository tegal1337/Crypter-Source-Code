VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOptions 
         Caption         =   "File Options"
         Begin VB.Menu mnuAdd 
            Caption         =   "Add File"
         End
         Begin VB.Menu mnuEdit 
            Caption         =   "Edit File"
         End
         Begin VB.Menu mnuRemove 
            Caption         =   "Remove File"
         End
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuCrypt 
      Caption         =   "Crypt"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, ByVal lpstrFile As Long, ByVal nMaxFile As Long, ByVal lpstrInitialDir As Long, ByVal lpstrDefExt As Long, ByVal lpstrFilter As Long, ByVal lpstrTitle As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Sub mnuAdd_Click()
Call AddFile
End Sub
Private Sub mnuAbout_Click()
Form3.Show
End Sub
Private Sub mnuSettings_Click()
Form2.Show
End Sub
Private Sub mnuCrypt_Click()
Call Form1.WriteCryptedFile
End Sub
Private Sub mnuRemove_Click()
On Error Resume Next
Form1.ListView1.ListItems.Remove (Form1.ListView1.SelectedItem.Index)
Unload Form5
End Sub

Private Sub mnuEdit_Click()
If Form1.ListView1.ListItems.Count = 0 Then
MsgBox "Nothing to edit...", vbCritical
Unload Form5
Exit Sub
End If
AreEditing = 1
Form4.Show
Unload Form5
End Sub
Private Function AddFile()
            Form4.Show
            AreEditing = 0
            Unload Form5
End Function
