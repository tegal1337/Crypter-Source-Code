Attribute VB_Name = "Modfunct"
Public Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function StrFormatByteSizeA Lib "shlwapi" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Fso As Object
Public tEst As Boolean
Public Splaats As String
Public SKeys As String
Public Alles As String
Public Const Splitter = "\?/"

Public Function FormatKB(ByVal Amount As Long) As String
    Dim Buffer As String
    Dim Result As String
    Buffer = Space$(255)
    Result = StrFormatByteSizeA(Amount, Buffer, Len(Buffer))
    If InStr(Result, vbNullChar) > 1 Then FormatKB = Left$(Result, InStr(Result, vbNullChar) - 1)
End Function

Public Function NodeExists(tv As TreeView, ByVal sKey As String) As Boolean
   Dim nd As Node
   On Error Resume Next
   Set nd = tv.Nodes(sKey)
   NodeExists = (Err = 0)
   Set nd = Nothing
End Function

Public Function Drivers(sState As String, Stype As String) As String
Select Case sState
            Case 2:
            If Left(Stype, 1) = "A" Then
            Drivers = "(Floppy"
            Else:            Drivers = "(Removable"
            End If
            Case 3: Drivers = "(Fixed"
            Case 4: Drivers = "(Remote"
            Case 5: Drivers = "(CD-Rom"
            Case 6: Drivers = "(RAM"
End Select
End Function

Public Function DriveIcon(Str As String) As Long
Select Case Str
        Case "(Removable": DriveIcon = 2
        Case "(Fixed": DriveIcon = 3
        Case "(Remote": DriveIcon = 4
        Case "(CD-Rom": DriveIcon = 5
        Case "(RAM": DriveIcon = 7
        Case "(Floppy": DriveIcon = 1
    End Select
End Function
