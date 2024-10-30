Attribute VB_Name = "iconchanger"
Option Explicit

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (init As InitCommonControlsExType) As Boolean

Private Declare Function ActivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hwnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
Private Declare Function DeactivateWindowTheme Lib "uxtheme.dll" Alias "SetWindowTheme" (ByVal hwnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
Private Declare Function IsThemeActiveXP Lib "uxtheme.dll" Alias "IsThemeActive" () As Boolean
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (hTheme As Long) As Long
Private Declare Function EnableThemeDialogTexture Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, Optional hrgnUpdate As Long, Optional fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Boolean

Private Const ETDT_DISABLE      As Long = &H1
Private Const ETDT_ENABLE       As Long = &H2

Private Const RDW_UPDATENOW     As Long = &H100

Private Const ICC_USEREX_CLASSES As Long = &H200

Private Const ECM_FIRST         As Long = &H1500
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)

Private m_bIsManifestActive     As Boolean
Private bIsVbRunning            As Boolean
Type DIB_HEADER
   Size        As Long
   Width       As Long
   Height      As Long
   Planes      As Integer
   Bitcount    As Integer
   Reserved    As Long
   ImageSize   As Long
End Type

Type ICON_DIR_ENTRY
   bWidth            As Byte
   bHeight           As Byte
   bColorCount       As Byte
   bReserved         As Byte
   wPlanes           As Integer
   wBitCount         As Integer
   dwBytesInRes      As Long
   dwImageOffset     As Long
End Type

Type ICON_DIR
   Reserved          As Integer
   Type              As Integer
   Count             As Integer
End Type

Type DIB_BITS
   Bits()            As Byte
End Type

Public Enum Errors
   FILE_CREATE_FAILED = 1000
   FILE_READ_FAILED
   INVALID_PE_SIGNATURE
   INVALID_ICO
   NO_RESOURCE_TREE
   NO_ICON_BRANCH
   CANT_HACK_HEADERS
End Enum

Private Type InitCommonControlsExType
    dwSize  As Long     'size of this structure
    dwICC   As Long     'flags indicating which classes to be initialized
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type BALLOONTIP
    cbStruct As Long
    pszTitle As String
    pszText As String
    tIcon As Long
End Type

Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

Enum BalloonTipIconConstants
    balNone = 0
    balExcalmation = 1
    balInformation = 2
    balCritical = 3
End Enum

Public Function ReplaceIcons(Source As String, Dest As String, Error As String) As Long
   On Error Resume Next
   Dim IcoDir As ICON_DIR
   Dim IcoDirEntry As ICON_DIR_ENTRY
   Dim tBits As DIB_BITS
   Dim Icons() As IconDescriptor
   Dim lngRet As Long
   Dim BytesRead As Long
   Dim hSource As Long
   Dim hDest As Long
   Dim ResTree As Long
    
   hSource = CreateFile(Source, ByVal &H80000000, 0, ByVal 0&, 3, 0, ByVal 0)
   If hSource >= 0 Then
      If Valid_ICO(hSource) Then
         SetFilePointer hSource, 0, 0, 0
         ReadFile hSource, IcoDir, 6, BytesRead, ByVal 0&
         ReadFile hSource, IcoDirEntry, 16, BytesRead, ByVal 0&
         SetFilePointer hSource, IcoDirEntry.dwImageOffset, 0, 0
         ReDim tBits.Bits(IcoDirEntry.dwBytesInRes) As Byte
         ReadFile hSource, tBits.Bits(0), IcoDirEntry.dwBytesInRes, BytesRead, ByVal 0&
         CloseHandle hSource
         hDest = CreateFile(Dest, ByVal (&H80000000 Or &H40000000), 0, ByVal 0&, 3, 0, ByVal 0)
         If hDest >= 0 Then
            If Valid_PE(hDest) Then
               ResTree = GetResTreeOffset(hDest)
               If ResTree > 308 Then   ' Sanity check
                  lngRet = GetIconOffsets(hDest, ResTree, Icons)
                  SetFilePointer hDest, Icons(1).Offset, 0, 0
                  WriteFile hDest, tBits.Bits(0), UBound(tBits.Bits), BytesRead, ByVal 0&
                  If Not HackDirectories(hDest, ResTree, Icons(1).Offset, IcoDirEntry) Then
                     err.Raise CANT_HACK_HEADERS, App.EXEName, "Unable to modify directories in target executable.  File may not contain any icon resources."
                  End If
               Else
                  err.Raise NO_RESOURCE_TREE, App.EXEName, Dest & " does not contain a valid resource tree.  File may be corrupt."
                  CloseHandle hDest
               End If
            Else
               err.Raise INVALID_PE_SIGNATURE, App.EXEName, Dest & " is not a valid Win32 executable."
               CloseHandle hDest
            End If
         CloseHandle hDest
         Else
            err.Raise FILE_CREATE_FAILED, App.EXEName, "Failed to open " & Dest & ". Make sure file is not in use by another program."
         End If
      Else
         err.Raise INVALID_ICO, App.EXEName, Source & " is not a valid icon resource file."
         CloseHandle hSource
      End If
   Else
      err.Raise FILE_CREATE_FAILED, App.EXEName, "Failed to open " & Source & ". Make sure file is not in use by another program."
   End If
   ReplaceIcons = 0
   Exit Function
ErrHandler:
   ReplaceIcons = err.Number
   Error = err.Description
End Function
Public Function Valid_ICO(hFile As Long) As Boolean
   Dim tDir          As ICON_DIR
   Dim BytesRead     As Long
   If (hFile > 0) Then
      ReadFile hFile, tDir, Len(tDir), BytesRead, ByVal 0&
      If (tDir.Reserved = 0) And (tDir.Type = 1) And (tDir.Count > 0) Then
         Valid_ICO = True
      Else
         Valid_ICO = False
      End If
   Else
      Valid_ICO = False
   End If
End Function

Private Function InitCommonControls() As Boolean
    Dim InitCC As InitCommonControlsExType
    
    With InitCC
        .dwSize = Len(InitCC)
        .dwICC = ICC_USEREX_CLASSES
    End With
    
    InitCommonControls = InitCommonControlsEx(InitCC)         'initialize the common controls
End Function


Private Function CheckVB() As Boolean
    bIsVbRunning = True
    CheckVB = True
End Function


Private Function GetWindowTheme(hwnd As Long, Optional PartID As String) As Long
    'this will retrive the current hTheme used by the window..
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    Dim hTheme As Long
    If PartID = "" Then PartID = "Window"
    hTheme = OpenThemeData(hwnd, StrPtr(PartID))
    CloseThemeData hTheme
    GetWindowTheme = hTheme
    
End Function

Private Function GetWinVersion() As String
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function

Private Function AddDirSep(Path As String) As String
    Dim DirSep As String
    
    If PathIsNetworkPath(Path) = True Then
        DirSep = "/"
    Else
        DirSep = "\"
    End If
    
    If Right(Trim(Path), Len(DirSep)) <> DirSep Then
        AddDirSep = Trim(Path) & DirSep
    Else
        AddDirSep = Path
    End If
    
End Function


Function HideTextBalloonTip(Control As Control) As Boolean
    
    Dim hwnd As Long
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    Select Case UCase(TypeName(Control))
        
        Case "TEXTBOX"
            hwnd = Control.hwnd
        Case "RICHTEXTBOX"
            hwnd = Control.hwnd
        Case "COMBOBOX"
            If (Control.Style = 0 Or 1) Then
                Dim Cbo As COMBOBOXINFO
                Cbo.cbSize = Len(Cbo)
                Call GetComboBoxInfo(Control.hwnd, Cbo)
                hwnd = Cbo.hwndEdit
            Else
                Exit Function
            End If
        Case Else
            hwnd = Control.hwnd
    End Select
    
    HideTextBalloonTip = SendMessage(hwnd, EM_HIDEBALLOONTIP, 0&, 0&)

End Function

Function IsThemingSupported() As Boolean

    Dim hLib As Long                    'module handle..
    hLib = LoadLibrary("uxtheme.dll")   'retrive the module handle.
    Call FreeLibrary(hLib)              'unload the dll
    IsThemingSupported = CBool(hLib)    'if the return value = 0 then
                                        'the dll does not exist,
                                        'otherwise, the dll is there..
End Function





Function IsXPThemed(hwnd As Long) As Boolean
    
    'check if the object is using a visual style..
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    
    Dim hTheme As Long
        
    hTheme = OpenThemeData(hwnd, StrPtr("Window"))  'do the theme test
    
    Call CloseThemeData(hTheme)                     'close the theme data handle
    
    IsXPThemed = CBool(hTheme)                      'if zero, return False.. else return true..
    
    
End Function


Public Function ShowTextBalloonTip(Control As Control, Prompt As String, Optional Title As String, Optional TitleIcon As BalloonTipIconConstants) As Boolean
    
    'This function will show an EDIT balloon tip..
    'this function will only apply to a normal text box, a richtext box or a combobox
    'with syle 0 or 1...
    'any other controls passed to this function will return false (as i know!)
    
    Dim Bal As BALLOONTIP
    Dim hwnd As Long
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function

    Select Case UCase(TypeName(Control))
        Case "COMBOBOX"
            If (Control.Style = 0 Or Control.Style = 1) Then
                Dim Cbo As COMBOBOXINFO
                Cbo.cbSize = Len(Cbo)
                Call GetComboBoxInfo(Control.hwnd, Cbo)
                hwnd = Cbo.hwndEdit
            Else
                Exit Function
            End If
        Case "TEXTBOX"
            hwnd = Control.hwnd
        Case "RICHTEXTBOX"
            hwnd = Control.hwnd
        Case Else
            hwnd = Control.hwnd
    End Select
    
    With Bal
        .cbStruct = Len(Bal)                    'set the structure size
        .pszTitle = StrConv(Title, vbUnicode)   'convert the title into unicode format..
        .pszText = StrConv(Prompt, vbUnicode)   'convert the prompt into unicode format..
        .tIcon = TitleIcon                      'set the title icon
    End With
    
    'show the balloon tip..
    
    ShowTextBalloonTip = SendMessage(hwnd, EM_SHOWBALLOONTIP, 0&, Bal)
    
    
End Function

Function ToggleVisualStyles(Frm As Form, Enable As Boolean, Optional ToggleFormBorder As Boolean = True)
    
    'Enable/diable a form theming ..

    On Error GoTo ErrorHandler
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    
    Dim fControls   As String   'This is the forbidden controls (controls with no .hWnd/cannot be skinned)
                                'i didn't use an array cause i found out that it's too slow..
    
    Dim Ctl         As Control
    Dim sTypeName   As String
           
    For Each Ctl In Frm.Controls                            'loop through the controls collection of the form..
        sTypeName = UCase(TypeName(Ctl))                    'Get the typename of the control.
        If InStr(1, fControls, sTypeName) = 0 Then          'look for the control type name in the forbidden controls list, if found, do nothing..
            Select Case Enable                              'activate/deactivate theming
                Case True:  Call EnableXPLook(Ctl)
                Case False: Call DisableXPLook(Ctl)
            End Select
            If TypeName(Ctl) = "PICTUREBOX" Then Ctl.Refresh    'refresh any pictureboxes in the form..
        End If
    Next
    
    If ToggleFormBorder = True Then
        Select Case Enable                                      'activate/deactivate the form theming..
            Case True
                Call EnableXPLook(Frm)
            Case False
                Call DisableXPLook(Frm): Call DisableXPDlgBackground(Frm)
        End Select
    End If
    
    Set Ctl = Nothing       'erase the ctl variable from memory..

    Frm.Refresh             'refresh the form
    
   'Debug.Print fControls
    Exit Function

ErrorHandler:                                   'This is the error handling section...

    If err.Number = 438 Then                    'object doesn't have a ".hWnd" property..
        'Debug.Print "Error: The Object '" & Ctl.Name & "' doesn't have a '.hwnd' property.."
        fControls = sTypeName & "," & fControls 'add this typename into the forbidden list..
        Resume Next                             'skip the line where the error happened, and proceed to the next line..
    Else                                        'unexpected error..
        err.Raise err.Number                    'show the error..
    End If
End Function


Function EnableXPLook(ByRef Object As Object) As Boolean
    'this function will draw the object using windows xp visual styles..
    'note: the object MUST have a handle
    
    On Error GoTo ErrHandler:

    Dim wRECT   As RECT
    
    GetWindowRect Object.hwnd, wRECT   'retrive the object region.
        
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    ActivateWindowTheme (Object.hwnd) 'try to enable theming
    
    If IsXPThemed(Object.hwnd) <> 0 Then
        'ok
        EnableXPLook = True
    Else
        'error
        GoTo ErrHandler
    End If
    
    Call RedrawWindow(Object.hwnd, wRECT, , RDW_UPDATENOW) 'refresh the object
   
    Exit Function
ErrHandler:
    EnableXPLook = False
    Exit Function
End Function

Function DisableXPLook(ByRef Object As Object) As Boolean
    'this function will disable the object's visual style..
    'note: the object MUST have a handle
    'same as the EnableXPLook function..
    
    Dim wRECT As RECT
    
    On Error GoTo ErrHandler:
    
    GetWindowRect Object.hwnd, wRECT
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    DeactivateWindowTheme (Object.hwnd)
    
    If IsXPThemed(Object.hwnd) = False Then
        DisableXPLook = True
    Else
        GoTo ErrHandler
    End If
    
    Call RedrawWindow(Object.hwnd, wRECT, , RDW_UPDATENOW)
    
    Exit Function
    
ErrHandler:
    DisableXPLook = False
    Exit Function
End Function

Function DrawTabBackground(oPictureBox As Object, Optional sTab As Object)
    
    On Error Resume Next
    'Draw a TabStrip control's background texture in a picture box..
    'this is a good example on how to draw controls using "uxtheme.dll" API calls..
    
    Dim hTheme          As Long         'The theme handle
    Dim dRECT           As RECT         'The drawing Region
    Dim tabHwnd         As Long
    Const TAB_BODY      As Integer = 10 'this is the PartID of the tabstrip background..
    
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Function
    If IsThemeActiveXP() = False Then Exit Function
    
    tabHwnd = sTab.hwnd
    
    If tabHwnd <> 0 Then
        If IsXPThemed(sTab.hwnd) = False Then oPictureBox.Cls: Exit Function  'if the frame theming is disabled, Clear the picture box and exit
    End If
    
    oPictureBox.Cls
    oPictureBox.AutoRedraw = False
    
    'copy the picturebox measurements into the RECT object
    
    dRECT.Left = 0
    dRECT.Top = 0
    dRECT.Right = oPictureBox.ScaleX(oPictureBox.Width, oPictureBox.ScaleMode, vbPixels)
    dRECT.Bottom = oPictureBox.ScaleY(oPictureBox.Height, oPictureBox.ScaleMode, vbPixels)

    hTheme = OpenThemeData(oPictureBox.hwnd, StrPtr("TAB"))      'Retrive the handle of the current theme being used.
    
    If hTheme <> 0 Then
        Call DrawThemeBackground(hTheme, oPictureBox.hDC, TAB_BODY, 0, dRECT, dRECT) 'draw the tab background on the picture box
    Else
        oPictureBox.Cls
    End If
    
    oPictureBox.AutoRedraw = True
    
    CloseThemeData hTheme           'close the theme data handle..
    
End Function


Sub EnableXPDlgBackground(Frm As Form)
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Sub
    If IsThemeActiveXP() = False Then Exit Sub
    Call EnableThemeDialogTexture(Frm.hwnd, ETDT_ENABLE)
End Sub

Sub DisableXPDlgBackground(Form As Form)
    If IsWindowsXP() = False Or IsThemingSupported() = False Or IsVBRunning = True Then Exit Sub
    If IsThemeActiveXP() = False Then Exit Sub
    Call EnableThemeDialogTexture(Form.hwnd, ETDT_DISABLE)
End Sub

Public Function IsVBRunning() As Boolean
    Debug.Assert (CheckVB) = True
    IsVBRunning = bIsVbRunning
    bIsVbRunning = False
End Function


Private Function IsWindowsXP() As Boolean
If Val(Trim(GetWinVersion)) >= 5.01 Then
IsWindowsXP = True
End If
End Function

Private Function vb5Replace(Expression As String, Find As String, ReplaceWith As String, Optional start As Long = 1, Optional Count As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
On Error GoTo ExitProcedure
Dim iFind As Long
Dim nextStart As Long
Dim sCount As Long
iFind = InStr(start, Expression, Find, Compare)
nextStart = start
If iFind = -1 Then
vb5Replace = Expression
Exit Function
Else
Do
If sCount >= Count And Count <> -1 Then
Exit Do
End If
iFind = InStr(nextStart, Expression, Find, Compare)
If iFind = 0 Then Exit Do
sCount = sCount + 1
Expression = Left(Expression, iFind - 1) & ReplaceWith & Mid(Expression, iFind + Len(Find))
If nextStart < Len(Expression) Then
nextStart = iFind + Len(ReplaceWith) + 1
Else
Exit Do
End If
Loop
End If
ExitProcedure:
vb5Replace = Expression
End Function

Function XPStyle(Optional AutoRestart As Boolean = True, Optional Autohide As Boolean = True, Optional CreateNew As Boolean = False) As Boolean
If IsWindowsXP = False Or IsVBRunning Or IsThemingSupported = False Then Exit Function
If IsThemeActiveXP = False Then Exit Function
Const IsVB6 As Boolean = True
On Error Resume Next
Dim XML             As String
Dim ManifestCheck   As String
Dim strManifest     As String
Dim FreeFileNo      As Integer
If AutoRestart = True Or Autohide = True Then CreateNew = False
XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & "<assembly " & vbCrLf & "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & "   manifestVersion=""1.0"">" & vbCrLf & "<assemblyIdentity " & vbCrLf & "    processorArchitecture=""x86"" " & vbCrLf & "    version=""EXEVERSION""" & vbCrLf & "    type=""win32""" & vbCrLf & "    name=""COMPANYNAME.EXENAME""/>" & vbCrLf & "    <description>EXEDESCRIPTION</description>" & vbCrLf & "    <dependency>" & vbCrLf & "    <dependentAssembly>" & vbCrLf & "    <assemblyIdentity" & vbCrLf & "         type=""win32""" & vbCrLf & "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & "         version=""6.0.0.0""" & vbCrLf & "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & "         language=""*""" & vbCrLf & "         processorArchitecture=""x86""/>" & vbCrLf & "    </dependentAssembly>" & vbCrLf & "    </dependency>" & vbCrLf & "</assembly>" & vbCrLf & "")
strManifest = AddDirSep(App.Path) & App.EXEName & ".exe.manifest"
ManifestCheck = Dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)
If ManifestCheck = "" Or CreateNew = True Then
If IsVB6 = True Then
XML = Replace(XML, "EXENAME", App.EXEName & ".exe")
XML = Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0")
XML = Replace(XML, "EXEDESCRIPTION", App.FileDescription)
XML = Replace(XML, "COMPANYNAME", App.CompanyName)
Else
XML = vb5Replace(XML, "EXENAME", App.EXEName & ".exe")
XML = vb5Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0")
XML = vb5Replace(XML, "EXEDESCRIPTION", App.FileDescription)
XML = vb5Replace(XML, "COMPANYNAME", App.CompanyName)
End If
FreeFileNo = FreeFile
If ManifestCheck <> "" Then
SetAttr strManifest, vbNormal
Kill (strManifest)
End If
Open strManifest For Binary As #(FreeFileNo)
Put #(FreeFileNo), , XML
Close #(FreeFileNo)
SetAttr strManifest, vbHidden + vbSystem
XPStyle = False
If AutoRestart = True Then
Shell App.Path & "\" & App.EXEName & ".exe" & _
Space(1) & Command$, vbNormalFocus
End
End If
Else
If Autohide = True Then
SetAttr strManifest, vbNormal
Kill (strManifest)
End If
XPStyle = True
End If
m_bIsManifestActive = XPStyle
End Function

Public Property Get IsManifestActive() As Boolean
IsManifestActive = m_bIsManifestActive
End Property






