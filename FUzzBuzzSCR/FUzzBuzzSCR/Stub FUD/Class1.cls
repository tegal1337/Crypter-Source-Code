VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mFakeClass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'=============================================================================================================
'
' cResource Class Module
' ----------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : January 4, 2002
'
' VB Versions : 5.0 / 6.0
'
' Requires    : Nothing
'
' Description : This class module makes it easy for you to extract all kinds of files and information from
'               resources both external as well as internal to your project.  If you use the "DLL_*"
'               functions, you can extract resources from an ActiveX DLL or standard C-Style DLL.  If you
'               use the "RES_*" functions, you can extract resources from a resource file (.RES) that is
'               included as part of your Visual Basic project.  This gives you the most flexibility and
'               performance to suit your needs.
'
' Example Use :
'
' Option Explicit
' Private Const SND_MEMORY = &H4
' Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal Sound As Long, ByVal lngFlags As Long) As Long
' Private Sub Form_Load()
'   Dim strPath   As String
'   Dim RES       As cResource
'   Dim objTemp   As StdPicture
'   Dim strTemp   As String
'   Dim lngTemp   As Long
'   Dim lngSize   As Long
'   Dim bytTemp() As Byte
'   ' Setup the form to correctly display the BITMAP, CURSOR, & ICON that we are going to give it
'   Me.AutoRedraw = True
'   Me.Visible = True
'   Me.MousePointer = vbCustom
'   ' Get the path to the DLL
'   strPath = App.Path
'   If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
'   strPath = strPath & "Resource.dll"
'   ' Initialize the "cResource" class module
'   Set RES = New cResource
'   RES.LibraryFilePath = strPath
'   ' Load a string from the DLL
'   If RES.DLL_LoadString(101, strTemp) = True Then Me.Caption = strTemp
'   ' Load a bitmap from the DLL
'   If RES.DLL_LoadBitmap(101, objTemp) = True Then Set Me.Picture = objTemp
'   ' Load an icon from the DLL
'   If RES.DLL_LoadIcon(101, objTemp) = True Then Set Me.Icon = objTemp
'   ' Load a cursor from the DLL
'   If RES.DLL_LoadCursor(101, objTemp) = True Then Set Me.MouseIcon = objTemp
'   Me.Refresh
'   ' Load a .WAV file from the DLL and play it
'   If RES.DLL_LoadData(101, lngTemp, lngSize, "WAVE") = True Then sndPlaySound lngTemp, SND_MEMORY
'   RES.SaveData_Pointer "C:\TEST.WAV", lngTemp, lngSize
'   MsgBox "Finished loading information from the resource DLL, now loading information from the current project's resource (.RES) file", vbOKOnly + vbInformation, "  Click OK To Continue"
'   ' Clear the previous information
'   Set objTemp = Nothing
'   strTemp = ""
'   lngTemp = 0
'   Set Me.Picture = Nothing
'   Set Me.Icon = Nothing
'   Set Me.MouseIcon = Nothing
'   Me.Cls
'   Me.Caption = ""
'   MsgBox "Previous data cleared from form", vbOKOnly + vbInformation, "  Click OK To Continue"
'   ' Load a string from the RES
'   If RES.RES_LoadString(101, strTemp) = True Then Me.Caption = strTemp
'   ' Load a bitmap from the RES
'   If RES.RES_LoadBitmap(101, objTemp) = True Then Set Me.Picture = objTemp
'   ' Load an icon from the RES
'   If RES.RES_LoadIcon(101, objTemp) = True Then Set Me.Icon = objTemp
'   ' Load a cursor from the RES
'   If RES.RES_LoadCursor(101, objTemp) = True Then Set Me.MouseIcon = objTemp
'   Me.Refresh
'   ' Load a .WAV file from the DLL and play it
'   If RES.RES_LoadData(101, bytTemp, lngTemp, lngSize, "WAVE") = True Then sndPlaySound lngTemp, SND_MEMORY
'   RES.SaveData_Array "C:\TEST.WAV", bytTemp
'   ' Cleanup
'   Erase bytTemp
'   Set objTemp = Nothing
'   Set RES = Nothing
' End Sub
'
'_____________________________________________________________________________________________________________
' SEE ALSO:
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'
' LoadResource:
' http://msdn.microsoft.com/library/en-us/winui/resource_2i05.asp
'
' Resource Functions:
' http://msdn.microsoft.com/library/en-us/winui/resource_69o3.asp
'
' LoadLibrary:
' http://msdn.microsoft.com/library/en-us/dllproc/dll_1o8p.asp
'
' Dynamic-Link Library Functions:
' http://msdn.microsoft.com/library/en-us/dllproc/dll_0qr7.asp
'
' LoadResData:
' http://msdn.microsoft.com/library/en-us/vb98/html/vbmthloadresdata.asp
'
' LoadResPicture:
' http://msdn.microsoft.com/library/en-us/vb98/html/vbmthloadrespicture.asp
'
' LoadResString:
' http://msdn.microsoft.com/library/en-us/vb98/html/vbmthloadresstring.asp
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


' Type - OleCreatePictureIndirect
Private Type GUID
  Data1    As Long
  Data2    As Integer
  Data3    As Integer
  Data4(7) As Byte
End Type

' Type - OleCreatePictureIndirect / OleLoadPicture
Private Type PICTDESC_BMP 'picType = PICTYPE_BITMAP
  cbSizeOfStruct As Long 'UINT     // Size of the PICTDESC structure.
  PicType        As Long 'UINT     // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
  hBitmap        As Long 'HBITMAP  // The HBITMAP identifying the bitmap assigned to the picture object.
  hPal           As Long 'HPALETTE // The HPALETTE identifying the color palette for the bitmap.
End Type

' Type - OleCreatePictureIndirect / OleLoadPicture
Private Type PICTDESC_ICON 'picType = PICTYPE_ICON
  cbSizeOfStruct As Long 'UINT  // Size of the PICTDESC structure.
  PicType        As Long 'UINT  // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
  hIcon          As Long 'HICON // The HICON identifying the icon assigned to the picture object.
End Type

' Type - OleCreatePictureIndirect / OleLoadPicture
Private Type PICTDESC_META 'picType = PICTYPE_METAFILE
  cbSizeOfStruct As Long 'UINT      // Size of the PICTDESC structure.
  PicType        As Long 'UINT      // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
  hMeta          As Long 'HMETAFILE // The HMETAFILE handle identifying the metafile assigned to the picture object.
  xExt           As Long 'int       // Horizontal extent of the metafile in HIMETRIC units.
  yExt           As Long 'int       // Vertical extent of the metafile in HIMETRIC units.
End Type

' Type - OleCreatePictureIndirect / OleLoadPicture
Private Type PICTDESC_EMETA 'picType = PICTYPE_ENHMETAFILE
  cbSizeOfStruct As Long 'UINT         // Size of the PICTDESC structure.
  PicType        As Long 'UINT         // Type of picture described by this structure, which can be any of the following values: PICTYPE_UNINITIALIZED, PICTYPE_NONE, PICTYPE_BITMAP, PICTYPE_METAFILE, PICTYPE_ICON, PICTYPE_ENHMETAFILE
  hEMF           As Long 'HENHMETAFILE // The HENHMETAFILE identifying the enhanced metafile to assign to the picture object.
End Type

' Type - GetIconInfo.pIconInfo
Private Type ICONINFO
  fIcon    As Long 'BOOL    // Specifies whether this structure defines an icon or a cursor. A value of TRUE (1) specifies an icon; FALSE (0) specifies a cursor
  xHotspot As Long 'DWORD   // Specifies the x-coordinate of a cursor's hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  yHotspot As Long 'DWORD   // Specifies the y-coordinate of the cursor's hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  hbmMask  As Long 'HBITMAP // Specifies the icon bitmask bitmap. If this structure defines a black and white icon, this bitmask is formatted so that the upper half is the icon AND bitmask and the lower half is the icon XOR bitmask. Under this condition, the height should be an even multiple of two. If this structure defines a color icon, this mask only defines the AND bitmask of the icon.
  hbmColor As Long 'HBITMAP // Handle to the icon color bitmap. This member can be optional if this structure defines a black and white icon. The AND bitmask of hbmMask is applied with the SRCAND flag to the destination; subsequently, the color bitmap is applied (using XOR) to the destination by using the SRCINVERT flag.
End Type

' Constants - OleCreateBitmapIndiect (Return Values)
Private Const S_OK = 0                   ' The new picture object was created successfully.
Private Const E_NOINTERFACE = &H80004002 ' The object does not support the interface specified in riid.
Private Const E_POINTER = &H80004003     ' The address in pPictDesc or ppvObj is not valid. For example, it may be NULL.
Private Const E_INVALIDARG = &H80000003  ' One or more arguments are invalid
Private Const E_OUTOFMEMORY = &H8007000E ' Ran out of memory
Private Const E_UNEXPECTED = &H8000FFFF  ' Catastrophic failure

' Constants - PICTDESC.picType
Private Const PICTYPE_UNINITIALIZED = -1 ' The picture object is currently uninitialized.
Private Const PICTYPE_NONE = 0           ' A new picture object is to be created without an initialized state. This value is valid only in the PICTDESC structure.
Private Const PICTYPE_BITMAP = 1         ' The picture type is a bitmap. When this value occurs in the PICTDESC structure, it means that the bmp field of that structure contains the relevant initialization parameters.
Private Const PICTYPE_METAFILE = 2       ' The picture type is a metafile. When this value occurs in the PICTDESC structure, it means that the wmf field of that structure contains the relevant initialization parameters.
Private Const PICTYPE_ICON = 3           ' The picture type is an icon. When this value occurs in the PICTDESC structure, it means that the icon field of that structure contains the relevant initialization parameters.
Private Const PICTYPE_ENHMETAFILE = 4    ' The picture type is a Win32-enhanced metafile. When this value occurs in the PICTDESC structure, it means that the emf field of that structure contains the relevant initialization parameters.

' Constants - GetObjectType.hGdiObject
Private Const OBJ_PEN = 1          ' Pen
Private Const OBJ_BRUSH = 2        ' Brush
Private Const OBJ_DC = 3           ' Device Context
Private Const OBJ_METADC = 4       ' Metafile Device Context
Private Const OBJ_PAL = 5          ' Palette
Private Const OBJ_FONT = 6         ' Font
Private Const OBJ_BITMAP = 7       ' BITMAP
Private Const OBJ_REGION = 8       ' Region
Private Const OBJ_METAFILE = 9     ' MetaFile
Private Const OBJ_MEMDC = 10       ' Memory Device Context
Private Const OBJ_EXTPEN = 11      ' Extended Pen
Private Const OBJ_ENHMETADC = 12   ' Enhanced MetaFile Device Context
Private Const OBJ_ENHMETAFILE = 13 ' Enhanced MetaFile
Private Const OBJ_COLORSPACE = 14  ' Color Space

' Constants - FindResource.lpType
Private Const RT_ACCELERATOR = 9   ' Accelerator table
Private Const RT_ANICURSOR = 21    ' Animated cursor
Private Const RT_ANIICON = 22      ' Animated icon
Private Const RT_BITMAP = 2        ' Bitmap resource
Private Const RT_CURSOR = 1        ' Hardware-dependent cursor resource
Private Const RT_DIALOG = 5        ' Dialog box
Private Const RT_FONT = 8          ' Font resource
Private Const RT_FONTDIR = 7       ' Font directory resource
Private Const RT_GROUP_CURSOR = 12 ' Hardware-independent cursor resource
Private Const RT_GROUP_ICON = 14   ' Hardware-independent icon resource
Private Const RT_HTML = 23         ' HTML document
Private Const RT_ICON = 3          ' Hardware-dependent icon resource
Private Const RT_MENU = 4          ' Menu resource
Private Const RT_MESSAGETABLE = 11 ' Message-table entry
Private Const RT_RCDATA = 10       ' Application-defined resource (raw data)
Private Const RT_STRING = 6        ' String-table entry
Private Const RT_VERSION = 16      ' Version resource

' Win32 Declarations - Resource Related
Private Declare Function FindResource Lib "KERNEL32" Alias "FindResourceA" (ByVal hLib As Long, ByVal strName As String, ByVal strType As String) As Long
Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLib As Long) As Long 'BOOL
Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal strFilePath As String) As Long
Private Declare Function LoadBitmap Lib "USER32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lngBitmapID As Long) As Long
Private Declare Function LoadCursor Lib "USER32" Alias "LoadCursorA" (ByVal hLib As Long, ByVal lngCursorID As Long) As Long
Private Declare Function LoadIcon Lib "USER32" Alias "LoadIconA" (ByVal hLib As Long, ByVal lngIconID As Long) As Long
Private Declare Function LoadString Lib "USER32" Alias "LoadStringA" (ByVal hLib As Long, ByVal ResourceID As Long, ByVal lpBuffer As String, ByVal nBufferSize As Long) As Long
Private Declare Function LoadResource Lib "KERNEL32" (ByVal hLib As Long, ByVal hRes As Long) As Long
Private Declare Function LockResource Lib "KERNEL32" (ByVal hRes As Long) As Long
Private Declare Function SizeofResource Lib "KERNEL32" (ByVal hModule As Long, ByVal hResInfo As Long) As Long

' Win32 Declarations - Misc
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function DeleteColorSpace Lib "gdi32" (ByVal hColorSpace As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEnhMetaFile As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMetaFile As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hGDIObj As Long) As Long
Private Declare Function DestroyCursor Lib "USER32" (ByVal hCursor As Long) As Long
Private Declare Function DestroyIcon Lib "USER32" (ByVal hIcon As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hGdiObject As Long) As Long
Private Declare Function GetIconInfo Lib "USER32" (ByVal IconOrCursor As Long, ByRef pICONINFO As ICONINFO) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef PicDesc As Any, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As StdPicture) As Long

' Local variables
Private hLib        As Long
Private strFilePath As String


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Private Sub Class_Terminate()
  CloseLibrary
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


' Gets or sets the path to the current resource DLL
Public Property Get LibraryFilePath() As String
  LibraryFilePath = strFilePath
End Property
Public Property Let LibraryFilePath(ByVal NewValue As String)
  
  Dim lngErrNum As Long
  
  ' Close the old library
  CloseLibrary
  
  ' Make sure specified file exists
  If Dir(NewValue, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Property
  
  ' Load the new library
  hLib = LoadLibrary(NewValue & Chr(0))
  If hLib <> 0 Then
    strFilePath = NewValue
  Else
    lngErrNum = Err.LastDllError
    Err.Raise lngErrNum, "cResource.LibraryFilePath_Let (" & NewValue & ")", "The LoadLibrary API failed with the error number " & CStr(lngErrNum) & " while trying to load the file '" & NewValue & "'"
  End If
  
End Property

' Returns the handle to the currently opened library
Public Property Get hLibrary() As Long
  hLibrary = hLib
End Property


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'=============================================================================================================
' CloseLibrary
'
' Purpose :
' Closes the currently opened library and cleans up any memory used by it
'
' Param                   Use
' ------------------------------------
' None
'
' Return
' ------
' None
'
'=============================================================================================================
Public Sub CloseLibrary()
  
  If hLib <> 0 Then FreeLibrary hLib
  strFilePath = ""
  hLib = 0
  
End Sub

'=============================================================================================================
' DLL_LoadString
'
' Purpose :
' Loads the specified string from the currently loaded DLL.
'
' WARNING: Do not set the "lngMaxLen" parameter above 1 million unless needed because string memory allocation
'          will eat up memory quickly if you do.  Also, set it lower than the default if you know you don't
'          need that much room.
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the string resource
' Return_String        Returns the string resource specified
' lngMaxLen            Optional. Specifies the maximum length of the string resource.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function DLL_LoadString(ByVal intResourceID As Integer, ByRef Return_String As String, Optional ByVal lngMaxLen As Long = 256) As Boolean
  
  Dim strID   As String
  Dim strTemp As String
  Dim hRes    As Long
  Dim lngLen  As Long
  
  ' Set the default return
  Return_String = ""
  
  ' Validata parameters
  If hLib = 0 Then Exit Function
  If intResourceID < 1 Or intResourceID > 32767 Then Exit Function
  If lngMaxLen < 1 Then Exit Function
  
  ' Get the string
  strTemp = String(lngMaxLen, Chr(0))
  lngLen = LoadString(hLib, intResourceID, strTemp, lngMaxLen)
  If lngLen = 0 Then GoTo CleanUp
  strTemp = Left(strTemp, lngLen)
  
  Return_String = strTemp
  DLL_LoadString = True
  
CleanUp:
  
  strTemp = ""
  
End Function

'=============================================================================================================
' DLL_LoadIcon
'
' Purpose :
' Loads the specified icon from the currently loaded DLL and returns it as an OLE StdPicture object.
'
' IMPORTANT: The caller of this function is responsible for cleaning up the ICON picture object returned
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the icon resource
' Return_Icon          Optional. Returns the resource icon specified as an StdPicture object
' Return_hICON         Optional. Returns the handle to the resource icon specified
' blnReturnHandleOnly  Optional. If TRUE, the "Return_Icon" parameter will not return an StdPicture object
'                      of the icon resource specified.  Setting this to FALSE saves processing and memory.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function DLL_LoadIcon(ByVal intResourceID As Integer, Optional ByRef Return_Icon As StdPicture, Optional ByRef Return_hICON As Long, Optional ByVal blnReturnHandleOnly As Boolean = False) As Boolean
  
  Dim hIco As Long
  
  ' Set the default return
  CleanUpHandle Return_hICON
  Set Return_Icon = Nothing
  
  ' Validata parameters
  If hLib = 0 Then Exit Function
  If intResourceID < 1 Or intResourceID > 32767 Then Exit Function
  
  ' Extract the icon resource
  hIco = LoadIcon(hLib, intResourceID)
  If hIco = 0 Then Exit Function
  
  Return_hICON = hIco
  If blnReturnHandleOnly = True Then
    DLL_LoadIcon = True
  Else
    Set Return_Icon = CreateOlePicture(hIco, vbPicTypeIcon)
    If Not Return_Icon Is Nothing Then
      DLL_LoadIcon = True
    Else
      DestroyIcon hIco
    End If
  End If
  
End Function

'=============================================================================================================
' DLL_LoadCursor
'
' Purpose :
' Loads the specified cursor from the currently loaded DLL and returns it as an OLE StdPicture object.
'
' IMPORTANT: The caller of this function is responsible for cleaning up the CURSOR picture object returned
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the cursor resource
' Return_Cursor        Optional. Returns the resource cursor specified as an StdPicture object
' Return_hCURSOR       Optional. Returns the handle to the resource cursor specified
' blnReturnHandleOnly  Optional. If TRUE, the "Return_Cursor" parameter will not return an StdPicture object
'                      of the cursor resource specified.  Setting this to FALSE saves processing and memory.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function DLL_LoadCursor(ByVal intResourceID As Integer, Optional ByRef Return_Cursor As StdPicture, Optional ByRef Return_hCURSOR As Long, Optional ByVal blnReturnHandleOnly As Boolean = False) As Boolean
  
  Dim hCur As Long
  
  ' Set the default return
  CleanUpHandle Return_hCURSOR
  Set Return_Cursor = Nothing
  
  ' Validata parameters
  If hLib = 0 Then Exit Function
  If intResourceID < 1 Or intResourceID > 32767 Then Exit Function
  
  ' Extract the cursor resource
  hCur = LoadCursor(hLib, intResourceID)
  If hCur = 0 Then Exit Function
  
  Return_hCURSOR = hCur
  If blnReturnHandleOnly = True Then
    DLL_LoadCursor = True
  Else
    Set Return_Cursor = CreateOlePicture(hCur, vbPicTypeIcon)
    If Not Return_Cursor Is Nothing Then
      DLL_LoadCursor = True
    Else
      DestroyCursor hCur
    End If
  End If
  
End Function

'=============================================================================================================
' DLL_LoadBitmap
'
' Purpose :
' Loads the specified bitmap from the currently loaded DLL and returns it as an OLE StdPicture object.
'
' IMPORTANT: The caller of this function is responsible for cleaning up the BITMAP picture object returned
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the bitmap resource
' Return_Bitmap        Optional. Returns the resource bitmap specified as an StdPicture object
' Return_hBITMAP       Optional. Returns the handle to the resource bitmap specified
' blnReturnHandleOnly  Optional. If TRUE, the "Return_Bitmap" parameter will not return an StdPicture object
'                      of the bitmap resource specified.  Setting this to FALSE saves processing and memory.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function DLL_LoadBitmap(ByVal intResourceID As Integer, Optional ByRef Return_Bitmap As StdPicture, Optional ByRef Return_hBITMAP As Long, Optional ByVal blnReturnHandleOnly As Boolean = False) As Boolean
  
  Dim hBmp As Long
  
  ' Set the default return
  CleanUpHandle Return_hBITMAP
  Set Return_Bitmap = Nothing
  
  ' Validata parameters
  If hLib = 0 Then Exit Function
  If intResourceID < 1 Or intResourceID > 32767 Then Exit Function
  
  ' Extract the bitmap resource
  hBmp = LoadBitmap(hLib, intResourceID)
  If hBmp = 0 Then Exit Function
  
  Return_hBITMAP = hBmp
  If blnReturnHandleOnly = True Then
    DLL_LoadBitmap = True
  Else
    Set Return_Bitmap = CreateOlePicture(hBmp, vbPicTypeBitmap)
    If Not Return_Bitmap Is Nothing Then
      DLL_LoadBitmap = True
    Else
      DeleteObject hBmp
    End If
  End If
  
End Function

'=============================================================================================================
' DLL_LoadData
'
' Purpose :
' Loads the specified data file from the currently loaded DLL and returns a pointer to where it is loaded into
' memory.
'
' NOTE: Data files loaded from a resource DLL do *NOT* have to be explicitly cleaned up like Bitmaps, Icons,
'       and Cursors loaded from a resource DLL do.  The reason for this is when the DLL is released by
'       calling the "FreeLibrary" API, these data files are cleaned up automatically.
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the data resource
' Return_DataPointer   Returns the a pointer to where the data file specified has been loaded into memory.
' Return_DataSize      Optional. Returns the size of the data file that is returned.
' strResourceType      Optional. Specifies the "type" of data file that is loaded.  The Visual Basic default
'                      is "CUSTOM".
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function DLL_LoadData(ByVal intResourceID As Integer, ByRef Return_DataPointer As Long, Optional ByRef Return_DataSize As Long, Optional ByVal strResourceType As String = "CUSTOM") As Boolean
  
  Dim strTemp As String
  Dim hRes    As Long
  Dim hData   As Long
  Dim lpData  As Long
  
  ' Set the default return
  Return_DataPointer = 0
  Return_DataSize = 0
  
  ' Validata parameters
  If hLib = 0 Then Exit Function
  If intResourceID < 1 Or intResourceID > 32767 Then Exit Function
  If strResourceType = "" Then Exit Function
  
  ' Find the resource to get
  strTemp = "#" & CStr(intResourceID)
  hRes = FindResource(hLib, strTemp & Chr(0), strResourceType & Chr(0))
  If hRes = 0 Then Exit Function
  
  ' Get a handle to the resource we want
  hData = LoadResource(hLib, hRes)
  If hData = 0 Then Exit Function
  
  ' Lock the resource and get a pointer to the data
  lpData = LockResource(hData)
  If lpData = 0 Then Exit Function
  
  ' Get the size ove the resource in bytes
  Return_DataSize = SizeofResource(hLib, hRes)
  
  Return_DataPointer = lpData
  DLL_LoadData = True
  
End Function

'=============================================================================================================
' RES_LoadString
'
' Purpose :
' Loads the specified string resource from the current project's resource (.RES) file.
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the string resource
' Return_String        Returns the specified string resource
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function RES_LoadString(ByVal intResourceID As Integer, ByRef Return_String As String) As Boolean
On Error Resume Next
  
  Dim lngErrNum As Long
  
  ' Set the default return
  Return_String = ""
  
  ' Load the string from the resource in the current project
  Err.Clear
  Return_String = LoadResString(intResourceID)
  
  ' Check for errors
  lngErrNum = Err.Number
  If lngErrNum = 0 Then RES_LoadString = True
  Err.Clear
  
End Function

'=============================================================================================================
' RES_LoadIcon
'
' Purpose :
' Loads the specified icon resource from the current project's resource (.RES) file.
'
' IMPORTANT: The caller of this function is responsible for cleaning up the ICON picture object returned
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the icon resource
' Return_Icon          Returns the specified icon resource as an OLE StdPicture object.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function RES_LoadIcon(ByVal intResourceID As Integer, ByRef Return_Icon As StdPicture) As Boolean
On Error Resume Next
  
  Dim lngErrNum As Long
  
  ' Set the default return
  Set Return_Icon = Nothing
  
  ' Load the icon from the resource in the current project
  Err.Clear
  Set Return_Icon = LoadResPicture(intResourceID, vbResIcon)
  
  ' Check for errors
  lngErrNum = Err.Number
  If Not Return_Icon Is Nothing And lngErrNum = 0 Then RES_LoadIcon = True
  Err.Clear
  
End Function

'=============================================================================================================
' RES_LoadCursor
'
' Purpose :
' Loads the specified cursor resource from the current project's resource (.RES) file.
'
' IMPORTANT: The caller of this function is responsible for cleaning up the CURSOR picture object returned
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the cursor resource
' Return_Cursor        Returns the specified cursor resource as an OLE StdPicture object.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function RES_LoadCursor(ByVal intResourceID As Integer, ByRef Return_Cursor As StdPicture) As Boolean
On Error Resume Next
  
  Dim lngErrNum As Long
  
  ' Set the default return
  Set Return_Cursor = Nothing
  
  ' Load the cursor from the resource in the current project
  Err.Clear
  Set Return_Cursor = LoadResPicture(intResourceID, vbResCursor)
  
  ' Check for errors
  lngErrNum = Err.Number
  If Not Return_Cursor Is Nothing And lngErrNum = 0 Then RES_LoadCursor = True
  Err.Clear
  
End Function

'=============================================================================================================
' RES_LoadBitmap
'
' Purpose :
' Loads the specified bitmap resource from the current project's resource (.RES) file.
'
' IMPORTANT: The caller of this function is responsible for cleaning up the BITMAP picture object returned
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the bitmap resource
' Return_Bitmap        Returns the specified bitmap resource as an OLE StdPicture object.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function RES_LoadBitmap(ByVal intResourceID As Integer, ByRef Return_Bitmap As StdPicture) As Boolean
On Error Resume Next
  
  Dim lngErrNum As Long
  
  ' Set the default return
  Set Return_Bitmap = Nothing
  
  ' Load the bitmap from the resource in the current project
  Err.Clear
  Set Return_Bitmap = LoadResPicture(intResourceID, vbResBitmap)
  
  ' Check for errors
  lngErrNum = Err.Number
  If Not Return_Bitmap Is Nothing And lngErrNum = 0 Then RES_LoadBitmap = True
  Err.Clear
  
End Function

'=============================================================================================================
' RES_LoadData
'
' Purpose :
' Loads the specified data file resource from the current project's resource (.RES) file.
'
' Param                Use
' ------------------------------------
' intResourceID        Specifies the ID (1 to 32767) of the data file resource
' Return_DataByteArray Returns a BYTE array that represents the data bits of the specified data file.
' Return_DataPointer   Optional. Returns a pointer to the data file returned by the "Return_DataByteArray" parameter
' Return_DataSize      Optional. Returns the size (in bytes) of the specified data file returned.
' strResourceType      Optional. Specifies the "type" of data file being returned.  The Visual Basic default
'                      for data files is "CUSTOM"
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function RES_LoadData(ByVal intResourceID As Integer, ByRef Return_DataByteArray() As Byte, Optional ByRef Return_DataPointer As Long, Optional ByRef Return_DataSize As Long, Optional ByVal strResourceType As String = "CUSTOM") As Boolean
On Error Resume Next
  
  Dim lngErrNum As Long
  
  ' Set the default return
  Erase Return_DataByteArray
  Return_DataPointer = 0
  Return_DataSize = 0
  
  ' Load the custom file from the resource in the current project
  Err.Clear
  Return_DataByteArray = LoadResData(intResourceID, strResourceType)
  
  ' Check for errors
  lngErrNum = Err.Number
  If lngErrNum = 0 Then
    RES_LoadData = True
    Return_DataPointer = VarPtr(Return_DataByteArray(0))
    Return_DataSize = UBound(Return_DataByteArray)
  End If
  Err.Clear
  
End Function

'=============================================================================================================
' SaveData_Pointer
'
' Purpose :
' This function takes the specified data pointer and saves out the data file it points to based on the size
' of the data file specified.
'
' Param                Use
' ------------------------------------
' strSavePath          Specifies the full path of the file to save out to.
' DataPointer          Specifies a pointer to the spot in memory where the data file to be saved resides.
' DataSize             Specifies the size of the data file that the "DataPointer" parameter points to.
' blnOverwriteIfExists Optional. If set to TRUE and the file specified in the "strSavePath" parameter already
'                      exists, the existing file will be overwritten with the new one.  If set to FALSE, the
'                      existing file is left alone and the specified data file is not written out.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function SaveData_Pointer(ByVal strSavePath As String, ByVal DataPointer As Long, ByVal DataSize As Long, Optional ByVal blnOverwriteIfExists As Boolean = True) As Boolean
  
  Dim FileNum    As Integer
  Dim bytArray() As Byte
  
  ' Validate parameters
  strSavePath = Trim(strSavePath)
  If strSavePath = "" Then Exit Function
  If DataPointer = 0 Then Exit Function
  If DataSize < 1 Then Exit Function
  
  ' Check if the file exists
  If FileExists(strSavePath) = True Then
    If blnOverwriteIfExists = False Then
      SaveData_Pointer = True
      Exit Function
    Else
      Kill strSavePath
    End If
  End If
  
  ' Take the data and copy it to a byte array to utilize the "Put" function
  ReDim bytArray(DataSize) As Byte
  CopyMemory bytArray(0), ByVal DataPointer, DataSize
  
  ' Save the information to file
  FileNum = FreeFile
  Open strSavePath For Binary As #FileNum
    Put #FileNum, 1, bytArray()
  Close #FileNum
  
  ' Saved successfully
  SaveData_Pointer = True
  
  ' Cleanup
  Erase bytArray
  
End Function

'=============================================================================================================
' SaveData_Array
'
' Purpose :
' This function takes the specified data file in the form of a BYTE array and saves it out to the specified
' file.
'
' Param                Use
' ------------------------------------
' strSavePath          Specifies the full path of the file to save out to.
' DataArray            Specifies the BYTE array that represents the data file to save out.
' blnOverwriteIfExists Optional. If set to TRUE and the file specified in the "strSavePath" parameter already
'                      exists, the existing file will be overwritten with the new one.  If set to FALSE, the
'                      existing file is left alone and the specified data file is not written out.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function SaveData_Array(ByVal strSavePath As String, ByRef DataArray() As Byte, Optional ByVal blnOverwriteIfExists As Boolean = True) As Boolean
  
  Dim FileNum As Integer
  
  ' Validate parameters
  strSavePath = Trim(strSavePath)
  If strSavePath = "" Then Exit Function
  
  ' Check if the file exists
  If FileExists(strSavePath) = True Then
    If blnOverwriteIfExists = False Then
      SaveData_Array = True
      Exit Function
    Else
      Kill strSavePath
    End If
  End If
  
  ' Save the information to file
  FileNum = FreeFile
  Open strSavePath For Binary As #FileNum
    Put #FileNum, 1, DataArray()
  Close #FileNum
  
  ' Saved successfully
  SaveData_Array = True
  
End Function


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


'=============================================================================================================
'
' CreateOlePicture
'
' This function takes the handle to a picture (Bitmap, Icon, Metafile, or Enhanced Metafile) and creates
' an OLE StdPicture object from it that can be used like the "Picture" properties of such VB objects as
' Form's, PictureBox's, ImageBox's, etc.
'
' Parameter:              Use:
' --------------------------------------------------
' PictureHandle           Handle to the picture to create.
'                           - If PictureType = vbPicTypeBitmap    : this must be a handle to a HBITMAP
'                           - If PictureType = vbPicTypeIcon      : this must be a handle to a HICON
'                           - If PictureType = vbPicTypeMetafile  : this must be a handle to a HMETAFILE
'                           - If PictureType = vbPicTypeEMetafile : this must be a handle to a HENHMETAFILE
' PictureType             Specifies the type of picture object to create.  These are the different types
'                         of pictures that can be specified:
'                           vbPicTypeBitmap     <-- DEFAULT
'                           vbPicTypeEMetafile
'                           vbPicTypeIcon
'                           vbPicTypeMetafile
' BitmapPalette           Optional. Specifies the handle to a Palette to use in the createion process.
' MetaHeight              Optional. If the PictureType is vbPicTypeMetafile, the height of the Metafile
'                         must be provided by this parameter.
' Metawidth               Optional. If the PictureType is vbPicTypeMetafile, the width of the Metafile
'                         must be provided by this parameter.
' Return_ErrNum           Optional. If an error occurs, the error number will be returned here.
' Return_ErrDesc          Optional. If an error occurs, the error description will be returned here.
'
' Return:
' -------
' If the function succeeds, the return is TRUE
' If the function fails, the return is FALSE
'
'=============================================================================================================
Private Function CreateOlePicture(ByVal PictureHandle As Long, ByVal PictureType As PictureTypeConstants, Optional ByVal BitmapPalette As Long = 0, Optional ByVal MetaHeight As Long = -1, Optional ByVal MetaWidth As Long = -1, Optional ByRef Return_ErrNum As Long, Optional ByRef Return_ErrDesc As String) As StdPicture
On Error Resume Next
  
  Dim ReturnValue   As Long
  Dim PicInfo_BMP   As PICTDESC_BMP
  Dim PicInfo_ICON  As PICTDESC_ICON
  Dim PicInfo_EMETA As PICTDESC_EMETA
  Dim PicInfo_META  As PICTDESC_META
  Dim ThePicture    As StdPicture 'IPicture
  Dim rIID          As GUID
  
  ' Clear the return variables
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Make sure the variable(s) passed are valid
  If PictureHandle = 0 Then
    Return_ErrNum = -1
    Return_ErrDesc = "Invalid bitmap handle"
  ElseIf PictureType = vbPicTypeNone Then
    Return_ErrNum = -1
    Return_ErrDesc = "Invalid picture type specified."
  ElseIf PictureType = vbPicTypeMetafile Then
    If MetaHeight = -1 Or MetaWidth = -1 Then
      Return_ErrNum = -1
      Return_ErrDesc = "Invalid metafile dimentions specified."
    End If
  End If
  
  ' Set the correct interface identifier GUID for the "OleCreatePictureIndirect" API
  With rIID
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
  End With
  
  ' Set the appropriate type depending on the type of picture
  Select Case PictureType
    Case vbPicTypeBitmap
      PicInfo_BMP.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_BMP.PicType = PICTYPE_BITMAP
      PicInfo_BMP.hBitmap = PictureHandle
      PicInfo_BMP.hPal = BitmapPalette
      ReturnValue = OleCreatePictureIndirect(PicInfo_BMP, rIID, 1, ThePicture)
      
    Case vbPicTypeIcon
      PicInfo_ICON.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_ICON.PicType = PICTYPE_ICON
      PicInfo_ICON.hIcon = PictureHandle

      ReturnValue = OleCreatePictureIndirect(PicInfo_ICON, rIID, 1, ThePicture)
      
    Case vbPicTypeMetafile
      PicInfo_META.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_META.PicType = PICTYPE_METAFILE
      PicInfo_META.hMeta = PictureHandle
      PicInfo_META.xExt = MetaWidth
      PicInfo_META.yExt = MetaHeight
      ReturnValue = OleCreatePictureIndirect(PicInfo_META, rIID, 1, ThePicture)
      
    Case vbPicTypeEMetafile
      PicInfo_EMETA.cbSizeOfStruct = Len(PicInfo_BMP)
      PicInfo_EMETA.PicType = PICTYPE_ENHMETAFILE
      PicInfo_EMETA.hEMF = PictureHandle
      ReturnValue = OleCreatePictureIndirect(PicInfo_BMP, rIID, 1, ThePicture)
  End Select
  
  ' Check the result
  If ReturnValue <> S_OK Then GoTo ErrorTrap
  
  ' Return the new picture
  Set CreateOlePicture = ThePicture
  
  Exit Function
  
ErrorTrap:
  
  Return_ErrNum = ReturnValue
  Select Case ReturnValue
    Case E_NOINTERFACE
      Return_ErrDesc = "The object does not support the interface specified in riid."
    Case E_POINTER
      Return_ErrDesc = "The address in pPictDesc or ppvObj is not valid. For example, it may be NULL."
    Case E_INVALIDARG
      Return_ErrDesc = "One or more arguments are invalid."
    Case E_OUTOFMEMORY
      Return_ErrDesc = "Ran out of memory."
    Case E_UNEXPECTED
      Return_ErrDesc = "Catastrophic Failure."
    Case Else
      Return_ErrDesc = "Unknown Error."
  End Select
  
End Function

' Attempt to figure out what type of object is being passed and delete it if it can find out what it is
Private Sub CleanUpHandle(ByRef hGdiObject As Long)
  
  Dim lngType         As Long
  Dim IconInformation As ICONINFO
  
  If hGdiObject = 0 Then Exit Sub
  
  ' If the object passed is an icon or cursor, destroy it
  If GetIconInfo(hGdiObject, IconInformation) <> 0 Then
    If IconInformation.fIcon = 1 Then
      DestroyIcon hGdiObject
      hGdiObject = 0
      Exit Sub
    Else
      DestroyCursor hGdiObject
      hGdiObject = 0
      Exit Sub
    End If
  End If
  
  ' If we can't determin the type of object passed, simply exit
  If GetObjectType(hGdiObject) = 0 Then
    hGdiObject = 0
    Exit Sub
  End If
  
  ' Go through the returned value and clean up the object appropriately
  Select Case lngType
    Case OBJ_PEN:         DeleteObject hGdiObject
    Case OBJ_EXTPEN:      DeleteObject hGdiObject
    Case OBJ_BRUSH:       DeleteObject hGdiObject
    Case OBJ_PAL:         DeleteObject hGdiObject
    Case OBJ_FONT:        DeleteObject hGdiObject
    Case OBJ_BITMAP:      DeleteObject hGdiObject
    Case OBJ_REGION:      DeleteObject hGdiObject
    Case OBJ_MEMDC:       DeleteDC hGdiObject
    Case OBJ_METAFILE:    DeleteMetaFile hGdiObject
    Case OBJ_ENHMETAFILE: DeleteEnhMetaFile hGdiObject
    Case OBJ_COLORSPACE:  DeleteColorSpace hGdiObject
  End Select
  hGdiObject = 0
  
End Sub

' Function that checks if a file exists
Private Function FileExists(ByVal FilePath As String) As Boolean
  
  If Dir(FilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
    FileExists = True
  Else
    FileExists = False
  End If
  
End Function

