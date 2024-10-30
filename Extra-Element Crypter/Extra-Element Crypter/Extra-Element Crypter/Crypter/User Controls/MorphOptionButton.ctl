VERSION 5.00
Begin VB.UserControl MorphOptionCheck 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
   ToolboxBitmap   =   "MorphOptionButton.ctx":0000
End
Attribute VB_Name = "MorphOptionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'* MorphOptionCheck v2.20 - Ownerdrawn OptionButton/Checkbox UserControl *
'* Written July 6, 2005, Matthew R. Usner of Planet Source Code.         *
'* Modified April, 2006 MRU for new subclassing and general improvements *
'* Copyright ©2006, Matthew R. Usner, All Rights Reserved.               *
'*************************************************************************
'* This usercontrol replaces the intrinsic VB OptionButton and CheckBox  *
'* controls.  Many aspects of the control's graphics are modifiable.     *
'* Control incorporates Version 2.1 of Paul Caton's usercontrol sub-     *
'* classing to achieve mouseover color highlight effects, mouse enter /  *
'* leave, and proper focus handling.  Select icons for checkmarks if     *
'* desired.  Unicode support is also incorporated.  Control may be rend- *
'* ered transparent.  If you use this feature,  set the .Transparent     *
'* property to True, and also set the .ContainerName property to the     *
'* name of the container (the name of the frame, etc) that the control   *
'* is in.  Keep this property blank (null string) if control is placed   *
'* directly on the form.                                                 *
'*************************************************************************
'* Legal:  Redistribution of this code, whole or in part, as source code *
'* or in binary form, alone or as part of a larger distribution or prod- *
'* uct, is forbidden for any commercial or for-profit use without the    *
'* author's explicit written permission.                                 *
'*                                                                       *
'* Redistribution of this code, as source code or in binary form, with   *
'* or without modification, is permitted provided that the following     *
'* conditions are met:                                                   *
'*                                                                       *
'* Redistributions of source code must include this list of conditions,  *
'* and the following acknowledgment:                                     *
'*                                                                       *
'* This code was developed by Matthew R. Usner.                          *
'* Source code, written in Visual Basic, is freely available for non-    *
'* commercial, non-profit use.                                           *
'*                                                                       *
'* Redistributions in binary form, as part of a larger project, must     *
'* include the above acknowledgment in the end-user documentation.       *
'* Alternatively, the above acknowledgment may appear in the software    *
'* itself, if and where such third-party acknowledgments normally appear.*
'*************************************************************************
'* Credits and Thanks to:                                                *
'* Paul Caton, for the self-subclassing usercontrol code (Version II).   *
'* Carles P.V., for the gradient paint routine.                          *
'* Richard Mewett, for the Unicode support.                              *
'*************************************************************************

Option Explicit

Private Enum TRACKMOUSEEVENT_FLAGS
   TME_HOVER = &H1&
   TME_LEAVE = &H2&
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
   cbSize                             As Long
   dwFlags                            As TRACKMOUSEEVENT_FLAGS
   hwndTrack                          As Long
   dwHoverTime                        As Long
End Type

Private bTrack                        As Boolean
Private bTrackUser32                  As Boolean
Private bInCtrl                       As Boolean
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'========================= subclasser declarations ===============================
' windows messages to be intercepted by subclassing.
Private Const WM_MOUSEMOVE            As Long = &H200
Private Const WM_MOUSELEAVE           As Long = &H2A3
Private Const WM_SETFOCUS             As Long = &H7
Private Const WM_KILLFOCUS            As Long = &H8
Private Const WM_CLOSE                As Long = &H10
Private Const WM_DESTROY              As Long = &H2

Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H34                                  'WndProc execution offset
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data storage index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data storage index of the subclassed window's hWnd
Private Const IDX_WNDPROC   As Long = 8                                     'Thunk data storage index of the original WndProc
Private Const IDX_BTABLE    As Long = 10                                    'Thunk data storage index of the Before table
Private Const IDX_ATABLE    As Long = 11                                    'Thunk data storage index of the After table
Private Const IDX_PARM_USER As Long = 12                                    'Thunk data storage index of the User-defined callback parameter data index

Private z_Base              As Long                                         'Data pointer base
Private z_TblEnd            As Long                                         'End of the owner object's vTable
Private z_SC(61)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'==================================================================================================

'  declares for Unicode support.
Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
   dwOSVersionInfoSize                As Long
   dwMajorVersion                     As Long
   dwMinorVersion                     As Long
   dwBuildNumber                      As Long
   dwPlatformId                       As Long
   szCSDVersion                       As String * 128        '  Maintenance string for PSS usage
End Type
Private mWindowsNT                    As Boolean
Private Const DT_CALCRECT             As Long = &H400        ' if used, DrawText API just calculates rectangle.
Private Const DT_SINGLELINE           As Long = &H20         ' strip cr/lf from string before draw.
Private Const DT_NOPREFIX             As Long = &H800        ' ignore access key ampersand.
Private Const DT_LEFT                 As Long = &H0          ' draw from left edge of rectangle.
Private Const DT_NOCLIP               As Long = &H100        ' ignores right edge of rectangle when drawing.

'  graphics api declares.
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' transparency declares
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, ByRef lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long

'  declares for Carles P.V.'s gradient paint routine.
Private Type BITMAPINFOHEADER
   biSize                             As Long
   biWidth                            As Long
   biHeight                           As Long
   biPlanes                           As Integer
   biBitCount                         As Integer
   biCompression                      As Long
   biSizeImage                        As Long
   biXPelsPerMeter                    As Long
   biYPelsPerMeter                    As Long
   biClrUsed                          As Long
   biClrImportant                     As Long
End Type
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Const DIB_RGB_COLORS          As Long = 0
Private Const PI                      As Single = 3.14159265358979
Private Const TO_DEG                  As Single = 180 / PI
Private Const TO_RAD                  As Single = PI / 180
Private Const INT_ROT                 As Long = 1000

'  enum for determining if the control is a checkbox or option button.
Public Enum ControlTypeOptions
   [CheckBox]                                                ' control is a checkbox.
   [OptionButton]                                            ' control is an option button.
End Enum

Public Enum CBAlignmentOptions
   [Align Left]                                              ' checkbox is displayed in left side of control.
   [Align Right]                                             ' checkbox is displayed in right side of control.
End Enum

Public Enum MouseOverOptions
   [None]                                                    ' no graphics changes on mouseover.
   [Border]                                                  ' border color is changed on mouseover.
   [CheckBox Border]                                         ' checkbox border color changed on mouseover.
   [Both]                                                    ' both of the above are changed on mouseover.
End Enum

'  for use in determining if mouse is in control.
Private Type POINTAPI
   x As Long                                                 ' horizontal pixel position.
   y As Long                                                 ' vertical pixel position.
End Type
Private MousePos As POINTAPI

'  rectangle structure for API drawing of text onto control.
Private Type RECT
   Left                               As Long
   Top                                As Long
   Right                              As Long
   Bottom                             As Long
End Type

' property default constants.
Private Const m_def_BackAngle = 90                           ' horizontal gradient.
Private Const m_def_BackColor1 = &H4040&                     ' gold gradient background.
Private Const m_def_BackColor2 = &HC0FFFF                    ' gold gradient background.
Private Const m_def_BackMiddleOut = True                     ' background gradient is middle-out.
Private Const m_def_BorderColor = 0                          ' black border.
Private Const m_def_BorderCurvature = 0                      ' no border curvature.
Private Const m_def_BorderIfTransparent = True               ' border is drawn when control is transparent.
Private Const m_def_BorderWidth = 1                          ' 1-pixel border width.
Private Const m_def_Caption = "Option_Check"                 ' caption.
Private Const m_def_CaptionColor = 0                         ' black caption color.
Private Const m_def_CheckBorderColor = 0                     ' black checkbox border
Private Const m_def_CheckBoxAlignment = 0                    ' checkbox aligned to left.
Private Const m_def_CheckBoxAngle = 90                       ' horizontal checkbox gradient.
Private Const m_def_CheckBoxColor1 = 0                       ' black checkbox background.
Private Const m_def_CheckBoxColor2 = 0                       ' black checkbox background.
Private Const m_def_CheckBoxMiddleOut = True                 ' checkbox gradient is middle-out.
Private Const m_def_CheckColor = &HFFFF&                     ' yellow checkmark.
Private Const m_def_ContainerName = ""                       ' default: control is on form.
Private Const m_def_ControlType = 0                          ' checkbox control by default.
Private Const m_def_DisBackColor1 = &H404040                 ' a shade of grey.
Private Const m_def_DisBackColor2 = &HC0C0C0                 ' a shade of grey.
Private Const m_def_DisBorderColor = &H404040                ' a shade of grey.
Private Const m_def_DisCaptionColor = &H808080               ' a shade of grey.
Private Const m_def_DisCheckBorderColor = &H808080           ' a shade of grey.
Private Const m_def_DisCheckBoxColor1 = &HB0B0B0             ' a shade of grey.
Private Const m_def_DisCheckBoxColor2 = &HB0B0B0             ' a shade of grey.
Private Const m_def_Enabled = True                           ' control is enabled.
Private Const m_def_FocusRectColor = &H0                     ' black custom focus rectangle.
Private Const m_def_MouseOverActions = 0                     ' no recoloring when mouseover.
Private Const m_def_MOverBorderColor = 0                     ' black control border when mouseover.
Private Const m_def_MOverCheckBoxColor = 0                   ' black checkbox border when mouseover.
Private Const m_def_PicForCheck = False                      ' use checkmarks instead of pictures.
Private Const m_def_ShowFocusRect = True                     ' display focus rectangle.
Private Const m_def_Transparent = False                      ' control opaque, not transparent.
Private Const m_def_Value = 0                                ' option not selected/checked.

'  these variables allow the control to switch between enabled/disabled appearances.
Private ActiveBackColor1              As OLE_COLOR           ' current first background gradient color.
Private ActiveBackColor2              As OLE_COLOR           ' current second background gradient color.
Private ActiveBorderColor             As OLE_COLOR           ' current control border color.
Private ActiveCaptionColor            As OLE_COLOR           ' current caption text color.
Private ActiveCheckBorderColor        As OLE_COLOR           ' current checkbox frame color.
Private ActiveCheckBoxColor1          As OLE_COLOR           ' current first checkbox gradient color.
Private ActiveCheckBoxColor2          As OLE_COLOR           ' current second checkbox gradient color.

' property variables.
Private m_BackAngle                   As Single              ' angle of control background gradient.
Private m_BackColor1                  As OLE_COLOR           ' the first color of the control's background gradient.
Private m_BackColor2                  As OLE_COLOR           ' the second color of the control's background gradient.
Private m_BackMiddleOut               As Boolean             ' if True, control background gradient is middle-out.
Private m_BorderColor                 As OLE_COLOR           ' the color of the control's border.
Private m_BorderCurvature             As Integer             ' amount of curvature the control's corners are to have.
Private m_BorderIfTransparent         As Boolean             ' if True, border is drawn in Transparent mode.
Private m_BorderWidth                 As Long                ' the width of the control's border.
Private m_Caption                     As String              ' the text to display alongside the checkbox.
Private m_CaptionColor                As OLE_COLOR           ' the color of the caption text.
Private m_CheckBorderColor            As OLE_COLOR           ' the color of the checkbox border.
Private m_CheckBoxAlignment           As CBAlignmentOptions  ' left or right checkbox alignment enum.
Private m_CheckBoxAngle               As Single              ' angle of checkbox background gradient.
Private m_CheckBoxColor1              As OLE_COLOR           ' the first color of the checkbox gradient.
Private m_CheckBoxColor2              As OLE_COLOR           ' the second color of the checkbox gradient.
Private m_CheckBoxMiddleOut           As Boolean             ' if True, checkbox background gradient is middle-out.
Private m_CheckColor                  As OLE_COLOR           ' the color of the checkmark.
Private m_ContainerName               As String              ' container name control is in; blank if on form.
Private m_ControlType                 As ControlTypeOptions  ' allows selection of the type of control to display.
Private m_DisBackColor1               As OLE_COLOR           ' disabled background gradient color 1.
Private m_DisBackColor2               As OLE_COLOR           ' disabled background gradient color 2.
Private m_DisBorderColor              As OLE_COLOR           ' disabled border color.
Private m_DisCaptionColor             As OLE_COLOR           ' disabled caption color.
Private m_DisCheckBorderColor         As OLE_COLOR           ' disabled checkbox border color.
Private m_DisCheckBoxColor1           As OLE_COLOR           ' disabled checkbox gradient color 1.
Private m_DisCheckBoxColor2           As OLE_COLOR           ' disabled checkbox gradient color 2.
Private m_Enabled                     As Boolean             ' control's enabled status.
Private m_FocusRectColor              As OLE_COLOR           ' the color of the custom focus rectangle.
Private m_Font                        As Font                ' the font used to draw the caption.
Private m_MouseOverActions            As MouseOverOptions    ' what gets colored when mouse is over control?
Private m_MOverBorderColor            As OLE_COLOR           ' control border color when mouse is over control.
Private m_MOverCheckBoxColor          As OLE_COLOR           ' checkbox border color when mouse is over control.
Private m_PicChecked                  As Picture             ' picture to display when control is checked.
Private m_PicForCheck                 As Boolean             ' if True, pics display instead of tickmarks.
Private m_PicUnchecked                As Picture             ' picture to display when control is unchecked.
Private m_ShowFocusRect               As Boolean             ' display focus rectangle flag.
Private m_Transparent                 As Boolean             ' if True, only checkbox and caption are visible.
Private m_Value                       As Integer             ' the selected status of the optionbutton/checkbox.

' event declarations.
Public Event MouseEnter()
Public Event MouseLeave()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Click()

Private HasFocus As Boolean                                  ' indicates if MorphOptionCheck has focus.
Private CheckBoxX As Long                                    ' x coordinate of left edge of checkbox.
Private KeyIsDown As Boolean                                 ' indicates if a key is being pressed.
Private MouseIsDown As Boolean                               ' indicates if the left mouse mutton is down.
Private SaveBorderColor As Long                              ' original control border color.
Private SaveCheckBoxBorderColor  As Long                     ' original checkbox border color.
Private MaxPicWidth As Long                                  ' width in pixels of widest check icon.

'-- Added By Gary Noble (Phantom Man)
Private Const ClientSideNotAvailable  As Long = 398          ' Client Side Not Available VB error code.

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

'*************************************************************************
'* track the mouse leaving the indicated window.                         *
'*************************************************************************

   Dim tme As TRACKMOUSEEVENT_STRUCT

   If bTrack Then
      With tme
         .cbSize = Len(tme)
         .dwFlags = TME_LEAVE
         .hwndTrack = lng_hWnd
       End With
       If bTrackUser32 Then
         Call TrackMouseEvent(tme)
       Else
         Call TrackMouseEventComCtl(tme)
       End If
   End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<< Event-Handling Routines >>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_Initialize()

'*************************************************************************
'* the first event in the control's life cycle.                          *
'*************************************************************************

   Dim OS As OSVERSIONINFO

   OS.dwOSVersionInfoSize = Len(OS)
   Call GetVersionEx(OS)
   mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_Show()

'*************************************************************************
'* sets max icon width, gets enabled status  and redraws control.        *
'*************************************************************************

'  determine width of widest check icon.
   MaxPicWidth = PicWidth

'  load appropriate display properties based on enabled state of control.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

   RedrawControl

End Sub

Private Sub UserControl_Resize()

'*************************************************************************
'* used only for resizing/redisplay in design mode.                      *
'*************************************************************************

'  determine the x coordinate of the left edge of the checkbox.
   GetCheckBoxX
   RedrawControl

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* not used in the control, but passed along to the rest of the project. *
'*************************************************************************

   If m_Enabled Then
      KeyIsDown = True
      RaiseEvent KeyDown(KeyCode, Shift)
   End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* allows the space bar to "click" the control like vb equivalents.      *
'*************************************************************************

   KeyIsDown = False

'  "click" control if the control is enabled, key pressed is space bar and mouse is not down.
   If m_Enabled And KeyCode = 32 And Not MouseIsDown Then
      If m_ControlType = OptionButton And Not m_Value Then
'        if control is in OptionButton mode and it hasn't already been selected,
'        select it and deselect all other option buttons in container or form.
         ProcessMorphOptionButtons
      Else
         If m_ControlType = CheckBox Then
'           if control is in CheckBox mode, reverse its selection status.
            m_Value = IIf(m_Value = vbChecked, vbUnchecked, vbChecked)
            RedrawControl
         End If
      End If
'     pass along the KeyUp and KeyPress events to project regardless of key pressed.
      RaiseEvent KeyUp(KeyCode, Shift)
      RaiseEvent KeyPress(KeyCode)
   End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* not used in the control, but passed along to the rest of the project. *
'*************************************************************************

   If m_Enabled Then
      MouseIsDown = True
      RaiseEvent MouseDown(Button, Shift, x, y)
   End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* sets MorphOptionButton's value to True (only on left button click).   *
'*************************************************************************

   MouseIsDown = False

'  if mouse has left control, ignore.  This allows you to "back out" of a click, so
'  to speak, by holding down mouse button, dragging the mouse out, and releasing.
   GetCursorPos MousePos
   If WindowFromPoint(MousePos.x, MousePos.y) <> UserControl.hwnd Then
      Exit Sub   ' don't send event.
   End If

'  only bother if the control is enabled, left mouse button is clicked, and no key is pressed.
   If m_Enabled And Button = vbLeftButton And Not KeyIsDown Then
      If m_ControlType = OptionButton Then
         If Not m_Value Then
'           if control is in OptionButton mode and it hasn't already been selected,
'           select it and deselect all other option buttons in container or form.
            ProcessMorphOptionButtons
         End If
      Else
'        if control is in CheckBox mode, reverse its selection status.
         m_Value = IIf(m_Value = vbChecked, vbUnchecked, vbChecked)
         RedrawControl
      End If
'     pass along both the MouseUp and Click events.
      RaiseEvent MouseUp(Button, Shift, x, y)
      RaiseEvent Click
   End If

End Sub

Private Sub UserControl_Terminate()
   On Error GoTo Catch
   sc_Terminate    ' terminate subclassing.
Catch:
End Sub

Private Sub ProcessMorphOptionButtons()

'*************************************************************************
'* sets MorphOptionButton's value to True (only on left button click).   *
'* Then loops through all like controls in the container or form that    *
'* this MorphOptionButton is in and sets them to False if not already so.*
'*************************************************************************

   Dim ctl As Control

'  set the MorphOptionButton's value to True and redraw it.
   m_Value = True
   RedrawControl

'  loop through each control in the MorphOptionButton's container or form.
   For Each ctl In UserControl.Parent.Controls
'     if the control we're looking at is a MorphOptionCheck...
      If TypeOf ctl Is MorphOptionCheck Then
'        and the MorphOptionCheck is in the same container/form...
         If ctl.Container.hwnd = UserControl.ContainerHwnd Then
'           and the MorphOptionCheck is not THIS MorphOptionCheck...
            If ctl.hwnd <> UserControl.hwnd Then
'              and the MorphOptionCheck is set to True...
               If ctl.Value = True Then
                  ctl.Value = False
               End If
            End If
         End If
      End If
   Next

End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics Routines  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub RedrawControl()

'*************************************************************************
'* the master routine for displaying textbox and its contents.           *
'*************************************************************************

'  if the control is not in Transparent mode, display the gradient background and border.
   If Not m_Transparent Then
      SetBackGround
      SetBorder
   Else
      MakeControlTransparent
      If m_BorderIfTransparent Then
         SetBorder
      End If
   End If
   SetCheckBox        ' display the checkbox background/border/checkmark.
   SetText            ' display the caption.

End Sub

Private Sub MakeControlTransparent()

'*************************************************************************
'* displays what's behind the control, thereby effectively rendering the *
'* control transparent.  NOTE:  The control's container MUST have the    *
'* .hDC property exposed.  Original code by Dana Seaman and  Matthew R.  *
'* Usner, with mods by Gary Noble (Phantom Man) and LaVolpe.             *
'*************************************************************************

   On Error GoTo ErrNoClientSide

   Dim pX           As Long
   Dim pY           As Long
   Dim ctl          As Control
   Dim h1           As Long
   Dim BorderOffset As Long
   Dim bAPIway      As Boolean
   Dim wRect        As RECT
   Dim wPT          As POINTAPI

   If UserControl.BorderStyle Then
      If UserControl.Appearance Then
         BorderOffset = 2
      Else
         BorderOffset = 1
      End If
   End If

   If LenB(m_ContainerName) = 0 Then
'     the container resides on a Form, so use parent hDC.
        If UserControl.Parent.hwnd = UserControl.ContainerHwnd Then
            UserControl.Parent.AutoRedraw = True
            h1 = UserControl.Parent.hDC
        Else
            bAPIway = True
        End If
   Else
'     the container resides in another container.
      For Each ctl In Parent.Controls
'        find container.
         If UCase(ctl.Name) = UCase(m_ContainerName) Then 'Found our container
            ctl.AutoRedraw = True 'AutoRedraw must be True
            h1 = ctl.hDC 'Get the container's hDC
            Exit For
         End If
      Next
   End If

'  get offsets for BitBlt.
   If Not bAPIway Then
      If UserControl.Extender.Container.ScaleMode = vbTwips Then
         pX = UserControl.Extender.Left \ Screen.TwipsPerPixelX
         pY = UserControl.Extender.Top \ Screen.TwipsPerPixelY
      Else
         pX = UserControl.Extender.Left
         pY = UserControl.Extender.Top
      End If
   Else
      GetWindowRect UserControl.hwnd, wRect
      wPT.x = wRect.Left
      wPT.y = wRect.Top
      ScreenToClient UserControl.ContainerHwnd, wPT
      pX = wPT.x
      pY = wPT.y
      OffsetRect wRect, -wRect.Left + wPT.x, -wRect.Top + wPT.y
      ShowWindow UserControl.hwnd, 0
      RedrawWindow UserControl.ContainerHwnd, wRect, ByVal 0&, &H1
      DoEvents
      h1 = GetDC(UserControl.ContainerHwnd)
   End If
'  copy background to usercontrol DC.
   BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, h1, pX + BorderOffset, pY + BorderOffset, vbSrcCopy

   If bAPIway Then
        ReleaseDC h1, UserControl.ContainerHwnd
        ShowWindow UserControl.hwnd, 1
   End If

'  -- following code added by Gary Noble (Phantom Man)
   Exit Sub

ErrNoClientSide:

'  -- This Will Take Care Of The Form Closing Event
'  -- Before, If You Closed The Parent The Control Was Still Trying
'  -- To Get The DC Of The Parent That Did Not Exist, If It Doesn't Exist You Can 't BitBlt The BackGround
   If Err.Number <> ClientSideNotAvailable Then
      Resume Next
   Else
'     -- No Client Side - Bail!
      Exit Sub
   End If

End Sub

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient.                           *
'*************************************************************************

   PaintGradient hDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(ActiveBackColor1), _
                 TranslateColor(ActiveBackColor2), m_BackAngle, m_BackMiddleOut

End Sub

Private Sub SetCheckBox()

'*************************************************************************
'* controls the display of control's checkbox or (un)selected picture.   *
'*************************************************************************

   If Not m_PicForCheck Then
      DisplayCheckBox
   Else
      DisplayPicture
   End If

End Sub

Private Sub DisplayCheckBox()

'*************************************************************************
'* displays checkbox border, gradient and appropriate checkmark.         *
'*************************************************************************

'  display the checkbox's background gradient.
   PaintGradient hDC, CheckBoxX, (ScaleHeight - 15) / 2, 15, 15, TranslateColor(ActiveCheckBoxColor1), _
                 TranslateColor(ActiveCheckBoxColor2), m_CheckBoxAngle, m_CheckBoxMiddleOut

'  draw a one-pixel border around the checkbox.
   DrawRectangle CheckBoxX, (ScaleHeight - 15) / 2, CheckBoxX + 15, (ScaleHeight - 15) / 2 + 15, ActiveCheckBorderColor

'  display the checkmark (or icon), if necessary.
   If m_Value And m_ControlType = OptionButton Then
      SetOptionButtonCheckMark
   Else
      If m_Value = vbChecked Then
         SetCheckBoxCheckMark
      End If
   End If

End Sub

Private Sub DisplayPicture()

'*************************************************************************
'* displays picture according to control status (checked or unchecked).  *
'*************************************************************************

   Select Case m_ControlType
      Case OptionButton
         If m_Value Then
            DisplayPic m_PicChecked
         Else
            DisplayPic m_PicUnchecked
         End If
      Case CheckBox
         If m_Value = vbChecked Then
            DisplayPic m_PicChecked
         Else
            DisplayPic m_PicUnchecked
         End If
   End Select

End Sub

Private Sub DisplayPic(Pic As StdPicture)

'*************************************************************************
'* displays appropriate icon in lieu of drawn checkbox.                  *
'*************************************************************************

   If IsThere(Pic) Then
      PaintPicture Pic, CheckBoxX, (UserControl.ScaleHeight / 2) - Int((ScaleX(Pic.Height, vbHimetric, vbPixels)) / 2)
   End If

End Sub

Private Function IsThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks to see if a picture exists by checking its dimensions.         *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub SetOptionButtonCheckMark()

'*************************************************************************
'* draws the 'selected' arrow in OptionButton mode.                      *
'*************************************************************************

   Dim hPO           As Long    ' selected pen object.
   Dim hPN           As Long    ' pen object for drawing checkmark.
   Dim R             As Long    ' loop and result variable for api calls.
   Dim X1            As Long    ' the x coordinate of the start of the checkmark.
   Dim Y1            As Long    ' the y coordinate of the start of the checkmark vertical line.
   Dim Y2            As Long    ' the y coordinate of the end of the checkmark vertical line.
   Dim DrawDirection As Long    ' draw from left to right or right to left?

'  determine x and y coordinates of first part of check arrow to draw and the direction to draw.
   If m_CheckBoxAlignment = [Align Left] Then
      X1 = CheckBoxX + 5
      DrawDirection = 1
   Else
      X1 = CheckBoxX + 9
      DrawDirection = -1
   End If
   Y1 = ScaleHeight \ 2 - 5
   Y2 = 11

'  draw the check arrow.
   hPN = CreatePen(0, 1, m_CheckColor)
   hPO = SelectObject(hDC, hPN)
   MoveTo hDC, X1, Y1, ByVal 0&
   For R = 1 To 6
      LineTo hDC, X1, Y1 + Y2
      X1 = X1 + DrawDirection
      Y1 = Y1 + 1
      Y2 = Y2 - 2
      MoveTo hDC, X1, Y1, ByVal 0&
   Next R

'  delete the pen object.
   R = SelectObject(hDC, hPO)
   R = DeleteObject(hPN)

End Sub

Private Sub SetCheckBoxCheckMark()

'*************************************************************************
'* draws a tick mark in the checkbox in CheckBox mode.                   *
'*************************************************************************

   Dim I     As Long    ' loop variable.
   Dim x     As Long    ' the x coordinate of current pixel being drawn.
   Dim y As Long        ' the y coordinate of current pixel being drawn.

   x = CheckBoxX + 1
   y = (ScaleHeight \ 2) - 7

   For I = 9 To 12: SetPixelV hDC, x + I, y + 3, m_CheckColor: Next I
   For I = 8 To 11: SetPixelV hDC, x + I, y + 4, m_CheckColor: Next I
   For I = 7 To 10: SetPixelV hDC, x + I, y + 5, m_CheckColor: Next I
   For I = 1 To 2: SetPixelV hDC, x + I, y + 6, m_CheckColor: Next I
   For I = 6 To 9: SetPixelV hDC, x + I, y + 6, m_CheckColor: Next I
   For I = 1 To 3: SetPixelV hDC, x + I, y + 7, m_CheckColor: Next I
   For I = 5 To 8: SetPixelV hDC, x + I, y + 7, m_CheckColor: Next I
   For I = 1 To 7: SetPixelV hDC, x + I, y + 8, m_CheckColor: Next I
   For I = 2 To 6: SetPixelV hDC, x + I, y + 9, m_CheckColor: Next I
   For I = 3 To 5: SetPixelV hDC, x + I, y + 10, m_CheckColor: Next I
   SetPixelV hDC, x + 4, y + 11, m_CheckColor

End Sub

Private Sub SetBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvature.     *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim hBrush As Long    ' the brush pattern used to 'paint' the border.
   Dim hRgn1  As Long    ' the outer boundary of the border region.
   Dim hRgn2  As Long    ' the inner boundary of the border region.

'  create and combine the outer and inner border regions.
   hRgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, m_BorderCurvature, m_BorderCurvature)
   hRgn2 = CreateRoundRectRgn(m_BorderWidth, m_BorderWidth, ScaleWidth - m_BorderWidth, _
                              ScaleHeight - m_BorderWidth, m_BorderCurvature, m_BorderCurvature)
   CombineRgn hRgn2, hRgn1, hRgn2, 3

'  create the color brush and fill the combined region.
   hBrush = CreateSolidBrush(TranslateColor(ActiveBorderColor))
   FillRgn hDC, hRgn2, hBrush

'  set the control region to match the created region.
   SetWindowRgn hwnd, hRgn1, True

'  free the memory.
   DeleteObject hRgn2
   DeleteObject hBrush
   DeleteObject hRgn1

ErrHandler:
   Exit Sub

End Sub

Private Sub GetCheckBoxX()

'*************************************************************************
'* gets X coordinate of left edge of checkbox/pic based on alignment.    *
'*************************************************************************

   If Not m_PicForCheck Then
      If m_CheckBoxAlignment = [Align Left] Then
         CheckBoxX = m_BorderWidth + 4
      Else
         CheckBoxX = ScaleWidth - m_BorderWidth - 19
      End If
   Else
      MaxPicWidth = PicWidth
      If m_CheckBoxAlignment = [Align Left] Then
         CheckBoxX = m_BorderWidth + 4
      Else
         CheckBoxX = ScaleWidth - m_BorderWidth - MaxPicWidth - 1
      End If
   End If

End Sub

Private Function PicWidth() As Long

'*************************************************************************
'* gets width of largest checkbox icon for text positioning purposes.    *
'*************************************************************************

   Dim Wid1 As Long    ' width of 'checked' icon (if it exists).
   Dim Wid2 As Long    ' width of 'unchecked' icon (if it exists).

'  obtain the widths of the 'checked' and 'unchecked' icons (if they exist).
   If IsThere(m_PicChecked) Then
      Wid1 = m_PicChecked.Width
   End If
   If IsThere(m_PicUnchecked) Then
      Wid2 = m_PicUnchecked.Width
   End If

'  determine the larger width of the two icons.
   If Wid2 > Wid1 Then
      Wid1 = Wid2
   End If

'  return the largest width in pixels.
   PicWidth = ScaleX(Wid1, vbHimetric, vbPixels)

End Function

Private Sub SetText()

'*************************************************************************
'* displays the caption text.  Selected text is displayed using the      *
'* SelTextColor value.                                                   *
'*************************************************************************

   If Not m_Font Is Nothing Then

      Dim R           As RECT    ' the rectangle that defines the text draw area.
      Dim tHeight     As Long    ' the height of the text.
      Dim tWidth      As Long    ' the width of the text.
      Dim Clearance   As Long    ' to right- or left-justify text.

'     get the height and width of the text based on the selected font.
      tHeight = TextHeight(m_Caption)
      tWidth = TextWidthU(hDC, m_Caption)

'     determine left edge of text based on checkbox/icon mode, and checkbox/icon alignment.
      If Not m_PicForCheck Then
         If m_CheckBoxAlignment = [Align Left] Then
            Clearance = TextWidthU(hDC, "n") + 20
         Else
            Clearance = TextWidthU(hDC, "n")
         End If
      Else
'        right- or left-justify the text in the control based on icon alignment.
         If m_CheckBoxAlignment = [Align Left] Then
            Clearance = ScaleWidth - tWidth - m_BorderWidth - TextWidthU(hDC, "i")
         Else
            Clearance = m_BorderWidth + TextWidthU(hDC, "i")
         End If
      End If

'     set the text color.
      UserControl.ForeColor = TranslateColor(ActiveCaptionColor)

'     define the text drawing area rectangle size.
      With R
         .Left = Clearance
         .Top = (ScaleHeight - tHeight) / 2
         .Bottom = R.Top + tHeight
         .Right = .Left + tWidth
      End With

'     display the text using DrawText API.
      DrawText UserControl.hDC, m_Caption, -1, R, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP

'     if the MorphOptionCheck has focus and the .ShowFocusRect property is True, draw
'     a focus rectangle around the text, adding a one-pixel clearance around the text.
      If HasFocus And m_ShowFocusRect Then
         With R
            DrawRectangle .Left - 1, .Top - 1, .Right + 1, .Bottom + 1, m_FocusRectColor
         End With
      End If

   End If

End Sub

Private Sub DrawRectangle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long)

'*************************************************************************
'* draws the checkbox and focus rectangles.                              *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim hBrush As Long    ' the brush pattern used to 'paint' the border.
   Dim hRgn1  As Long    ' the outer boundary of the rectangle region.
   Dim hRgn2  As Long    ' the inner boundary of the rectangle region.

'  create the outer region.
   hRgn1 = CreateRoundRectRgn(X1, Y1, X2, Y2, 0, 0)
'  create the inner region.
   hRgn2 = CreateRoundRectRgn(X1 + 1, Y1 + 1, X2 - 1, Y2 - 1, 0, 0)
   
'  combine the regions into one border region.
   CombineRgn hRgn2, hRgn1, hRgn2, 3

'  create and apply the color brush.
   hBrush = CreateSolidBrush(TranslateColor(lcolor))
   FillRgn hDC, hRgn2, hBrush

'  free the memory.
   DeleteObject hRgn2
   DeleteObject hBrush
   DeleteObject hRgn1

ErrHandler:
   Exit Sub

End Sub

Private Function TextWidthU(ByVal hDC As Long, sString As String) As Long

'*************************************************************************
'* a better alternative to the VB method .TextWidth.  Thanks LaVolpe!    *
'*************************************************************************

   Dim TextRect As RECT

   SetRect TextRect, 0, 0, 0, 0
   DrawText hDC, sString, -1, TextRect, DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_LEFT
   TextWidthU = TextRect.Right + 1

End Function

Private Sub DrawText(ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)

'*************************************************************************
'* draws the text with Unicode support based on OS version.              *
'* Thanks to Richard Mewett.                                             *
'*************************************************************************

   If mWindowsNT Then
      DrawTextW hDC, StrPtr(lpString), nCount, lpRect, wFormat
   Else
      DrawTextA hDC, lpString, nCount, lpRect, wFormat
   End If

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* translates ole color into COLORREF long for drawing purposes.         *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub PaintGradient(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, _
                          ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, _
                          ByVal Angle As Single, ByVal bMOut As Boolean)

'*************************************************************************
'* Carles P.V.'s routine, modified by Matthew R. Usner for middle-out    *
'* gradient capability.  Original submission at PSC, txtCodeID=60580.    *
'*************************************************************************

   Dim uBIH      As BITMAPINFOHEADER
   Dim lBits()   As Long
   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, b1 As Long
   Dim R2        As Long, G2 As Long, b2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim I         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long
 
   If (Width > 0 And Height > 0) Then

'     Matthew R. Usner - solves weird problem of when angle is
'     >= 91 and <= 270, the colors invert in MiddleOut mode.
      If bMOut And Angle >= 91 And Angle <= 270 Then
         g = Color1
         Color1 = Color2
         Color2 = g
      End If

'     -- Right-hand [+] (ox=0º)
      Angle = -Angle + 90

'     -- Normalize to [0º;360º]
      Angle = Angle Mod 360
      If (Angle < 0) Then
         Angle = 360 + Angle
      End If

'     -- Get quadrant (0 - 3)
      lQuad = Angle \ 90

'     -- Normalize to [0º;90º]
        Angle = Angle Mod 90

'     -- Calc. gradient length ('distance')
      If (lQuad Mod 2 = 0) Then
         AngleDiag = Atn(Width / Height) * TO_DEG
      Else
         AngleDiag = Atn(Height / Width) * TO_DEG
      End If
      AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
      Angle = Angle * TO_RAD
      g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'     -- Decompose colors
      If (lQuad > 1) Then
         lClr = Color1
         Color1 = Color2
         Color2 = lClr
      End If
      R1 = (Color1 And &HFF&)
      G1 = (Color1 And &HFF00&) \ 256
      b1 = (Color1 And &HFF0000) \ 65536
      R2 = (Color2 And &HFF&)
      G2 = (Color2 And &HFF00&) \ 256
      b2 = (Color2 And &HFF0000) \ 65536

'     -- Get color distances
      dR = R2 - R1
      dG = G2 - G1
      dB = b2 - b1

'     -- Size gradient-colors array
      ReDim lGrad(0 To g - 1)
      ReDim lGrad2(0 To g - 1)

'     -- Calculate gradient-colors
      iEnd = g - 1
      If (iEnd = 0) Then
'        -- Special case (1-pixel wide gradient)
         lGrad2(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
         For I = 0 To iEnd
            lGrad2(I) = b1 + (dB * I) \ iEnd + 256 * (G1 + (dG * I) \ iEnd) + 65536 * (R1 + (dR * I) \ iEnd)
         Next I
      End If

'     'if block' added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
      If bMOut Then
         k = 0
         For I = 0 To iEnd Step 2
            lGrad(k) = lGrad2(I)
            k = k + 1
         Next I
         For I = iEnd - 1 To 1 Step -2
            lGrad(k) = lGrad2(I)
            k = k + 1
         Next I
      Else
         For I = 0 To iEnd
            lGrad(I) = lGrad2(I)
         Next I
      End If

'     -- Size DIB array
      ReDim lBits(Width * Height - 1) As Long
      iEnd = Width - 1
      jEnd = Height - 1
      Scan = Width

'     -- Render gradient DIB
      Select Case lQuad

         Case 0, 2
            luSin = Sin(Angle) * INT_ROT
            luCos = Cos(Angle) * INT_ROT
            Offset = 0
            jIn = 0
            For j = 0 To jEnd
               For I = 0 To iEnd
                  lBits(I + Offset) = lGrad((I * luSin + jIn) \ INT_ROT)
               Next I
               jIn = jIn + luCos
               Offset = Offset + Scan
            Next j

         Case 1, 3
            luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
            luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
            Offset = jEnd * Scan
            jIn = 0
            For j = 0 To jEnd
               For I = 0 To iEnd
                  lBits(I + Offset) = lGrad((I * luSin + jIn) \ INT_ROT)
               Next I
               jIn = jIn + luCos
               Offset = Offset - Scan
            Next j

      End Select

'     -- Define DIB header
      With uBIH
         .biSize = 40
         .biPlanes = 1
         .biBitCount = 32
         .biWidth = Width
         .biHeight = Height
      End With

'     -- Paint it!
      Call StretchDIBits(hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

    End If

End Sub

Private Sub InitControlDisplay()

'*************************************************************************
'* gets appropriate display settings for control.                        *
'*************************************************************************

'  save the default border and checkbox border colors so that when the
'  mouse cursor leaves the control we can restore the original color(s).
   SaveBorderColor = m_BorderColor
   SaveCheckBoxBorderColor = m_CheckBorderColor

'  determine the x coordinate of the left edge of the checkbox.
   GetCheckBoxX

'  load appropriate display properties based on enabled state of control.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

   RedrawControl

End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Properties  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to default values.                              *
'*************************************************************************

   Set m_Font = Ambient.Font
   Set m_PicChecked = LoadPicture("")
   Set m_PicUnchecked = LoadPicture("")
   m_BackAngle = m_def_BackAngle
   m_BackColor1 = m_def_BackColor1
   m_BackColor2 = m_def_BackColor2
   m_BackMiddleOut = m_def_BackMiddleOut
   m_BorderColor = m_def_BorderColor
   m_BorderCurvature = m_def_BorderCurvature
   m_BorderIfTransparent = m_def_BorderIfTransparent
   m_BorderWidth = m_def_BorderWidth
   m_Caption = m_def_Caption
   m_CaptionColor = m_def_CaptionColor
   m_CheckBorderColor = m_def_CheckBorderColor
   m_CheckBoxAlignment = m_def_CheckBoxAlignment
   m_CheckBoxAngle = m_def_CheckBoxAngle
   m_CheckBoxColor1 = m_def_CheckBoxColor1
   m_CheckBoxColor2 = m_def_CheckBoxColor2
   m_CheckBoxMiddleOut = m_def_CheckBoxMiddleOut
   m_CheckColor = m_def_CheckColor
   m_ContainerName = m_def_ContainerName
   m_ControlType = m_def_ControlType
   m_DisBackColor1 = m_def_DisBackColor1
   m_DisBackColor2 = m_def_DisBackColor2
   m_DisBorderColor = m_def_DisBorderColor
   m_DisCaptionColor = m_def_DisCaptionColor
   m_DisCheckBorderColor = m_def_DisCheckBorderColor
   m_DisCheckBoxColor1 = m_def_DisCheckBoxColor1
   m_DisCheckBoxColor2 = m_def_DisCheckBoxColor2
   m_Enabled = m_def_Enabled
   m_FocusRectColor = m_def_FocusRectColor
   m_MouseOverActions = m_def_MouseOverActions
   m_MOverBorderColor = m_def_MOverBorderColor
   m_MOverCheckBoxColor = m_def_MOverCheckBoxColor
   m_PicForCheck = m_def_PicForCheck
   m_ShowFocusRect = m_def_ShowFocusRect
   m_Transparent = m_def_Transparent
   m_Value = m_def_Value

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* load property values from storage.                                    *
'*************************************************************************

   With PropBag
      Set m_Font = .ReadProperty("Font", Ambient.Font)
      Set UserControl.Font = m_Font
      Set m_PicChecked = .ReadProperty("PicChecked", Nothing)
      Set m_PicUnchecked = .ReadProperty("PicUnchecked", Nothing)
      m_BackAngle = .ReadProperty("BackAngle", m_def_BackAngle)
      m_BackColor1 = .ReadProperty("BackColor1", m_def_BackColor1)
      m_BackColor2 = .ReadProperty("BackColor2", m_def_BackColor2)
      m_BackMiddleOut = .ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
      m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
      m_BorderCurvature = .ReadProperty("BorderCurvature", m_def_BorderCurvature)
      m_BorderIfTransparent = .ReadProperty("BorderIfTransparent", m_def_BorderIfTransparent)
      m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
      m_Caption = .ReadProperty("Caption", m_def_Caption)
      m_CaptionColor = .ReadProperty("CaptionColor", m_def_CaptionColor)
      m_CheckBorderColor = .ReadProperty("CheckBorderColor", m_def_CheckBorderColor)
      m_CheckBoxAlignment = .ReadProperty("CheckBoxAlignment", m_def_CheckBoxAlignment)
      m_CheckBoxAngle = .ReadProperty("CheckBoxAngle", m_def_CheckBoxAngle)
      m_CheckBoxColor1 = .ReadProperty("CheckBoxColor1", m_def_CheckBoxColor1)
      m_CheckBoxColor2 = .ReadProperty("CheckBoxColor2", m_def_CheckBoxColor2)
      m_CheckBoxMiddleOut = .ReadProperty("CheckBoxMiddleOut", m_def_CheckBoxMiddleOut)
      m_CheckColor = .ReadProperty("CheckColor", m_def_CheckColor)
      m_ContainerName = .ReadProperty("ContainerName", m_def_ContainerName)
      m_ControlType = .ReadProperty("ControlType", m_def_ControlType)
      m_DisBackColor1 = .ReadProperty("DisBackColor1", m_def_DisBackColor1)
      m_DisBackColor2 = .ReadProperty("DisBackColor2", m_def_DisBackColor2)
      m_DisBorderColor = .ReadProperty("DisBorderColor", m_def_DisBorderColor)
      m_DisCaptionColor = .ReadProperty("DisCaptionColor", m_def_DisCaptionColor)
      m_DisCheckBorderColor = .ReadProperty("DisCheckBorderColor", m_def_DisCheckBorderColor)
      m_DisCheckBoxColor1 = .ReadProperty("DisCheckBoxColor1", m_def_DisCheckBoxColor1)
      m_DisCheckBoxColor2 = .ReadProperty("DisCheckBoxColor2", m_def_DisCheckBoxColor2)
      m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
      m_FocusRectColor = .ReadProperty("FocusRectColor", m_def_FocusRectColor)
      m_MouseOverActions = .ReadProperty("MouseOverActions", m_def_MouseOverActions)
      m_MOverBorderColor = .ReadProperty("MOverBorderColor", m_def_MOverBorderColor)
      m_MOverCheckBoxColor = .ReadProperty("MOverCheckBoxColor", m_def_MOverCheckBoxColor)
      m_PicForCheck = .ReadProperty("PicForCheck", m_def_PicForCheck)
      m_ShowFocusRect = .ReadProperty("ShowFocusRect", m_def_ShowFocusRect)
      m_Transparent = .ReadProperty("Transparent", m_def_Transparent)
      m_Value = .ReadProperty("Value", m_def_Value)
   End With

   InitControlDisplay

'  start up the subclassing if not in design mode.
   If Ambient.UserMode Then
      bTrack = True
      With UserControl
         sc_Subclass .hwnd                            ' subclass the control's window handle.
         sc_AddMsg .hwnd, WM_MOUSEMOVE, MSG_AFTER     ' for mouse enter detect.
         sc_AddMsg .hwnd, WM_MOUSELEAVE, MSG_AFTER    ' for mouse leave detect.
         sc_AddMsg .hwnd, WM_SETFOCUS, MSG_AFTER      ' for got focus detect.
         sc_AddMsg .hwnd, WM_KILLFOCUS, MSG_AFTER     ' for lost focus detect.
      End With
   
   End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write property values to storage.                                     *
'*************************************************************************

   With PropBag
      .WriteProperty "BackAngle", m_BackAngle, m_def_BackAngle
      .WriteProperty "BackColor1", m_BackColor1, m_def_BackColor1
      .WriteProperty "BackColor2", m_BackColor2, m_def_BackColor2
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut
      .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
      .WriteProperty "BorderCurvature", m_BorderCurvature, m_def_BorderCurvature
      .WriteProperty "BorderIfTransparent", m_BorderIfTransparent, m_def_BorderIfTransparent
      .WriteProperty "BorderWidth", m_BorderWidth, m_def_BorderWidth
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "CaptionColor", m_CaptionColor, m_def_CaptionColor
      .WriteProperty "CheckBorderColor", m_CheckBorderColor, m_def_CheckBorderColor
      .WriteProperty "CheckBoxAlignment", m_CheckBoxAlignment, m_def_CheckBoxAlignment
      .WriteProperty "CheckBoxAngle", m_CheckBoxAngle, m_def_CheckBoxAngle
      .WriteProperty "CheckBoxColor1", m_CheckBoxColor1, m_def_CheckBoxColor1
      .WriteProperty "CheckBoxColor2", m_CheckBoxColor2, m_def_CheckBoxColor2
      .WriteProperty "CheckBoxMiddleOut", m_CheckBoxMiddleOut, m_def_CheckBoxMiddleOut
      .WriteProperty "CheckColor", m_CheckColor, m_def_CheckColor
      .WriteProperty "ContainerName", m_ContainerName, m_def_ContainerName
      .WriteProperty "ControlType", m_ControlType, m_def_ControlType
      .WriteProperty "DisBackColor1", m_DisBackColor1, m_def_DisBackColor1
      .WriteProperty "DisBackColor2", m_DisBackColor2, m_def_DisBackColor2
      .WriteProperty "DisBorderColor", m_DisBorderColor, m_def_DisBorderColor
      .WriteProperty "DisCaptionColor", m_DisCaptionColor, m_def_DisCaptionColor
      .WriteProperty "DisCheckBorderColor", m_DisCheckBorderColor, m_def_DisCheckBorderColor
      .WriteProperty "DisCheckBoxColor1", m_DisCheckBoxColor1, m_def_DisCheckBoxColor1
      .WriteProperty "DisCheckBoxColor2", m_DisCheckBoxColor2, m_def_DisCheckBoxColor2
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "FocusRectColor", m_FocusRectColor, m_def_FocusRectColor
      .WriteProperty "Font", m_Font, Ambient.Font
      .WriteProperty "MouseOverActions", m_MouseOverActions, m_def_MouseOverActions
      .WriteProperty "MOverBorderColor", m_MOverBorderColor, m_def_MOverBorderColor
      .WriteProperty "MOverCheckBoxColor", m_MOverCheckBoxColor, m_def_MOverCheckBoxColor
      .WriteProperty "PicChecked", m_PicChecked, Nothing
      .WriteProperty "PicForCheck", m_PicForCheck, m_def_PicForCheck
      .WriteProperty "PicUnchecked", m_PicUnchecked, Nothing
      .WriteProperty "ShowFocusRect", m_ShowFocusRect, m_def_ShowFocusRect
      .WriteProperty "Transparent", m_Transparent, m_def_Transparent
      .WriteProperty "Value", m_Value, m_def_Value
   End With

End Sub

Public Property Get BackAngle() As Single
Attribute BackAngle.VB_Description = "The angle of the control's background gradient."
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   RedrawControl
End Property

Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "The first color of the background gradient."
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   If m_Enabled Then
      ActiveBackColor1 = New_BackColor1
   End If
   PropertyChanged "BackColor1"
   RedrawControl
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "The second color of the background gradient."
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   If m_Enabled Then
      ActiveBackColor2 = New_BackColor2
   End If
   PropertyChanged "BackColor2"
   RedrawControl
End Property

Public Property Get BackMiddleOut() As Boolean
Attribute BackMiddleOut.VB_Description = "The middle-out status of the background gradient.  If True, gradient goes from Color1>Color2>Color1."
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "The color of the control's border."
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   If m_Enabled Then
      ActiveBorderColor = New_BorderColor
   End If
   PropertyChanged "BorderColor"
   RedrawControl
End Property

Public Property Get BorderCurvature() As Integer
Attribute BorderCurvature.VB_Description = "The amount of curvature of the control's corners.  0 means no curvature."
   BorderCurvature = m_BorderCurvature
End Property

Public Property Let BorderCurvature(ByVal New_BorderCurvature As Integer)
   m_BorderCurvature = New_BorderCurvature
   PropertyChanged "BorderCurvature"
   RedrawControl
End Property

Public Property Get BorderIfTransparent() As Boolean
   BorderIfTransparent = m_BorderIfTransparent
End Property

Public Property Let BorderIfTransparent(ByVal New_BorderIfTransparent As Boolean)
   m_BorderIfTransparent = New_BorderIfTransparent
   PropertyChanged "BorderIfTransparent"
   RedrawControl
End Property

Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_Description = "The width, in pixels, of the control's border."
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Long)
   m_BorderWidth = New_BorderWidth
   PropertyChanged "BorderWidth"
   RedrawControl
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text to display in the control."
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal new_caption As String)
   m_Caption = new_caption
   PropertyChanged "Caption"
   RedrawControl
End Property

Public Property Get CaptionColor() As OLE_COLOR
Attribute CaptionColor.VB_Description = "The color of the caption text."
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   If m_Enabled Then
      ActiveCaptionColor = New_CaptionColor
   End If
   PropertyChanged "CaptionColor"
   RedrawControl
End Property

Public Property Get CheckBorderColor() As OLE_COLOR
Attribute CheckBorderColor.VB_Description = "The color of the 1-pixel wide checkbox border."
   CheckBorderColor = m_CheckBorderColor
End Property

Public Property Let CheckBorderColor(ByVal New_CheckBorderColor As OLE_COLOR)
   m_CheckBorderColor = New_CheckBorderColor
   If m_Enabled Then
      ActiveCheckBorderColor = New_CheckBorderColor
   End If
   PropertyChanged "CheckBorderColor"
   RedrawControl
End Property

Public Property Get CheckBoxAlignment() As CBAlignmentOptions
Attribute CheckBoxAlignment.VB_Description = "Aligns the checkbox to the left or right side of the control."
   CheckBoxAlignment = m_CheckBoxAlignment
End Property

Public Property Let CheckBoxAlignment(ByVal New_CheckBoxAlignment As CBAlignmentOptions)
   m_CheckBoxAlignment = New_CheckBoxAlignment
   PropertyChanged "CheckBoxAlignment"
   GetCheckBoxX    '  determine the x coordinate of the left edge of the checkbox.
   RedrawControl
End Property

Public Property Get CheckBoxAngle() As Single
Attribute CheckBoxAngle.VB_Description = "The angle of the gradient in the checkbox."
   CheckBoxAngle = m_CheckBoxAngle
End Property

Public Property Let CheckBoxAngle(ByVal New_CheckBoxAngle As Single)
   m_CheckBoxAngle = New_CheckBoxAngle
   PropertyChanged "CheckBoxAngle"
   RedrawControl
End Property

Public Property Get CheckBoxColor1() As OLE_COLOR
Attribute CheckBoxColor1.VB_Description = "The first color of the checkbox gradient."
   CheckBoxColor1 = m_CheckBoxColor1
End Property

Public Property Let CheckBoxColor1(ByVal New_CheckBoxColor1 As OLE_COLOR)
   m_CheckBoxColor1 = New_CheckBoxColor1
   If m_Enabled Then
      ActiveCheckBoxColor1 = New_CheckBoxColor1
   End If
   PropertyChanged "CheckBoxColor1"
   RedrawControl
End Property

Public Property Get CheckBoxColor2() As OLE_COLOR
Attribute CheckBoxColor2.VB_Description = "The second color of the checkbox gradient."
   CheckBoxColor2 = m_CheckBoxColor2
End Property

Public Property Let CheckBoxColor2(ByVal New_CheckBoxColor2 As OLE_COLOR)
   m_CheckBoxColor2 = New_CheckBoxColor2
   If m_Enabled Then
      ActiveCheckBoxColor2 = New_CheckBoxColor2
   End If
   PropertyChanged "CheckBoxColor2"
   RedrawControl
End Property

Public Property Get CheckBoxMiddleOut() As Boolean
Attribute CheckBoxMiddleOut.VB_Description = "The middle-out status of the checkbox gradient.  If True, gradient goes from Color1>Color2>Color1."
   CheckBoxMiddleOut = m_CheckBoxMiddleOut
End Property

Public Property Let CheckBoxMiddleOut(ByVal New_CheckBoxMiddleOut As Boolean)
   m_CheckBoxMiddleOut = New_CheckBoxMiddleOut
   PropertyChanged "CheckBoxMiddleOut"
   RedrawControl
End Property

Public Property Get CheckColor() As OLE_COLOR
Attribute CheckColor.VB_Description = "The color of the check symbol."
   CheckColor = m_CheckColor
End Property

Public Property Let CheckColor(ByVal New_CheckColor As OLE_COLOR)
   m_CheckColor = New_CheckColor
   PropertyChanged "CheckColor"
   RedrawControl
End Property

Public Property Get ContainerName() As String
   ContainerName = m_ContainerName
End Property

Public Property Let ContainerName(ByVal New_ContainerName As String)
   m_ContainerName = New_ContainerName
   PropertyChanged "ContainerName"
   RedrawControl
End Property

Public Property Get ControlType() As ControlTypeOptions
Attribute ControlType.VB_Description = "Lets user determine which control to emulate (CheckBox or OptionButton)."
   ControlType = m_ControlType
End Property

Public Property Let ControlType(ByVal New_ControlType As ControlTypeOptions)
   If Ambient.UserMode Then Err.Raise 382    ' property is read-only at runtime.
   m_ControlType = New_ControlType
   PropertyChanged "ControlType"
End Property

Public Property Get DisBackColor1() As OLE_COLOR
Attribute DisBackColor1.VB_Description = "The first background gradient color when the control is disabled."
Attribute DisBackColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DisBackColor1 = m_DisBackColor1
End Property

Public Property Let DisBackColor1(ByVal New_DisBackColor1 As OLE_COLOR)
   m_DisBackColor1 = New_DisBackColor1
   If Not m_Enabled Then
      ActiveBackColor1 = New_DisBackColor1
   End If
   PropertyChanged "DisBackColor1"
   RedrawControl
End Property

Public Property Get DisBackColor2() As OLE_COLOR
Attribute DisBackColor2.VB_Description = "The second background gradient color when the control is disabled."
Attribute DisBackColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DisBackColor2 = m_DisBackColor2
End Property

Public Property Let DisBackColor2(ByVal New_DisBackColor2 As OLE_COLOR)
   m_DisBackColor2 = New_DisBackColor2
   If Not m_Enabled Then
      ActiveBackColor2 = New_DisBackColor2
   End If
   PropertyChanged "DisBackColor2"
   RedrawControl
End Property

Public Property Get DisBorderColor() As OLE_COLOR
Attribute DisBorderColor.VB_Description = "The border color when the control is disabled."
Attribute DisBorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DisBorderColor = m_DisBorderColor
End Property

Public Property Let DisBorderColor(ByVal New_DisBorderColor As OLE_COLOR)
   m_DisBorderColor = New_DisBorderColor
   If Not m_Enabled Then
      ActiveBorderColor = New_DisBorderColor
   End If
   PropertyChanged "DisBorderColor"
   RedrawControl
End Property

Public Property Get DisCaptionColor() As OLE_COLOR
Attribute DisCaptionColor.VB_Description = "The caption color when the control is disabled."
Attribute DisCaptionColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DisCaptionColor = m_DisCaptionColor
End Property

Public Property Let DisCaptionColor(ByVal New_DisCaptionColor As OLE_COLOR)
   m_DisCaptionColor = New_DisCaptionColor
   If Not m_Enabled Then
      ActiveCaptionColor = New_DisCaptionColor
   End If
   PropertyChanged "DisCaptionColor"
   RedrawControl
End Property

Public Property Get DisCheckBorderColor() As OLE_COLOR
Attribute DisCheckBorderColor.VB_Description = "The checkbox border color when the control is disabled."
Attribute DisCheckBorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DisCheckBorderColor = m_DisCheckBorderColor
End Property

Public Property Let DisCheckBorderColor(ByVal New_DisCheckBorderColor As OLE_COLOR)
   m_DisCheckBorderColor = New_DisCheckBorderColor
   If Not m_Enabled Then
      ActiveCheckBorderColor = New_DisCheckBorderColor
   End If
   PropertyChanged "DisCheckBorderColor"
   RedrawControl
End Property

Public Property Get DisCheckBoxColor1() As OLE_COLOR
Attribute DisCheckBoxColor1.VB_Description = "The first gradient color of the checkbox when the control is disabled."
Attribute DisCheckBoxColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DisCheckBoxColor1 = m_DisCheckBoxColor1
End Property

Public Property Let DisCheckBoxColor1(ByVal New_DisCheckBoxColor1 As OLE_COLOR)
   m_DisCheckBoxColor1 = New_DisCheckBoxColor1
   If Not m_Enabled Then
      ActiveCheckBoxColor1 = New_DisCheckBoxColor1
   End If
   PropertyChanged "DisCheckBoxColor1"
   RedrawControl
End Property

Public Property Get DisCheckBoxColor2() As OLE_COLOR
Attribute DisCheckBoxColor2.VB_Description = "The second gradient color of the checkbox when the control is disabled."
Attribute DisCheckBoxColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
   DisCheckBoxColor2 = m_DisCheckBoxColor2
End Property

Public Property Let DisCheckBoxColor2(ByVal New_DisCheckBoxColor2 As OLE_COLOR)
   m_DisCheckBoxColor2 = New_DisCheckBoxColor2
   If Not m_Enabled Then
      ActiveCheckBoxColor2 = New_DisCheckBoxColor2
   End If
   PropertyChanged "DisCheckBoxColor2"
   RedrawControl
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/returns whether or not the control can be accessed."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Misc"
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal new_Enabled As Boolean)
   m_Enabled = new_Enabled
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
'     when the control is disabled, it cannot be checked.  set the value
'     to unchecked, according to the type of control being displayed.
      If m_ControlType = CheckBox Then
         m_Value = vbUnchecked
      Else
         m_Value = False
      End If
      GetDisabledDisplayProperties
   End If
   PropertyChanged "Enabled"
   RedrawControl
End Property

Public Property Get FocusRectColor() As OLE_COLOR
Attribute FocusRectColor.VB_Description = "The color of the 1-pixel wide custom focus rectangle."
   FocusRectColor = m_FocusRectColor
End Property

Public Property Let FocusRectColor(ByVal New_FocusRectColor As OLE_COLOR)
   m_FocusRectColor = New_FocusRectColor
   PropertyChanged "FocusRectColor"
   RedrawControl
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "The font to display the caption text with."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Font.VB_UserMemId = -512
   Set Font = m_Font
End Property

Public Property Set Font(ByVal new_font As StdFont)
   Set m_Font = new_font
   Set UserControl.Font = m_Font
   PropertyChanged "Font"
   RedrawControl
End Property

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "The device context for the control."
   hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "The window handle for the control."
   hwnd = UserControl.hwnd
End Property

Public Property Get MouseOverActions() As MouseOverOptions
Attribute MouseOverActions.VB_Description = "Determines which parts of the control's display change when mouse pointer enters control (border, checkbox border, or both)."
   MouseOverActions = m_MouseOverActions
End Property

Public Property Let MouseOverActions(ByVal New_MouseOverActions As MouseOverOptions)
   m_MouseOverActions = New_MouseOverActions
   PropertyChanged "MouseOverActions"
   RedrawControl
End Property

Public Property Get MOverBorderColor() As OLE_COLOR
Attribute MOverBorderColor.VB_Description = "The color of the border when the mouse pointer has entered the control."
   MOverBorderColor = m_MOverBorderColor
End Property

Public Property Let MOverBorderColor(ByVal New_MOverBorderColor As OLE_COLOR)
   m_MOverBorderColor = New_MOverBorderColor
   PropertyChanged "MOverBorderColor"
   RedrawControl
End Property

Public Property Get MOverCheckBoxColor() As OLE_COLOR
Attribute MOverCheckBoxColor.VB_Description = "The color of the checkbox border when the mouse pointer has entered the control."
   MOverCheckBoxColor = m_MOverCheckBoxColor
End Property

Public Property Let MOverCheckBoxColor(ByVal New_MOverCheckBoxColor As OLE_COLOR)
   m_MOverCheckBoxColor = New_MOverCheckBoxColor
   PropertyChanged "MOverCheckBoxColor"
   RedrawControl
End Property

Public Property Get PicChecked() As Picture
Attribute PicChecked.VB_Description = "The icon to display when the control is checked."
   Set PicChecked = m_PicChecked
End Property

Public Property Set PicChecked(ByVal New_PicChecked As Picture)
   Set m_PicChecked = New_PicChecked
   PropertyChanged "PicChecked"
   RedrawControl
End Property

Public Property Get PicForCheck() As Boolean
Attribute PicForCheck.VB_Description = "If True, displays user-selected icons in lieu of standard checkmarks."
   PicForCheck = m_PicForCheck
End Property

Public Property Let PicForCheck(ByVal New_PicForCheck As Boolean)
   m_PicForCheck = New_PicForCheck
   PropertyChanged "PicForCheck"
   RedrawControl
End Property

Public Property Get PicUnchecked() As Picture
Attribute PicUnchecked.VB_Description = "The icon to display when the control is unchecked."
   Set PicUnchecked = m_PicUnchecked
End Property

Public Property Set PicUnchecked(ByVal New_PicUnchecked As Picture)
   Set m_PicUnchecked = New_PicUnchecked
   PropertyChanged "PicUnchecked"
   RedrawControl
End Property

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "If True, displays a custom 1-pixel focus rectangle around the caption text when the control has the focus."
   ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
   m_ShowFocusRect = New_ShowFocusRect
   PropertyChanged "ShowFocusRect"
   RedrawControl
End Property

Public Property Get Transparent() As Boolean
   Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
   m_Transparent = New_Transparent
   PropertyChanged "Transparent"
   RedrawControl
End Property

Public Property Get Value() As Integer
Attribute Value.VB_Description = "MorphOptionCheck status (Checked, Unchecked, True, False)."
   Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
   m_Value = New_Value
   PropertyChanged "Value"
   RedrawControl
End Property

Private Sub GetEnabledDisplayProperties()

'*************************************************************************
'* applies enabled graphics properties to the active display properties. *
'*************************************************************************

   ActiveBackColor1 = m_BackColor1
   ActiveBackColor2 = m_BackColor2
   ActiveBorderColor = m_BorderColor
   ActiveCaptionColor = m_CaptionColor
   ActiveCheckBorderColor = m_CheckBorderColor
   ActiveCheckBoxColor1 = m_CheckBoxColor1
   ActiveCheckBoxColor2 = m_CheckBoxColor2

End Sub

Private Sub GetDisabledDisplayProperties()

'*************************************************************************
'* applies disabled graphics properties to active display properties.    *
'*************************************************************************

   ActiveBackColor1 = m_DisBackColor1
   ActiveBackColor2 = m_DisBackColor2
   ActiveBorderColor = m_DisBorderColor
   ActiveCaptionColor = m_DisCaptionColor
   ActiveCheckBorderColor = m_DisCheckBorderColor
   ActiveCheckBoxColor1 = m_DisCheckBoxColor1
   ActiveCheckBoxColor2 = m_DisCheckBoxColor2

End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Subclassing  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByVal nOrdinal As Long = 1) As Boolean         'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - User-defined callback parameter
'* bIdeSafety - Enable/disable IDE safety measures. Generally, bIdeSafety should be left as True, it's only necessary to disable IDE safety in a UserControl for design-time subclassing
'* nOrdinal   - Ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'*************************************************************************************************
Const CODE_LEN      As Long = 248                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data storage index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data storage index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data storage index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data storage index of the VirtualFree function address
Const IDX_OWNER     As Long = 7                                             'Thunk data storage index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 9                                             'Thunk data storage index of the callback method address
Const IDX_EBX       As Long = 15                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get our process ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    nAddr = zAddressOf(nOrdinal)                                            'Get the address of the specified ordinal method
    
    If nAddr = 0 Then                                                       'Ensure that we've found the ordinal method
      zError SUB_NAME, "Callback method not found"
      Exit Function
    End If
    
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    
    'Initialize the thunk machine-code
    z_SC(13) = &HD231C031: z_SC(14) = &HBBE58960: z_SC(16) = &H4339F631: z_SC(17) = &H4A21750C: z_SC(18) = &HE8287B8B: z_SC(19) = &H74&: z_SC(20) = &H75147539: z_SC(21) = &H21E80F: z_SC(22) = &HD2310000: z_SC(23) = &HE82C7B8B: z_SC(24) = &H60&: z_SC(25) = &H10C261: z_SC(26) = &H830C53FF: z_SC(27) = &HD77401F8: z_SC(28) = &H2874C085: z_SC(29) = &H2E8&: z_SC(30) = &HFFE9EB00: z_SC(31) = &H75FF3075: z_SC(32) = &H2875FF2C: z_SC(33) = &HFF2475FF: z_SC(34) = &H3FF2073: z_SC(35) = &H891053FF: z_SC(36) = &HBFF1C45: z_SC(37) = &H73395F75
    z_SC(38) = &H585A7404: z_SC(39) = &H6A2073FF: z_SC(40) = &H873FFFC: z_SC(41) = &H891453FF: z_SC(42) = &H7589285D: z_SC(43) = &H3045C72C: z_SC(44) = &H8000&: z_SC(45) = &H8920458B: z_SC(46) = &H4589145D: z_SC(47) = &HC4836124: z_SC(48) = &H1862FF04: z_SC(49) = &H2DE30F8B: z_SC(50) = &HA78C985: z_SC(51) = &H8B04C783: z_SC(52) = &HAFF22845: z_SC(53) = &H438D1F75: z_SC(54) = &H144D8D30: z_SC(55) = &H1C458D50: z_SC(56) = &HFF3075FF: z_SC(57) = &H75FF2C75: z_SC(58) = &H873FF28: z_SC(59) = &HFF525150: z_SC(60) = &H53FF1C73: z_SC(61) = &HC324&

    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_SC(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_SC(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_SC(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_SC(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_SC(IDX_OWNER) = ObjPtr(Me)                                            'Store my object address in the thunk data
    z_SC(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
  End If
  
  z_Base = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                   'Allocate executable memory

  If z_Base <> 0 Then                                                       'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_Base, "h" & lng_hWnd                                     'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    z_SC(IDX_EBX) = z_Base                                                  'Patch the thunk data address
    z_SC(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_SC(IDX_BTABLE) = z_Base + CODE_LEN                                    'Store the address of the before table in the thunk data
    z_SC(IDX_ATABLE) = z_Base + CODE_LEN + ((MSG_ENTRIES + 1) * 4)          'Store the address of the after table in the thunk data
    z_SC(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_Base + WNDPROC_OFF)     'Set the new WndProc, return the address of the original WndProc
    
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_SC(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_Base, VarPtr(z_SC(0)), CODE_LEN                         'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_Base, 0, MEM_RELEASE
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim I     As Long
  Dim nAddr As Long

  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
  Else
    With z_Funk
      For I = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        nAddr = .Item(I)                                                    'Map zData() to the hWnd thunk address
        If IsBadCodePtr(nAddr) = 0 Then                                     'Ensure that the thunk hasn't already released its memory
          z_Base = nAddr                                                    'Map the thunk memory to the zData() array
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next I                                                                'Next member of the collection
    End With
    
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg lng_hWnd, ALL_MESSAGES, IDX_BTABLE                            'Delete all before messages
      zDelMsg lng_hWnd, ALL_MESSAGES, IDX_ATABLE                            'Delete all after messages
    End If
    
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg lng_hWnd, uMsg, IDX_BTABLE                                    'Add the message to the before table
    End If
  
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg lng_hWnd, uMsg, IDX_ATABLE                                    'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg lng_hWnd, uMsg, IDX_BTABLE                                    'Delete the message from the before table
    End If
  
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg lng_hWnd, uMsg, IDX_ATABLE                                    'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim I      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map zData() to the thunk of the specified window handle
  z_Base = zData(nTable)                                                    'Map zData() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count

    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      Exit Sub
    End If

    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = 0 Then                                                  'If the element is free...
        zData(I) = uMsg                                                     'Use this element
        Exit Sub                                                            'Bail
      ElseIf zData(I) = uMsg Then                                           'If the message is already in the table...
        Exit Sub                                                            'Bail
      End If
    Next I                                                                  'Next message table entry

    nCount = I                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim I      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map zData() to the thunk of the specified window handle
  z_Base = zData(nTable)                                                    'Map zData() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = uMsg Then                                               'If the message is found...
        zData(I) = 0                                                        'Null the msg value -- also frees the element for re-use
        Exit Sub                                                            'Exit
      End If
    Next I                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "z_Base = zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_Base = z_Funk("h" & lng_hWnd)                                         'Get the thunk address
    zMap_hWnd = z_Base
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal private method, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(Optional ByVal nOrdinal As Long = 1) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  If z_TblEnd = 0 Then                                                      'First time through...
    RtlMoveMemory VarPtr(nAddr), ObjPtr(Me), 4                              'Get the address of this object instance
  
    If Not zProbe(nAddr + &H1C, I, bSub) Then                               'Probe for a Class method
      If Not zProbe(nAddr + &H6F8, I, bSub) Then                            'Probe for a Form method
        If Not zProbe(nAddr + &H7A4, I, bSub) Then                          'Probe for a UserControl method
          Exit Function                                                     'Bail...
        End If
      End If
    End If
    
    I = I + 4                                                               'Bump to the next entry
    j = I + 1024                                                            'Set a reasonable limit, scan 256 vTable entries
    
    Do While I < j
      RtlMoveMemory VarPtr(nAddr), I, 4                                     'Get the address stored in this vTable entry
      
      If IsBadCodePtr(nAddr) Then                                           'Is the entry an invalid code address?
        z_TblEnd = I                                                        'Cache the vTable end-point
        GoTo Found                                                          'Bad method signature, quit loop
      End If
  
      RtlMoveMemory VarPtr(bVal), nAddr, 1                                  'Get the byte pointed to by the vTable entry
      If bVal <> bSub Then                                                  'If the byte doesn't match the expected value...
        z_TblEnd = I                                                        'Cache the vTable end-point
        GoTo Found                                                          'Bad method signature, quit loop
      End If
      
      I = I + 4                                                             'Next vTable entry
    Loop
    
    Exit Function                                                           'Final method not found
  End If
  
Found:
  RtlMoveMemory VarPtr(zAddressOf), z_TblEnd - (nOrdinal * 4), 4            'Return the specified vTable entry address
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_Base + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_Base + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc(ByVal bBefore As Boolean, _
                     ByRef bHandled As Boolean, _
                     ByRef lReturn As Long, _
                     ByVal lng_hWnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal wParam As Long, _
                     ByVal lParam As Long, _
                     ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter.
'*************************************************************************************************
  
   Select Case uMsg

      Case WM_MOUSEMOVE
'        detect when mouse has entered the control.
         If m_Enabled And Not bInCtrl Then
            bInCtrl = True
            Call TrackMouseLeave(lng_hWnd)
            RaiseEvent MouseEnter
'           repaint control based on selected mouseover actions.
            Select Case m_MouseOverActions
               Case [Border]
                  ActiveBorderColor = m_MOverBorderColor
                  RedrawControl
               Case [CheckBox Border]
                  ActiveCheckBorderColor = m_MOverCheckBoxColor
                  RedrawControl
               Case [Both]
                  ActiveBorderColor = m_MOverBorderColor
                  ActiveCheckBorderColor = m_MOverCheckBoxColor
                  RedrawControl
            End Select
         End If

'     detect when mouse has left the control.
      Case WM_MOUSELEAVE
         bInCtrl = False
         RaiseEvent MouseLeave
'        restore default control appearance if any mouseover actions were specified.
         If m_MouseOverActions <> [None] Then
            ActiveBorderColor = SaveBorderColor
            ActiveCheckBorderColor = SaveCheckBoxBorderColor
            RedrawControl
         End If

'     detect when control has gained the focus.
      Case WM_SETFOCUS
         If m_Enabled Then
            HasFocus = True
            RedrawControl
         End If

'     detect when control has lost the focus.
      Case WM_KILLFOCUS
         HasFocus = False
         KeyIsDown = False
         MouseIsDown = False
         RedrawControl

   End Select
  
End Sub
