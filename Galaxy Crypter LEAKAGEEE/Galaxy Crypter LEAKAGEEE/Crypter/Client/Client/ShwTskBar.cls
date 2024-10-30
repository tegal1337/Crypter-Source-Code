VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShwTskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Win32 APIs used to toggle border styles.
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

' Used to get window style bits.
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

' Force total redraw that shows new styles.
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1

' Extended Style bits.
Private Const WS_EX_APPWINDOW = &H40000

' References to client form.
Private WithEvents m_Client As Form
Attribute m_Client.VB_VarHelpID = -1
Private WithEvents m_MdiClient As MDIForm
Attribute m_MdiClient.VB_VarHelpID = -1
Private m_hWnd As Long

' Member variables
Private m_AutoSysMenu As Boolean


' ************************************************
'  Initialize/Terminate
' ************************************************
Private Sub Class_Initialize()
   '
End Sub

Private Sub Class_Terminate()
   '
End Sub



Public Property Let ShowInTaskbar(ByVal Value As Boolean)
   ' Set WS_EX_APPWINDOW On or Off as requested.
   ' Toggling this value requires that we also toggle
   ' visibility, flipping the bit while invisible,
   ' forcing the taskbar to update on reshow.
   ' Using LockWindowUpdate prevents some flicker.
   Call LockWindowUpdate(m_hWnd)
   Call ShowWindow(m_hWnd, vbHide)
   Call FlipBitEx(WS_EX_APPWINDOW, Value)
   Call ShowWindow(m_hWnd, vbNormalFocus)
   Call LockWindowUpdate(0&)
End Property

Public Property Get ShowInTaskbar() As Boolean
   ' Return value of WS_EX_APPWINDOW bit.
   ShowInTaskbar = CBool(StyleEx And WS_EX_APPWINDOW)
End Property



Private Function FlipBit(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
   Dim nStyle As Long
   
   ' Retrieve current style bits.
   nStyle = GetWindowLong(m_hWnd, GWL_STYLE)
   
   ' Attempt to set requested bit On or Off,
   ' and redraw
   If Value Then
      nStyle = nStyle Or Bit
   Else
      nStyle = nStyle And Not Bit
   End If
   Call SetWindowLong(m_hWnd, GWL_STYLE, nStyle)
   Call Redraw
   
   ' Return success code.
   FlipBit = (nStyle = GetWindowLong(m_hWnd, GWL_STYLE))
End Function

Private Function FlipBitEx(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
   Dim nStyleEx As Long
   
   ' Retrieve current style bits.
   nStyleEx = GetWindowLong(m_hWnd, GWL_EXSTYLE)
   
   ' Attempt to set requested bit On or Off,
   ' and redraw.
   If Value Then
      nStyleEx = nStyleEx Or Bit
   Else
      nStyleEx = nStyleEx And Not Bit
   End If
   Call SetWindowLong(m_hWnd, GWL_EXSTYLE, nStyleEx)
   Call Redraw
   
   ' Return success code.
   FlipBitEx = (nStyleEx = GetWindowLong(m_hWnd, GWL_EXSTYLE))
End Function


Public Property Set Client(ByVal obj As Object)
   ' Clear cached handle.
   m_hWnd = 0
   
   ' Store object reference and handle to client.
   If TypeOf obj Is MDIForm Then
      Set m_Client = Nothing
      Set m_MdiClient = obj
      m_hWnd = m_MdiClient.hWnd
   ElseIf TypeOf obj Is Form Then
      Set m_MdiClient = Nothing
      Set m_Client = obj
      m_hWnd = m_Client.hWnd
   End If
End Property

Public Property Get Client() As Object
   ' Return reference to client.
   If Not m_Client Is Nothing Then
      Set Client = m_Client
   ElseIf Not m_MdiClient Is Nothing Then
      Set Client = m_MdiClient
   End If
End Property


Public Sub Redraw()
   ' Redraw window with new style.
   Const swpFlags As Long = _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or _
      SWP_NOZORDER Or SWP_NOSIZE
   SetWindowPos m_hWnd, 0, 0, 0, 0, 0, swpFlags
End Sub



