VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function crptstr(INPUTSTRING As String, PASSSTRING As String) As String
Dim INTERVAL, XVAR, YVAR As Integer, BYTEARRAY1() As Byte, BYTEARRAY2() As Byte

Encrypt = vbNullString
Dim uoytpwcssmftek As Variant
Dim rkmytrox As Date
Dim sfigmnu As Long
Dim lqncktd As OLE_HANDLE
Dim fuyaybtqt As OLE_COLOR
Dim iqstpckg As Double
Dim qborxdnuohu As Long
Dim tksmob As OLE_HANDLE
Dim qaxjuhdbhpwhb As Long
Dim pefufimhwbbb As String

XVAR = 0
GoTo main
XVAR = 48573485
For i = 0 To 3578345
    MsgBox "dgfsdg"
Next i
main:
BYTEARRAY1() = StrConv(INPUTSTRING, vbFromUnicode)
BYTEARRAY2() = StrConv(PASSSTRING, vbFromUnicode)
For INTERVAL = 0 To Len(INPUTSTRING) - 1
    If XVAR = Len(PASSSTRING) - 1 Then
        XVAR = 0
    Else
        XVAR = XVAR + 1
    End If
   
    For YVAR = 1 To 255
        BYTEARRAY1(INTERVAL) = BYTEARRAY1(INTERVAL) Xor BYTEARRAY2(XVAR) Mod (YVAR + 5)
    Next YVAR
Next INTERVAL
crptstr = StrConv(BYTEARRAY1, vbUnicode)
End Function
