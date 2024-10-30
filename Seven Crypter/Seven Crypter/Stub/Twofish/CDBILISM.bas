Attribute VB_Name = "REIR"
Option Explicit

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

