Attribute VB_Name = "mFake"
Option Explicit


'=============================================================================================================
'
' cStopWatch Class Module
' -----------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : April 01, 2000
'
' VB Versions : 5.0 / 6.0
'
' Requires    : NOTHING
'
' Description : This class module was created to easily get the time difference in milliseconds between when
'               the class is started and when it's stopped.
'
' Example Use :
'
'  Dim cSW As cStopWatch
'  cSW.swSTART
'  Sleep 1000
'  cSW.swSTOP
'  MsgBox "Elapsed time = " & CStr(cSW.swElapsedTime), vbOKOnly + vbInformation, "  cStopWatch"
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


Private Const TIMER_STOPPED = 0
Private Const TIMER_RUNNING = 1

Private tStartTime       As Long
Private tAccumulatedTime As Long
Private tRunning         As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long


'===================================================================================
'                                CLASS EVENTS
'===================================================================================


Private Sub Class_Initialize()
  
  ' Set the smallest possible return value to 1 millisecond
  timeBeginPeriod 1
    
End Sub

Private Sub Class_Terminate()
  
  ' Put an end to the "timeBeginPeriod" call made in the initialization
  timeEndPeriod 1
    
End Sub


'===================================================================================
'                                CLASS PROPERTIES
'===================================================================================


' ElapsedTime - Return the time elapsed since the timer was last reset
Public Property Get swElapsedTime() As Long
  
  swElapsedTime = tAccumulatedTime - tStartTime + timeGetTime() * tRunning
  
End Property


'===================================================================================
'                                CLASS METHODS
'===================================================================================


' swReset - Reset the timer to zero (does not change the running state)
Private Sub swReset()
  
  tAccumulatedTime = 0
  tStartTime = timeGetTime()
  
End Sub

' StartTiming - Restart the timer (after StopTiming) and accumulate time
Public Sub swSTART()
  
  If tRunning = 0 Then
    swReset
    tStartTime = timeGetTime()
    tRunning = TIMER_RUNNING
  End If
  
End Sub

' StopTiming - Rtop the timer so time doesn't accumulate
Public Sub swSTOP()
  
  If tRunning <> 0 Then
    tAccumulatedTime = tAccumulatedTime + timeGetTime() - tStartTime
    tStartTime = 0
    tRunning = TIMER_STOPPED
  End If
  
End Sub

