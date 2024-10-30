VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mFakeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'=============================================================================================================
'
' cTimer_NoSC Class Module
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
' SEE ALSO    : cTimer.cls / modTimer.bas
'               ( This version of the Timer class module is the No Subclassing version.  cTimer.cls /
'               modTimer.bas is the subclassing version )
'
' NOTE        : Use the Enabled property to start and stop this timer class
'
' WARNING     : You need to set the Enabled property to FALSE before closing a form or program that uses
'               this class - otherwise the loop that controls the time will keep it open.
'
' Description : This class module was created as a light-weight alternative to using the standard VB Timer
'               control.  You can either declare a variable in a form using the WithEvents keyword and use
'               this Class' Timer event in that form to run your timer functionality, or you can put your
'               code directly into this class' TimerEvent method below.
'
'               This class has advantages and disadvantages as listed below.
'
'
'   ADVANTAGES:                           |  DISADVANTAGES:
'_________________________________________|_______________________________________
'                                         |
'  1. You can use a timer in a program    | 1. If your system is slow, this timer
'  without having to have any forms to    | class will NOT be as accurate as the
'  put the control on                     | standard Timer control
'                                         |
'  2. If your system is a fast system,    | 2. In debug mode, you'll frequently get
'  this timer class will be more accurate | put into the loop in the StartTimer
'  than the standard Timer control *      | method which is what controls the time
'                                         | ( get out by pressing CTRL+SHIFT+F8 )
'  3. This class can double as a stop     |
'  watch in that you can at any point,    | 3. You need to set the Enabled property
'  after enabling the class, check the    | of this class to FALSE before exiting
'  elapsed time from start.               | your program or closing a form it's used
'                                         | on, or the looping used to time will keep
'  4. This class does not use subclassing | the program / form open.
'  like some other class based timers do. |
'  Subclassing makes your program hard to |
'  step through to debug and can easily   |
'  crash VB's design environment if you   |
'  are not carefull.                      |
'                                         |
'  5. You can declare a variable to       |
'  represent this class module in a form  |
'  using the WithEvents keyword and use   |
'  this class module just like a standard |
'  VB Timer control via this module's     |
'  Timer event.                           |
'_________________________________________|_______________________________________
'
'
' Example Use :
'
'  Private WithEvents cTime As cTimer_NoSC
'
'  Private Sub Command1_Click()
'    cTime.Enabled = Not cTime.Enabled
'  End Sub
'
'  Private Sub cTime_Timer()
'    Dim TheDisplay As String
'    If cTime.Enabled = False Then
'      Exit Sub
'    End If
'    TheDisplay = Text1.Text
'    Text1.Text = "Time = " & cTime.TimeElapsed & vbCrLf & TheDisplay
'  End Sub
'
'  Private Sub Form_Load()
'    Me.Show
'  ' Text1.MultiLine = True
'  ' Text1.ScrollBars = 3 ' Both
'    Set cTime = New cTimer_NoSC
'    cTime.Interval = 1000
'    cTime.Enabled = True
'  End Sub
'
'  Private Sub Form_Unload(Cancel As Integer)
'    cTime.Enabled = False
'  End Sub

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


Private tInterval As Long
Private tEnabled As Boolean
Private tTimeStart As Long

Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long


'=============================================================================================================
'                           CLASS EVENT DECLARATIONS
'=============================================================================================================


Public Event Timer()


'=============================================================================================================
'                           CLASS INITIALIZATION
'=============================================================================================================


' Class starts here
Private Sub Class_Initialize()
On Error Resume Next
  
  ' Set initial variables
  tInterval = 0
  tTimeStart = -1
  tEnabled = False
  
  ' Set the smallest possible time interval to 1 millisecond
  timeBeginPeriod 1
  
End Sub

' Class ends here
Private Sub Class_Terminate()
On Error Resume Next
  
  ' Terminating call for the "timeBeginPeriod" function
  timeEndPeriod 1
  
End Sub


'=============================================================================================================
'                           CLASS PROPERTIES
'=============================================================================================================


' Returns if the class is enabled or not
Public Property Get Enabled() As Boolean
On Error Resume Next
  
  Enabled = tEnabled
  
End Property

' Sets if the class is enabled or not
Public Property Let Enabled(ByVal NewValue As Boolean)
On Error Resume Next
  
  ' Turn on the timer if set to true, distroy it if it's set to false
  If NewValue = True And tEnabled = False Then
    tEnabled = True
    If tInterval > 0 Then
      tTimeStart = timeGetTime
      StartTimer
      Exit Property
    End If
  ElseIf NewValue = False Then
    tEnabled = False
    tTimeStart = -1
  End If
  
End Property

' Returns what the currently set interval is
Public Property Get Interval() As Long
  
  Interval = tInterval
  
End Property

' Sets what the timer's interval is
Public Property Let Interval(ByVal NewValue As Long)
  
  ' Make sure the interval set is valid
  If NewValue <= 0 Then
    tTimeStart = -1
    tInterval = 0
  Else
    tInterval = NewValue
  End If
  
End Property

' This property returns how much time in milliseconds has elapsed since
' the timer was started
Public Property Get TimeElapsed() As Long
  
  If tEnabled = False Or tTimeStart = -1 Then TimeElapsed = 0 Else TimeElapsed = Abs(timeGetTime - tTimeStart)
  
End Property


'=============================================================================================================
'                               CLASS METHODS
'=============================================================================================================


' This function is ment to find out what is the slowest increment of time
' in milliseconds your computer can accurately report time based on a
' time test.  Slower computers running this class module will not have
' the accuracy of faster computers because the timer events are based on
' a loop, not a subclass.
Public Function TestSmallestInterval() As Double
On Error Resume Next
  
  Dim MyCounter As Long
  Dim StartTime As Long
  Dim EndTime As Long
  Dim TestTimes(1 To 1000) As Long
  Dim TotalSoFar As Long
  
  ' Change the screen's cursor to the Hour Glass while running the test
  Screen.MousePointer = vbHourglass
  
  ' Get the time and do a compare 1000 times to see how fast the
  ' computer can do it and return the results
  For MyCounter = 1 To 1000
    If MyCounter = MyCounter Then DoEvents
    StartTime = timeGetTime
    EndTime = timeGetTime
    TestTimes(MyCounter) = Abs(EndTime - StartTime)
    DoEvents
  Next
  
  ' Add all the times together and devide them by the total to get the average
  For MyCounter = 1 To 1000
    TotalSoFar = TotalSoFar + TestTimes(MyCounter)
  Next
  TestSmallestInterval = TotalSoFar / 1000
  If TestSmallestInterval <= 1 Then
    TestSmallestInterval = 1
  End If
  
  Screen.MousePointer = vbDefault
  
End Function

' This is the internal function that actually controls the timer events.
Private Function StartTimer()
On Error Resume Next
  
  Dim TestTime As Long
  Dim CountTime As Long
  Static StartTime As Long
  
TheStart:
  
  ' Get the start time to test for every time around
  StartTime = timeGetTime
  
  Do
    
    ' If the user has disabled the timer or changed the
    ' interval to 0 then stop the timer
    If tEnabled = False Or tInterval <= 0 Then Exit Function
    DoEvents
    
    TestTime = timeGetTime
    CountTime = Abs(TestTime - StartTime)
  Loop While CountTime < tInterval
  
  ' Once at least the timer's interval has passed, fire the
  ' Timer event and reset for the next round
  TimerEvent
  GoTo TheStart
  
End Function


' PUT CODE FOR TIMER RELATED FUNCTIONALITY HERE...
' ------------------------------------------------
' or you can use the WithEvents keyword to declare this class in a form
' and put your code in the Timer event of the variable in that form
Public Function TimerEvent()
  
  If tEnabled = False Then
    Exit Function
  End If
  
  RaiseEvent Timer
  
End Function

