VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "                                                           PRIVATE VERSION"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   3750
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Form3
Form1.Show
End Sub
