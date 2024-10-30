Attribute VB_Name = "Module1"
Option Explicit
 
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const CREATE_NEW = 1
Public Const INVALID_HANDLE_VALUE = -1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const IDENTIFY_BUFFER_SIZE = 512
Public Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16

'GETVERSIONOUTPARAMS contains the data returned
'from the Get Driver Version function
Public Type GETVERSIONOUTPARAMS
   bVersion       As Byte 'Binary driver version.
   bRevision      As Byte 'Binary driver revision
   bReserved      As Byte 'Not used
   bIDEDeviceMap  As Byte 'Bit map of IDE devices
   fCapabilities  As Long 'Bit mask of driver capabilities
   dwReserved(3)  As Long 'For future use
End Type

'IDE registers
Public Type IDEREGS
   bFeaturesReg     As Byte 'Used for specifying SMART "commands"
   bSectorCountReg  As Byte 'IDE sector count register
   bSectorNumberReg As Byte 'IDE sector number register
   bCylLowReg       As Byte 'IDE low order cylinder value
   bCylHighReg      As Byte 'IDE high order cylinder value
   bDriveHeadReg    As Byte 'IDE drive/head register
   bCommandReg      As Byte 'Actual IDE command
   bReserved        As Byte 'reserved for future use - must be zero
End Type

'SENDCMDINPARAMS contains the input parameters for the
'Send Command to Drive function
Public Type SENDCMDINPARAMS
   cBufferSize     As Long     'Buffer size in bytes
   irDriveRegs     As IDEREGS  'Structure with drive register values.
   bDriveNumber    As Byte     'Physical drive number to send command to (0,1,2,3).
   bReserved(2)    As Byte     'Bytes reserved
   dwReserved(3)   As Long     'DWORDS reserved
   bBuffer()      As Byte      'Input buffer.
End Type

'Valid values for the bCommandReg member of IDEREGS.
Public Const IDE_ID_FUNCTION = &HEC            'Returns ID sector for ATA.
Public Const IDE_EXECUTE_SMART_FUNCTION = &HB0 'Performs SMART cmd.
                                                'Requires valid bFeaturesReg,
                                                'bCylLowReg, and bCylHighReg

'Cylinder register values required when issuing SMART command
Public Const SMART_CYL_LOW = &H4F
Public Const SMART_CYL_HI = &HC2

'Status returned from driver
Public Type DRIVERSTATUS
   bDriverError  As Byte          'Error code from driver, or 0 if no error
   bIDEStatus    As Byte          'Contents of IDE Error register
                                  'Only valid when bDriverError is SMART_IDE_ERROR
   bReserved(1)  As Byte
   dwReserved(1) As Long
 End Type

Public Type IDSECTOR
   wGenConfig                 As Integer
   wNumCyls                   As Integer
   wReserved                  As Integer
   wNumHeads                  As Integer
   wBytesPerTrack             As Integer
   wBytesPerSector            As Integer
   wSectorsPerTrack           As Integer
   wVendorUnique(2)           As Integer
   sSerialNumber(19)          As Byte
   wBufferType                As Integer
   wBufferSize                As Integer
   wECCSize                   As Integer
   sFirmwareRev(7)            As Byte
   sModelNumber(39)           As Byte
   wMoreVendorUnique          As Integer
   wDoubleWordIO              As Integer
   wCapabilities              As Integer
   wReserved1                 As Integer
   wPIOTiming                 As Integer
   wDMATiming                 As Integer
   wBS                        As Integer
   wNumCurrentCyls            As Integer
   wNumCurrentHeads           As Integer
   wNumCurrentSectorsPerTrack As Integer
   ulCurrentSectorCapacity    As Long
   wMultSectorStuff           As Integer
   ulTotalAddressableSectors  As Long
   wSingleWordDMA             As Integer
   wMultiWordDMA              As Integer
   bReserved(127)             As Byte
End Type

'Structure returned by SMART IOCTL commands
Public Type SENDCMDOUTPARAMS
  cBufferSize   As Long         'Size of Buffer in bytes
  DRIVERSTATUS  As DRIVERSTATUS 'Driver status structure
  bBuffer()    As Byte          'Buffer of arbitrary length for data read from drive
End Type

'Vendor specific feature register defines
'for SMART "sub commands"
Public Const SMART_ENABLE_SMART_OPERATIONS = &HD8

'Status Flags Values
Public Enum STATUS_FLAGS
   PRE_FAILURE_WARRANTY = &H1
   ON_LINE_COLLECTION = &H2
   PERFORMANCE_ATTRIBUTE = &H4
   ERROR_RATE_ATTRIBUTE = &H8
   EVENT_COUNT_ATTRIBUTE = &H10
   SELF_PRESERVING_ATTRIBUTE = &H20
End Enum

'IOCTL commands
Public Const DFP_GET_VERSION = &H74080
Public Const DFP_SEND_DRIVE_COMMAND = &H7C084
Public Const DFP_RECEIVE_DRIVE_DATA = &H7C088

Public Type ATTR_DATA
   AttrID As Byte
   AttrName As String
   AttrValue As Byte
   ThresholdValue As Byte
   WorstValue As Byte
   StatusFlags As STATUS_FLAGS
End Type

Public Type DRIVE_INFO
   bDriveType As Byte
   SerialNumber As String
   Model As String
   FirmWare As String
   Cilinders As Long
   Heads As Long
   SecPerTrack As Long
   BytesPerSector As Long
   BytesperTrack As Long
   NumAttributes As Byte
   Attributes() As ATTR_DATA
End Type

Public Enum IDE_DRIVE_NUMBER
   PRIMARY_MASTER
   PRIMARY_SLAVE
   SECONDARY_MASTER
   SECONDARY_SLAVE
   TERTIARY_MASTER
   TERTIARY_SLAVE
   QUARTIARY_MASTER
   QUARTIARY_SLAVE
End Enum

Public Declare Function CreateFile Lib "kernel32" _
   Alias "CreateFileA" _
  (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   lpSecurityAttributes As Any, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long
  
Public Declare Function DeviceIoControl Lib "kernel32" _
  (ByVal hDevice As Long, _
   ByVal dwIoControlCode As Long, _
   lpInBuffer As Any, _
   ByVal nInBufferSize As Long, _
   lpOutBuffer As Any, _
   ByVal nOutBufferSize As Long, _
   lpBytesReturned As Long, _
   lpOverlapped As Any) As Long
  
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (hpvDest As Any, _
   hpvSource As Any, _
   ByVal cbCopy As Long)
  
Public Type OSVERSIONINFO
   OSVSize As Long
   dwVerMajor As Long
   dwVerMinor As Long
   dwBuildNumber As Long
   PlatformID As Long
   szCSDVersion As String * 128
End Type

Public Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (LpVersionInformation As OSVERSIONINFO) As Long


Public Function GetDriveInfo(drvNumber As IDE_DRIVE_NUMBER) As DRIVE_INFO
    
   Dim hDrive As Long
   Dim di As DRIVE_INFO
   
   hDrive = SmartOpen(drvNumber)
   
   If hDrive <> INVALID_HANDLE_VALUE Then
   
      If SmartGetVersion(hDrive) = True Then
      
         With di
            .bDriveType = 0
            .NumAttributes = 0
            ReDim .Attributes(0)
            .bDriveType = 1
         End With
         
         If SmartCheckEnabled(hDrive, drvNumber) Then
            
            If IdentifyDrive(hDrive, IDE_ID_FUNCTION, drvNumber, di) = True Then
         
               GetDriveInfo = di
               
            End If   'IdentifyDrive
         End If   'SmartCheckEnabled
      End If   'SmartGetVersion
   End If   'hDrive <> INVALID_HANDLE_VALUE
   
   CloseHandle hDrive
   
End Function


Public Function IdentifyDrive(ByVal hDrive As Long, _
                               ByVal IDCmd As Byte, _
                               ByVal drvNumber As IDE_DRIVE_NUMBER, _
                               di As DRIVE_INFO) As Boolean
    
  'Function: Send an IDENTIFY command to the drive
  'drvNumber = 0-3
  'IDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
   Dim SCIP As SENDCMDINPARAMS
   Dim IDSEC As IDSECTOR
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim cbBytesReturned As Long
   
   With SCIP
      .cBufferSize = IDENTIFY_BUFFER_SIZE
      .bDriveNumber = CByte(drvNumber)
        
      With .irDriveRegs
         .bFeaturesReg = 0
         .bSectorCountReg = 1
         .bSectorNumberReg = 1
         .bCylLowReg = 0
         .bCylHighReg = 0
         .bDriveHeadReg = &HA0 'compute the drive number
         If Not IsWinNT4Plus Then
            .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
         End If
         'the command can either be IDE
         'identify or ATAPI identify.
         .bCommandReg = CByte(IDCmd)
      End With
   End With
   
   If DeviceIoControl(hDrive, _
                      DFP_RECEIVE_DRIVE_DATA, _
                      SCIP, _
                      Len(SCIP) - 4, _
                      bArrOut(0), _
                      OUTPUT_DATA_SIZE, _
                      cbBytesReturned, _
                      ByVal 0&) Then
                      
      CopyMemory IDSEC, bArrOut(16), Len(IDSEC)

      di.Model = StrConv(SwapBytes(IDSEC.sModelNumber), vbUnicode)
      di.SerialNumber = StrConv(SwapBytes(IDSEC.sSerialNumber), vbUnicode)
      
      IdentifyDrive = True
      
    End If
    
End Function


Public Function IsWinNT4Plus() As Boolean

  'returns True if running Windows NT4 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)

   If GetVersionEx(osv) = 1 Then
   
      IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                     (osv.dwVerMajor >= 4)
 
   End If

End Function


Public Function SmartCheckEnabled(ByVal hDrive As Long, _
                                   drvNumber As IDE_DRIVE_NUMBER) As Boolean
   
  'SmartCheckEnabled - Check if SMART enable
  'FUNCTION: Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
  'bDriveNum = 0-3
   Dim SCIP As SENDCMDINPARAMS
   Dim SCOP As SENDCMDOUTPARAMS
   Dim cbBytesReturned As Long
   
   With SCIP
   
      .cBufferSize = 0
      
      With .irDriveRegs
           .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
           .bSectorCountReg = 1
           .bSectorNumberReg = 1
           .bCylLowReg = SMART_CYL_LOW
           .bCylHighReg = SMART_CYL_HI

           .bDriveHeadReg = &HA0
            If Not IsWinNT4Plus Then
               .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
            End If
           .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
           
       End With
       
       .bDriveNumber = drvNumber
       
   End With
   
   SmartCheckEnabled = DeviceIoControl(hDrive, _
                                      DFP_SEND_DRIVE_COMMAND, _
                                      SCIP, _
                                      Len(SCIP) - 4, _
                                      SCOP, _
                                      Len(SCOP) - 4, _
                                      cbBytesReturned, _
                                      ByVal 0&)
End Function


Public Function SmartGetVersion(ByVal hDrive As Long) As Boolean
   
   Dim cbBytesReturned As Long
   Dim GVOP As GETVERSIONOUTPARAMS
   
   SmartGetVersion = DeviceIoControl(hDrive, _
                                     DFP_GET_VERSION, _
                                     ByVal 0&, 0, _
                                     GVOP, _
                                     Len(GVOP), _
                                     cbBytesReturned, _
                                     ByVal 0&)
   
End Function


Public Function SmartOpen(drvNumber As IDE_DRIVE_NUMBER) As Long

  'Open SMART to allow DeviceIoControl
  'communications and return SMART handle

   If IsWinNT4Plus() Then
      
      SmartOpen = CreateFile("\\.\PhysicalDrive" & CStr(drvNumber), _
                             GENERIC_READ Or GENERIC_WRITE, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                             ByVal 0&, _
                             OPEN_EXISTING, _
                             0&, _
                             0&)

   Else
      
      SmartOpen = CreateFile("\\.\SMARTVSD", _
                              0&, 0&, _
                              ByVal 0&, _
                              CREATE_NEW, _
                              0&, _
                              0&)
   End If
   
End Function


Public Function SwapBytes(b() As Byte) As Byte()
   
  'Note: VB4-32 and VB5 do not support the
  'return of arrays from a function. For
  'developers using these VB versions there
  'are two workarounds to this restriction:
  '
  '1) Change the return data type ( As Byte() )
  '   to As Variant (no brackets). No change
  '   to the calling code is required.
  '
  '2) Change the function to a sub, remove
  '   the last line of code (SwapBytes = b()),
  '   and take advantage of the fact the
  '   original byte array is being passed
  '   to the function ByRef, therefore any
  '   changes made to the passed data are
  '   actually being made to the original data.
  '   With this workaround the calling code
  '   also requires modification:
  '
  '      di.Model = StrConv(SwapBytes(IDSEC.sModelNumber), vbUnicode)
  '
  '   ... to ...
  '
  '      Call SwapBytes(IDSEC.sModelNumber)
  '      di.Model = StrConv(IDSEC.sModelNumber, vbUnicode)
   
   Dim bTemp As Byte
   Dim cnt As Long

   For cnt = LBound(b) To UBound(b) Step 2
      bTemp = b(cnt)
      b(cnt) = b(cnt + 1)
      b(cnt + 1) = bTemp
   Next cnt
      
   SwapBytes = b()
      
End Function


