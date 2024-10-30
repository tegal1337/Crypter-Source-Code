Attribute VB_Name = "mSound"
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Public Const SPI_GETWORKAREA = 48


Public Type RECT
    PIzq As Long
    PDer As Long
    PArr As Long
    PAbj As Long
    End Type
    ' *********************************************************************
'  Copyright ©1996-2001 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

'  flag values for uFlags parameter
Private Const SND_SYNC = &H0              '  play synchronously (default)
Private Const SND_ASYNC = &H1             '  play asynchronously
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_MEMORY = &H4            '  lpszSoundName points to a memory file
Private Const SND_ALIAS = &H10000         '  name is a WIN.INI [sounds] entry
Private Const SND_FILENAME = &H20000      '  name is a file name
Private Const SND_RESOURCE = &H40004      '  name is a resource name or atom
Private Const SND_ALIAS_ID = &H110000     '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS_START = 0         '  must be > 4096 to keep strings in same section of resource file
Private Const SND_LOOP = &H8              '  loop the sound until next sndPlaySound
Private Const SND_NOSTOP = &H10           '  don't stop any currently playing sound
Private Const SND_VALID = &H1F            '  valid flags          / ;Internal /
Private Const SND_NOWAIT = &H2000         '  don't wait if the driver is busy
Private Const SND_VALIDFLAGS = &H17201F   '  Set of valid flag bits.  Anything outside this range will raise an error
Private Const SND_RESERVED = &HFF000000   '  In particular these flags are reserved
Private Const SND_TYPE_MASK = &H170007

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'  waveform audio error return values
Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)       '  unsupported wave format
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)      '  header not prepared
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)            '  device is synchronous
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)       '  last error in range

Private m_snd() As Byte

Public Function PlaySoundResource1(ByVal SndID As Long) As Long
   Const Flags = SND_RESOURCE Or SND_ASYNC Or SND_NODEFAULT
   PlaySoundResource1 = PlaySound(CStr("#" & SndID), App.hInstance, Flags)
End Function

Public Function PlaySoundResource(ByVal SndID As Long) As Long
   Const Flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
   m_snd = LoadResData(SndID, "CUSTOM")
   PlaySoundResource = PlaySoundData(m_snd(0), 0, Flags)
End Function

