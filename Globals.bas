Attribute VB_Name = "Module2"
Option Explicit

'NOTE not all these variables are actually used
'     it would take a while to go thru and see
'     which are and arent and i`m in a hurry! :P
'
Public OrigFile(1 To 50000) As String * 2
Public NewFile(1 To 50000) As String * 2

Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
        'startinfo As STARTUPINFO
End Type
Public startinfo As STARTUPINFO


Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
        'procinfo As PROCESS_INFORMATION
End Type
Public procinfo As PROCESS_INFORMATION

Global X As Integer
Global TFile As String
Global Const Hite = 4900
Global CardInserted As Integer

Global BytesToRead As Integer
Global CardInfoBuffer As String

Global Const ResetBaud = 9600
Global Const DataBaud = 19200


Global UserTimeZone
Global UserZipCode


Global AtrLen
Global CheckINS
Global BytesTotalSent As Long

Global ByteToFlip(1 To 10000) As String * 2
Global SendStr(1 To 500) As String * 2
Global PPV(1 To 25) As String
Global R0byte As String
Global BufferIncount As Integer

Global OldCount As String
Global Current As String
Global CardDone As Long

Global TmpStr As String
Global ByteStr As String

Global InverseBuffer As String
Global BufferIn As String
Global WorkByte As String

Global ATR As String
Global PreATR As String
Global PostATR As String

Global DATA As String
Global preDATA As String
Global postDATA As String

Public InBuff As String * 1
Public port As String
Public icond As Boolean

Global tmpFUSE, tmpRATING
Global tmpCARDID, tmpIRD
Global tmpSPENDING, tmpTIMEZONE
Global tmpGUIDE, tmpTRASH
Global tmpTEMP, tmpUSW

Global xxx As Integer

Public temp1
Public Temp2
Public Temp3
Public Temp4
Public Cpos As Integer
Public Npos As Integer
Public Byte_Value
Public byteLen As Integer
Public Nibble As Integer
Public Term_Count As Long

Type dcbType
    DCBlength As Long
    BaudRate As Long
    Bits1 As Long
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer
End Type

Public Const ERR_INVALIDPROPERTY = 31000
Public Const fBinary = &H1&
Public Const fParity = &H2&
Public Const fOutxCtsFlow = &H4&
Public Const fOutxDsrFlow = &H8&
Public Const fDtrControl = &H30&
Public Const fDsrSensitivity = &H40&
Public Const fTXContinueOnXoff = &H80&
Public Const fOutX = &H100&
Public Const fInX = &H200&
Public Const fErrorChar = &H400&
Public Const fNull = &H800&
Public Const fRtsControl = &H3000&
Public Const fAbortOnError = &H4000&

Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public timeouts As COMMTIMEOUTS
Public hPort As Long
Public dwCommModemStatus As Long
Public numRead As Long
Public DCB As dcbType
Public written As Long

Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Boolean
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Boolean
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Boolean
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As dcbType) As Long
Public Declare Function GetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As dcbType) As Long
Public Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Public Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Public Declare Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Public Declare Function EscapeCommFunction Lib "kernel32" (ByVal hFile As Long, ByVal dwFunc As Long) As Boolean
Public Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Long, lpModemStat As Long) As Boolean
Public Declare Function GetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Boolean
Public Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Public Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, ByVal l As Long) As Long
Public Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As Long) As Long
Public Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Public Const PURGE_RXCLEAR = &H8
Public Const PURGE_TXCLEAR = &H4
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const INVALID_HANDLE_VALUE = -1
Public Const GMEM_FIXED = &H0
Public Const ERROR_IO_PENDING = 997
Public Const WAIT_TIMEOUT = &H102&
Public Const MS_CTS_ON = &H10&
Public Const MS_DSR_ON = &H20&
Public Const MS_RING_ON = &H40&
Public Const MS_RLSD_ON = &H80&
Public Const ERR_NOCOMMACCESS = 31010
Public Const ERR_UNINITIALIZED = 31011
Public Const ERR_MODEMSTATUS = 31012
Public Const ERR_READFAIL = 31013
Public Const ERR_EVENTFAIL = 31014
Public Const EV_RXCHAR = &H1                '  Any Character received
Public Const EV_RXFLAG = &H2                '  Received certain character
Public Const EV_TXEMPTY = &H4               '  Transmitt Queue Empty
Public Const EV_CTS = &H8                   '  CTS changed state
Public Const EV_DSR = &H10                  '  DSR changed state
Public Const EV_RLSD = &H20                 '  RLSD changed state
Public Const EV_BREAK = &H40                '  BREAK received
Public Const EV_ERR = &H80                  '  Line status error occurred
Public Const EV_RING = &H100                '  Ring signal detected
Public Const EV_PERR = &H200                '  Printer error occured
Public Const EV_RX80FULL = &H400            '  Receive BufferIn is 80 percent full
Public Const EV_EVENT1 = &H800              '  Provider specific event 1
Public Const EV_EVENT2 = &H1000             '  Provider specific event 2
Public Const CE_RXOVER = &H1                '  Receive Queue overflow
Public Const CE_OVERRUN = &H2               '  Receive Overrun Error
Public Const CE_RXPARITY = &H4              '  Receive Parity Error
Public Const CE_FRAME = &H8                 '  Receive Framing error
Public Const CE_BREAK = &H10                '  Break Detected
Public Const CE_TXFULL = &H100              '  TX Queue is full
Public Const NOPARITY = 0
Public Const ODDPARITY = 1
Public Const EVENPARITY = 2
Public Const MARKPARITY = 3
Public Const SPACEPARITY = 4
Public Const ONESTOPBIT = 0
Public Const ONE5STOPBITS = 1
Public Const TWOSTOPBITS = 2
Public Const IGNORE = 0
Public Const INFINITE = &HFFFF
Public Const CE_CTSTO = &H20
Public Const CE_DSRTO = &H40
Public Const CE_RLSDTO = &H80
Public Const CE_PTO = &H200
Public Const CE_IOE = &H400
Public Const CE_DNS = &H800
Public Const CE_OOP = &H1000
Public Const CE_MODE = &H8000
Public Const IE_BADID = (-1)
Public Const IE_OPEN = (-2)
Public Const IE_NOPEN = (-3)
Public Const IE_MEMORY = (-4)
Public Const IE_DEFAULT = (-5)
Public Const IE_HARDWARE = (-10)
Public Const IE_BYTESIZE = (-11)
Public Const IE_BAUDRATE = (-12)
Public Const EV_CTSS = &H400
Public Const EV_DSRS = &H800
Public Const EV_RLSDS = &H1000
Public Const SETXOFF = 1
Public Const SETXON = 2
Public Const SETRTS = 3
Public Const CLRRTS = 4
Public Const SETDTR = 5
Public Const CLRDTR = 6
Public Const RESETDEV = 7
Public Const GETMAXLPT = 8
Public Const GETMAXCOM = 9
Public Const GETBASEIRQ = 10
Public Const CBR_110 = &HFF10
Public Const CBR_300 = &HFF11
Public Const CBR_600 = &HFF12
Public Const CBR_1200 = &HFF13
Public Const CBR_2400 = &HFF14
Public Const CBR_4800 = &HFF15
Public Const CBR_9600 = &HFF16
Public Const CBR_14400 = &HFF17
Public Const CBR_19200 = &HFF18
Public Const CBR_38400 = &HFF1B
Public Const CBR_56000 = &HFF1F
Public Const CBR_128000 = &HFF23
Public Const CBR_256000 = &HFF27
Public Const CN_RECEIVE = &H1
Public Const CN_TRANSMIT = &H2
Public Const CN_EVENT = &H4
Public Const CSTF_CTSHOLD = &H1
Public Const CSTF_DSRHOLD = &H2
Public Const CSTF_RLSDHOLD = &H4
Public Const CSTF_XOFFHOLD = &H8
Public Const CSTF_XOFFSENT = &H10
Public Const CSTF_EOF = &H20
Public Const CSTF_TXIM = &H40
Public Const LPTx = &H80
Public Const DTR_CONTROL_DISABLE = &H0
Public Const DTR_CONTROL_ENABLE = &H1
Public Const DTR_CONTROL_HANDSHAKE = &H2
Public Const RTS_CONTROL_DISABLE = &H0
Public Const RTS_CONTROL_ENABLE = &H1
Public Const RTS_CONTROL_HANDSHAKE = &H2
Public Const RTS_CONTROL_TOGGLE = &H3

Public Const CardInfoStr = ("482A000080")
Public Const PPVinfoStr = ("485E080E95")
Public Const IRDinfoStr = ("4858000017")
Public Const IRDpacket = ("4852000004")

Global titleA$
Global titleB$
Global titleC$
Global titleD$
Global titleE$

