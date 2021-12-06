Attribute VB_Name = "Module1"

'API for playing Audio CD
Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long
    
'API for getting all drives on your PC
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
       (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Api to determine if you have a CD, floppy or Hard drive
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
       (ByVal nDrive As String) As Long
       
       
Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
       
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2


Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2

Dim CurrentVolLeft As Long
Dim CurrentVolRight As Long

Dim kkk As String

     

