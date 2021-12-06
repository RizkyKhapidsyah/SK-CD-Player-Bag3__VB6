VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generic Multiple CD Player"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   960
   End
   Begin VB.CommandButton Command8 
      Caption         =   "About..."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Close"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Eject"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CD 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pause"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Track>>"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<Track"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Available Audio CD Drives"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentCd As String
Dim mssg As String * 255
Public Sub Detect_CDs()

Dim SmallString As String
Dim NextDrive As String
Static z As Integer
       
alldrives$ = Space$(64)
'Get all drives on your PC as one long string
ret& = GetLogicalDriveStrings(Len(alldrives$), alldrives$)
'trim off any trailing spaces. AllDrives$
'now contains all the drive letters.
alldrives$ = Left$(alldrives$, ret&)


' "AllDrives$"  contains a string of all of your drives
'in your computer, but there is a character "chr$(0)"
'between each drive letter that we must filter out.
'We will use the "FOR NEXT" function to do this.
   
For k = 1 To Len(alldrives$)
  SmallString = Mid$(alldrives$, k, 1) 'Get one character at a time
  If SmallString = Chr$(0) Then
           SmallString = ""     'remove unwanted character
           DriveType& = GetDriveType(NextDrive) 'Check if it is a CD drive
           If DriveType = 5 Then
              If CD(0).Caption = "" Then 'Our first button needs to be updated before the others.
                CD(0).Caption = UCase$(NextDrive)
                CD(z).Visible = True
                CurrentCd = UCase$(NextDrive)
              Else
                'Since this is a CD drive, make a button for it.
                'This code below creates command buttons
                 z = z + 1
                 Load CD(z)
                 CD(z).Caption = UCase$(NextDrive)
                 CD(z).Left = (CD(z - 1).Left) + (CD(z - 1).Width)
                 CD(z).Visible = True
              End If
           End If
           NextDrive = "" 'Now that a drive was detected, clear the
                          'string for new info
    End If
      
NextDrive = NextDrive & SmallString
   
Next k

If CD(0).Caption = "" Then
  MsgBox "No Audio CDs were detected", vbInformation, ""
  End
Else
UpDate_Cds
End If

End Sub

Private Sub CD_Click(Index As Integer)
  i = mciSendString("stop cd", 0&, 0, 0)
  i = mciSendString("close cd", 0&, 0, 0)
  CurrentCd = CD(Index).Caption

UpDate_Cds
End Sub


Private Sub Command1_Click()
Dim numTracks As Integer
Dim CurTrack As Integer


'Get the current track
rc = mciSendString("status cd current track", mssg, 255, 0)
CurTrack = Str(mssg)

'Get total number of tracks
rc = mciSendString("status cd number of tracks wait", mssg, 255, 0)
numTracks = Str(mssg)

'Check to see if CD is playing
rc = mciSendString("status cd mode", mssg, 255, 0)

If Left$(mssg, 7) = "playing" Then
    If CurTrack = 1 Then
         rc = mciSendString("play cd from " & numTracks, mssg, 255, 0)
    Else
         rc = mciSendString("play cd from " & CurTrack - 1, mssg, 255, 0)
    End If
Else
    If CurTrack = 1 Then
         rc = mciSendString("seek cd to " & numTracks, mssg, 255, 0)
    Else
         rc = mciSendString("seek cd to " & CurTrack - 1, mssg, 255, 0)
    End If
End If
End Sub

Private Sub Command2_Click()
  i = mciSendString("play cd", 0&, 0, 0)
End Sub

Private Sub Command3_Click()
Dim mssg As String * 255
Dim numTracks As Integer
Dim CurTrack As Integer


'Get the current track
rc = mciSendString("status cd current track", mssg, 255, 0)
CurTrack = Str(mssg)

'Get total number of tracks
rc = mciSendString("status cd number of tracks wait", mssg, 255, 0)
numTracks = Str(mssg)

'Check to see if CD is playing
rc = mciSendString("status cd mode", mssg, 255, 0)

If Left$(mssg, 7) = "playing" Then
    If CurTrack = numTracks Then
         rc = mciSendString("play cd from 1", mssg, 255, 0)
    Else
         rc = mciSendString("play cd from " & CurTrack + 1, mssg, 255, 0)
    End If
Else
    If CurTrack = numTracks Then
         rc = mciSendString("seek cd to 1", mssg, 255, 0)
    Else
         rc = mciSendString("seek cd to " & CurTrack + 1, mssg, 255, 0)
    End If
End If
End Sub

Private Sub Command4_Click()
  i = mciSendString("pause cd wait", 0&, 0, 0)
End Sub

Private Sub Command5_Click()
  i = mciSendString("stop cd wait", 0&, 0, 0)
  i = mciSendString("seek cd to 1 wait", 0&, 0, 0)
End Sub


Private Sub Command6_Click()

i = mciSendString("set cd door open wait", mssg, 255, 0)

End Sub

Private Sub Command7_Click()
i = mciSendString("status cd mode", mssg, 255, 0)

If Left$(mssg, 4) = "open" Then
   i = mciSendString("set cd door closed wait", mssg, 255, 0)
End If

End Sub


Private Sub Command8_Click()
'Show form2 (About box) and disable form1
Form2.Show 1
End Sub


Private Sub Form_Load()
' If we're already running, then quit
If (App.PrevInstance = True) Then
    End
End If


Detect_CDs
End Sub


Private Sub Form_Unload(Cancel As Integer)

  i = mciSendString("stop cd", 0&, 0, 0)
  i = mciSendString("close cd", 0&, 0, 0)
  i = mciSendString("close all", 0&, 0, 0)
End Sub



Public Sub UpDate_Cds()
  i = mciSendString("open  " & CurrentCd & " type cdaudio alias cd wait shareable", 0&, 0, 0)
  i = mciSendString("set cd time format tmsf", 0&, 0, 0)


End Sub

Private Sub Timer1_Timer()

' Check if CD is in the player
i = mciSendString("status cd media present", mssg, 255, 0)

If Left$(mssg, 4) = "true" Then
   i = mciSendString("status cd position", mssg, 255, 0)
   track = CInt(Mid$(mssg, 1, 2))
   Min = CInt(Mid$(mssg, 4, 2))
   sec = CInt(Mid$(mssg, 7, 2))
   Label3.Caption = "[" & Format(track, "00") & "] " & Format(Min, "00") & ":" & Format(sec, "00")
Else
   
End If

End Sub


