VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Special Thanks goes out to the user that made this class.
'It was found at http://www.cddb.com over a year ago.
'But its not there now, good thing I found it again from a
'different site or this control would never be remade. :)
'I just don't have the time for small projects like this.
'I'm only redoing this control because of many users asking
'me to fix my first version. Well here it is, now that you
'have the source, you can fix when needed.

'NOTE TO DEVERS THAT USE THE CDDB CONTROL IN THERE APP:
'You might want to check the CDDB web site ASAP at
'http://www.cddb.com/dev/ Don't just look, read it all
'Learn whats going on. I'm almost 100% sure that this control
'will not work in time. This is why I'm giving you my source
'to this ActiveX Control. You might also want to check out
'there CDDB2 SDK.
'
'
'
'Michael L. Barker

Option Explicit

Private Type MCI_OPEN_PARMS
   dwCallback As Long
   wDeviceID As Long
   lpstrDeviceType As String
   lpstrElementName As String
   lpstrAlias As String
End Type

Private Type MCI_SET_PARMS
   dwCallback As Long
   dwTimeFormat As Long
   dwAudio As Long
End Type

Private Type MCI_STATUS_PARMS
   dwCallback As Long
   dwReturn As Long
   dwItem As Long
   dwTrack As Integer
End Type

Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" _
   (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long

Private Const MMSYSERR_NOERROR = 0

Private Const MCI_CLOSE = &H804
Private Const MCI_FORMAT_MSF = 2
Private Const MCI_OPEN = &H803
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_OPEN_TYPE = &H2000&
Private Const MCI_SET = &H80D
Private Const MCI_SET_TIME_FORMAT = &H400&

Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_STATUS_LENGTH = &H1&
Private Const MCI_STATUS_NUMBER_OF_TRACKS = &H3&
Private Const MCI_STATUS_POSITION = &H2&
Private Const MCI_TRACK = &H10&
Private Const MCI_STATUS = &H814

Private mciOpenParms As MCI_OPEN_PARMS
Private mciSetParms As MCI_SET_PARMS
Private mciStatusParms As MCI_STATUS_PARMS

Private Type TTrackInfo
   Minutes As Long
   Seconds As Long
   Frames As Long
   FrameOffset As Long           ' Starting location in frames (used by QueryString)
End Type

Private m_Error As Long          ' Error code from API call
Private m_CID As String          ' Computed disc id
Private m_Drive As String        ' Drive letter
Private m_DeviceID As Long       ' Device Id
Private m_NTracks As Integer     ' Number of tracks in CD
Private m_Length As Long         ' Length of CD in seconds
Private m_Tracks() As TTrackInfo ' Track info for each and every track on the CD
                                 ' Zero based. Last index used for storing lead-out
                                 ' position information.

Private Sub Class_Initialize()
   m_CID = "(unavailable)"
   m_Drive = ""
   m_Error = 0
   m_DeviceID = -1
   m_NTracks = 0
End Sub

Public Property Get DiscID() As String
   DiscID = m_CID
End Property

Public Property Get ErrorCode() As Long
   Error = m_Error
End Property

Public Sub Init(sDrive As String)
   Dim p1 As Integer
   m_Error = MMSYSERR_NOERROR
   m_Drive = sDrive
   If OpenCD Then
      Call LoadCDInfo
      CloseCD
   End If
   
End Sub

Private Sub Class_Terminate()
   If m_DeviceID <> -1 Then
      CloseCD
   End If
End Sub

Private Function OpenCD() As Boolean
   Dim Scode As Long, wDeviceID As Long
   OpenCD = False
   mciOpenParms.lpstrDeviceType = "cdaudio"
   mciOpenParms.lpstrElementName = m_Drive
   Scode = mciSendCommand(0, MCI_OPEN, (MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT), mciOpenParms)
   If Scode <> MMSYSERR_NOERROR Then
      m_Error = Scode
      Exit Function
   End If
   m_DeviceID = mciOpenParms.wDeviceID
   mciSetParms.dwTimeFormat = MCI_FORMAT_MSF
   Scode = mciSendCommand(m_DeviceID, MCI_SET, MCI_SET_TIME_FORMAT, mciSetParms)
   If Scode <> MMSYSERR_NOERROR Then
      m_Error = Scode
      Scode = mciSendCommand(m_DeviceID, MCI_CLOSE, 0, 0)  ' Dont forget to close it
      Exit Function
   End If
   OpenCD = True
End Function

Private Sub CloseCD()
   m_Error = mciSendCommand(m_DeviceID, MCI_CLOSE, 0, 0)
   m_DeviceID = -1
End Sub

Private Function LoadCDInfo() As Boolean
   Dim Scode As Long
   Dim p1 As Long, dwPosM As Long, dwPosS As Long, dwPosF As Long
   Dim dwLenM As Long, dwLenS As Long, dwLenF As Long, dwpos As Long
   Dim sum As Long, p2 As Long
   On Error Resume Next
   LoadCDInfo = False
   mciStatusParms.dwItem = MCI_STATUS_NUMBER_OF_TRACKS
   Scode = mciSendCommand(m_DeviceID, MCI_STATUS, MCI_STATUS_ITEM, mciStatusParms)
   If Scode <> MMSYSERR_NOERROR Then
      m_Error = Scode
      Exit Function
   End If
   m_NTracks = mciStatusParms.dwReturn
   ReDim m_Tracks(m_NTracks + 1) As TTrackInfo
   For p1 = 1 To m_NTracks
      mciStatusParms.dwItem = MCI_STATUS_POSITION
      mciStatusParms.dwTrack = p1
      Scode = mciSendCommand(m_DeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK, mciStatusParms)
      If Scode <> MMSYSERR_NOERROR Then
         m_Error = Scode
         Exit Function
      End If
      m_Tracks(p1 - 1).Frames = (mciStatusParms.dwReturn \ 65536) And &HFF
      m_Tracks(p1 - 1).Seconds = (mciStatusParms.dwReturn \ 256) And &HFF
      m_Tracks(p1 - 1).Minutes = (mciStatusParms.dwReturn) And &HFF
      m_Tracks(p1 - 1).FrameOffset = (m_Tracks(p1 - 1).Minutes * 60 * 75) + _
                                   (m_Tracks(p1 - 1).Seconds * 75) + _
                                   (m_Tracks(p1 - 1).Frames)
   Next p1
   mciStatusParms.dwItem = MCI_STATUS_LENGTH
   mciStatusParms.dwTrack = m_NTracks
   Scode = mciSendCommand(m_DeviceID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK, mciStatusParms)
   If Scode <> MMSYSERR_NOERROR Then
      m_Error = Scode
      Exit Function
   End If
   dwLenM = (mciStatusParms.dwReturn) And &HFF
   dwLenS = (mciStatusParms.dwReturn \ 256) And &HFF
   dwLenF = ((mciStatusParms.dwReturn \ 65536) And &HFF) + 1
   dwPosM = m_Tracks(m_NTracks - 1).Minutes
   dwPosS = m_Tracks(m_NTracks - 1).Seconds
   dwPosF = m_Tracks(m_NTracks - 1).Frames
   dwpos = (dwPosM * 60 * 75) + (dwPosS * 75) + dwPosF + _
           (dwLenM * 60 * 75) + (dwLenS * 75) + dwLenF
   m_Tracks(m_NTracks).Frames = dwpos Mod 75
   dwpos = dwpos \ 75
   m_Tracks(m_NTracks).Seconds = dwpos Mod 60
   dwpos = dwpos \ 60
   m_Tracks(m_NTracks).Minutes = dwpos
   m_Length = ((m_Tracks(m_NTracks).Minutes * 60) + (m_Tracks(m_NTracks).Seconds)) - _
              ((m_Tracks(0).Minutes * 60) + (m_Tracks(0).Seconds))
   sum = 0
   For p1 = 0 To m_NTracks - 1
      p2 = m_Tracks(p1).Minutes * 60 + m_Tracks(p1).Seconds
      Do While p2 > 0
         sum = sum + (p2 Mod 10)
         p2 = p2 \ 10
      Loop
   Next p1
   m_CID = LCase$(LeftZeroPad(Hex$(sum Mod &HFF), 2) & LeftZeroPad(Hex$(m_Length), 4) & LeftZeroPad(Hex$(m_NTracks), 2))
   LoadCDInfo = True
End Function

Public Function QueryString() As String
    Dim p1 As Integer, s As String
    On Error GoTo CHK
    s = "cddb query " & m_CID & " " & m_NTracks
    For p1 = 0 To m_NTracks - 1
        s = s & " " & Format$(m_Tracks(p1).FrameOffset)
    Next
    QueryString = s & " " & Format$(m_Tracks(m_NTracks).Minutes * 60) + (m_Tracks(m_NTracks).Seconds)
CHK:
    Select Case Err.Number
    Case 0
    Case 9
        QueryString = "Drive not ready. Try again."
        Exit Function
    Case Else
        MsgBox Err.Number & " " & Err.Description
        Exit Function
    End Select
End Function

Private Function LeftZeroPad(s As String, n As Integer) As String
   If Len(s) < n Then
      LeftZeroPad = String$(n - Len(s), "0") & s
   Else
      LeftZeroPad = s
   End If
End Function
