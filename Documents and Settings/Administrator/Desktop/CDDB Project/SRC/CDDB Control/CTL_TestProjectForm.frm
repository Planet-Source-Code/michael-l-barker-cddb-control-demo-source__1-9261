VERSION 5.00
Object = "*\A..\..\..\..\..\..\DOCUME~1\ADMINI~1\Desktop\CDDBPR~2\SRC\CDDBCO~1\CDDBControl.vbp"
Begin VB.Form CTL_TestProjectForm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "All Tracks"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Track #6"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin CDDBControl.CDDB CDDB1 
      Left            =   780
      Top             =   1260
      _ExtentX        =   1535
      _ExtentY        =   1058
   End
End
Attribute VB_Name = "CTL_TestProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CDDB1_AllServerMessages(Text As String)
    Debug.Print Text
End Sub


Private Sub CDDB1_CDInfo(AlbumName As String, ArtistName As String, TotalTracks As Long)
    MsgBox "Album Name: '" & AlbumName & "'" & vbCrLf & "ArtistName: '" & ArtistName & "'" & vbCrLf & "TotalTracks: '" & TotalTracks & "'"
End Sub


Private Sub CDDB1_Connected()
    MsgBox "Connected"
End Sub


Private Sub CDDB1_Disconnected()
    MsgBox "Disconnected"
    Command2.Enabled = True
    Command3.Enabled = True
End Sub


Private Sub Command1_Click()
'CDDB1.DriveLetter = "d"
CDDB1.Connect Random_US_site
End Sub


Private Sub Command2_Click()

MsgBox CDDB1.GetTrackName(5)

End Sub



Private Sub Command3_Click()
Dim a As Long

Debug.Print "": Debug.Print "": Debug.Print "":

Debug.Print "Track Number" & vbTab & "Track Name"
Debug.Print String(26, "-")

For a = 0 To CDDB1.GetTrackCount
    Debug.Print a + 1 & vbTab & vbTab & vbTab & vbTab & CDDB1.GetTrackName(a)
Next

Debug.Print "": Debug.Print "": Debug.Print "":

End Sub


