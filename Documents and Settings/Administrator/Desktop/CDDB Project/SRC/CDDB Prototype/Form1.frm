VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "CDDB Prototype"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CD Rom 2 [F:]"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   270
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3165
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CD-Rom 1 [D:]"
      Height          =   375
      Left            =   3465
      TabIndex        =   0
      Top             =   270
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   158
      TabIndex        =   3
      Top             =   930
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   158
      TabIndex        =   2
      Top             =   510
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   158
      TabIndex        =   1
      Top             =   150
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TrackCount As Long

Dim QueryString As String

Dim DiskID As String
Dim DiskCategory As String

Dim AlbumName As String
Dim ArtistName As String

Dim TrackTitles() As String  'Only goes to 10 tracks, but we redim it later

Dim TempStringArray$()
Dim TempStringArray2$()
Dim TempString As String
Private Sub Command1_Click()
Dim CDClass As CCd
Set CDClass = New CCd

CDClass.Init "D:"

QueryString = CDClass.QueryString
DiskID = CDClass.DiscID

Winsock1.LocalPort = 8880
Winsock1.RemotePort = 8880
Winsock1.Connect "us.cddb.com", 8880

Set CDClass = Nothing

End Sub


Private Sub Command2_Click()

Dim CDClass As CCd
Set CDClass = New CCd

CDClass.Init "E:"

QueryString = CDClass.QueryString
DiskID = CDClass.DiscID

Winsock1.LocalPort = 8880
Winsock1.RemotePort = 8880
Winsock1.Connect "us.cddb.com", 8880

Set CDClass = Nothing

End Sub


Private Sub Form_Load()
    Label1.BackColor = &H8000000F
    Label2.BackColor = &H8000000F
    Label3.BackColor = &H8000000F
    
    Label3.Caption = "Stick a CD in the drive, then press a CD Rom button."
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub


Private Sub Form_Unload(Cancel As Integer)

End

End Sub


Private Sub Winsock1_Close()
Dim a As Long

Label1.Caption = ArtistName
Label2.Caption = AlbumName

Label3.Caption = ""

For a = 0 To UBound(TrackTitles)
    'This is a poor way of doing this, but its just a prototype.
    Label3.Caption = Label3.Caption & a + 1 & ") " & TrackTitles(a) & vbCrLf
Next

Winsock1.Close
    
End Sub

Private Sub Winsock1_Connect()

Winsock1.SendData "cddb hello Michael michaelb Test_APP 1.0.0" & vbCrLf

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Text As String
Dim a As Long

Winsock1.GetData Text, vbString

'Debug Start
    Debug.Print Text
    'TODO: When builing control, make Event AllText and raise it here
'Debug End

If InStr(1, Text, "200 Hello and welcome ", vbTextCompare) Then
    'User Logged In
    Winsock1.SendData QueryString & vbCrLf
    Debug.Print "** User Logged In **"
End If

'Is there more then two ways to get the CD Info? :)

'CD Found #1
If InStr(1, Text, "200 ", vbTextCompare) Then
    If InStr(1, Text, " / ", vbTextCompare) Then
        
        'Get Category
        TempString = InStr(4, Text, " ", vbTextCompare) + 1 '1st Space
        DiskCategory = Mid(Text, TempString, InStr(TempString, Text, " ", vbTextCompare) - TempString)
        
            
        'We get the disk ID, There disk ID is differen't then
        'the one we got before we started. Why? I have no clue.
        'You're going to have to find that one out on your own.
        
        'Get the new DISK ID
        TempString = InStr(5, Text, " ", vbTextCompare) + 1 '1st Space
        DiskID = Mid(Text, TempString, InStr(TempString, Text, " ", vbTextCompare) - TempString)
        
        Winsock1.SendData "cddb read " & DiskCategory & " " & DiskID & vbCrLf
        
    End If
End If


'CD Found #2
If InStr(1, Text, "211 Found inexact matches, list follows (until terminating `.')", vbTextCompare) Then
    TempStringArray = Split(Text, vbCrLf)
    
    'We get the disk ID, There disk ID is differen't then
    'the one we got before we started. Why? I have no clue.
    'You're going to have to find that one out on your own.
    
    'Get the new DISK ID
    TempString = InStr(1, TempStringArray(1), " ", vbTextCompare) + 1
    TempString = Mid(TempStringArray(1), TempString, InStr(TempString, TempStringArray(1), " ", vbTextCompare) - TempString)
    DiskID = TempString
    
    'Get Category
    DiskCategory = Left(TempStringArray(1), InStr(1, TempStringArray(1), " ", vbTextCompare) - 1)

    Winsock1.SendData "cddb read " & DiskCategory & " " & DiskID & vbCrLf

    Debug.Print "** CD Found **"
    'To be safe, lets clear the TempString
    TempString = ""
End If

If InStr(1, Text, " CD database entry follows (until terminating `.')", vbTextCompare) Then

    TempStringArray = Split(Text, vbCrLf)
    
    'Just to play it safe, lets reset the TrackTitles() & TackCount
    ReDim TrackTitles(9) As String
    TrackCount = 0
        
    For a = 0 To UBound(TempStringArray)
        'Disk Artist / Album Name
        If InStr(1, TempStringArray(a), "DTITLE=", vbTextCompare) Then
            TempString = Mid(TempStringArray(a), 8, Len(TempStringArray(a)) - 7)

            TempStringArray2$ = Split(TempString, " / ")
            
            ArtistName = TempStringArray2$(0)
            AlbumName = TempStringArray2$(1)
            
            'To be safe, lets clear the TempString
            TempString = ""
        End If
        
        If InStr(1, TempStringArray(a), "TTITLE", vbTextCompare) Then
            'We are only getting the CD Title and Track Titles.
            'There is more info that we can get. But thats also
            'something you'll need to code.

            TempString = InStr(1, TempStringArray(a), "=", vbTextCompare) + 1
            TempString = Mid(TempStringArray(a), TempString, Len(TempStringArray(a)) - TempString + 1)
            
            ReDim Preserve TrackTitles(TrackCount) As String
            
            TrackTitles(TrackCount) = TempString
            
            TrackCount = TrackCount + 1
            
            Debug.Print "'" & TempString & "'"
            
            Winsock1.SendData "quit" & vbCrLf
            
            'To be safe, lets clear the TempString
            TempString = ""
        End If
    Next

End If


End Sub

