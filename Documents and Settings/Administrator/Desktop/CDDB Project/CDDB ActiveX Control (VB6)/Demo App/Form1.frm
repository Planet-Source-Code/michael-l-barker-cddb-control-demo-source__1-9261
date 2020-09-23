VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9199428E-6E93-4EAF-B611-977B263AE9B7}#1.0#0"; "CDDBControl_vb6.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CDDB ActiveX Control Demo App"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3780
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":000C
            Key             =   "kCD"
            Object.Tag             =   "tCD"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Info"
      Default         =   -1  'True
      Height          =   375
      Left            =   4860
      TabIndex        =   7
      Top             =   1080
      Width           =   1275
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "D"
      Top             =   420
      Width           =   255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   188
      TabIndex        =   6
      Top             =   1560
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   5212
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4650
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   188
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1140
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   188
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   3375
   End
   Begin CDDBControl.CDDB CDDB1 
      Left            =   5100
      Top             =   240
      _ExtentX        =   1535
      _ExtentY        =   1058
      DriveLetter     =   ""
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Drive Letter:"
      Height          =   195
      Left            =   3720
      TabIndex        =   2
      Top             =   180
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Artist Name:"
      Height          =   195
      Left            =   188
      TabIndex        =   4
      Top             =   900
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Album Name:"
      Height          =   195
      Left            =   188
      TabIndex        =   0
      Top             =   180
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CDDB1_AllServerMessages(Text As String)
    Debug.Print Text
End Sub


Private Sub CDDB1_CDInfo(AlbumName As String, ArtistName As String, TotalTracks As Long)

Dim a As Long
Dim ItemToAdd As ListItem

StatusBar1.SimpleText = "Getting Info..."

Text1.Text = AlbumName
Text2.Text = ArtistName

For a = 0 To TotalTracks

Set ItemToAdd = ListView1.ListItems.Add(, , , , 1)

StatusBar1.SimpleText = "... " & CDDB1.GetTrackName(a)

ItemToAdd.Text = a + 1
ItemToAdd.SubItems(1) = CDDB1.GetTrackName(a)

Next

StatusBar1.SimpleText = TotalTracks + 1 & " Tracks Found"

End Sub


Private Sub CDDB1_Connected()
    StatusBar1.SimpleText = "Connected"
End Sub

Private Sub CDDB1_Disconnected()
    StatusBar1.SimpleText = "Disconnected"
End Sub


Private Sub Command1_Click()

Text1.Text = ""
Text2.Text = ""
ListView1.ListItems.Clear

CDDB1.DriveLetter = Text3.Text
CDDB1.Connect Random_US_site

End Sub

Private Sub Form_Load()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

ListView1.FullRowSelect = True
ListView1.View = lvwReport

ListView1.ColumnHeaders.Add 1, , "Track Number", 2000
ListView1.ColumnHeaders.Add 2, , "Track Name", 3600

End Sub


