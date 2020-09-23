VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl CDDB 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
   InvisibleAtRuntime=   -1  'True
   Picture         =   "CDDB.ctx":0000
   PropertyPages   =   "CDDB.ctx":088B
   ScaleHeight     =   600
   ScaleWidth      =   870
   Windowless      =   -1  'True
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "CDDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'CDDB ActiveX Control For VB5/VB6
'Compiled With VB6 SP3
'Note: You will NEED VB6 to compile this again.
'      Because I used a VB6 Function 'Split'
'      For more info, open your MSDN Help (F1)
'      and jump to this URL:
'      mk:@MSITStore:C:\Program%20Files\Microsoft%20Visual%20Studio\MSDN\2000APR\1033\vbenlr98.chm::/html/vafctSplit.htm
'      Or, paste that in your browser. If the path
'      Doesn't match your path. It doesn't matter
'      It will still find it.
'June 24, 2000
'Known Bugs? Not sure, I won't be using this control
'I only remade it because it was a request.
'There might be errors, I redid this whole control
'from the start. Took less then a day to make.
'Author: Michael L. Barker

Option Explicit

Public Event AllServerMessages(Text As String)
Public Event CDInfo(AlbumName As String, ArtistName As String, TotalTracks As Long)
Public Event Connected()
Public Event Disconnected()


Dim TrackCount As Long
Dim QueryString As String
Dim DiskID As String
Dim DiskCategory As String
Dim AlbumName As String
Dim ArtistName As String
Dim TrackTitles() As String
Dim TempStringArray$()
Dim TempStringArray2$()
Dim TempString As String

Dim sDriveLetter As String

Public Enum CDDB_Server
    Random_US_site = 0
    San_Jose_CA_US = 1
    Santa_Clara_CA_US = 2
End Enum

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

sDriveLetter = PropBag.ReadProperty("DriveLetter", "D")

End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 870
    UserControl.Height = 600
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DriveLetter", sDriveLetter, "D"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Text As String
Dim a As Long

Winsock1.GetData Text, vbString

RaiseEvent AllServerMessages(Text)

If InStr(1, Text, "200 Hello and welcome ", vbTextCompare) Then
    Winsock1.SendData QueryString & vbCrLf
End If

If InStr(1, Text, "200 ", vbTextCompare) Then
    If InStr(1, Text, " / ", vbTextCompare) Then
        
        TempString = InStr(4, Text, " ", vbTextCompare) + 1
        DiskCategory = Mid(Text, TempString, InStr(TempString, Text, " ", vbTextCompare) - TempString)
        
        TempString = InStr(5, Text, " ", vbTextCompare) + 1
        DiskID = Mid(Text, TempString, InStr(TempString, Text, " ", vbTextCompare) - TempString)
        
        Winsock1.SendData "cddb read " & DiskCategory & " " & DiskID & vbCrLf
        
    End If
End If

If InStr(1, Text, "211 Found inexact matches, list follows (until terminating `.')", vbTextCompare) Then
    TempStringArray = Split(Text, vbCrLf)
    
    TempString = InStr(1, TempStringArray(1), " ", vbTextCompare) + 1
    TempString = Mid(TempStringArray(1), TempString, InStr(TempString, TempStringArray(1), " ", vbTextCompare) - TempString)
    DiskID = TempString
    
    DiskCategory = Left(TempStringArray(1), InStr(1, TempStringArray(1), " ", vbTextCompare) - 1)

    Winsock1.SendData "cddb read " & DiskCategory & " " & DiskID & vbCrLf

    Debug.Print "** CD Found **"
    TempString = ""
End If

If InStr(1, Text, " CD database entry follows (until terminating `.')", vbTextCompare) Then

    TempStringArray = Split(Text, vbCrLf)
    
    ReDim TrackTitles(1) As String
    TrackCount = 0
        
    For a = 0 To UBound(TempStringArray)
        If InStr(1, TempStringArray(a), "DTITLE=", vbTextCompare) Then
            TempString = Mid(TempStringArray(a), 8, Len(TempStringArray(a)) - 7)

            TempStringArray2$ = Split(TempString, " / ")
            
            ArtistName = TempStringArray2$(0)
            AlbumName = TempStringArray2$(1)
            
            TempString = ""
        End If
        
        If InStr(1, TempStringArray(a), "TTITLE", vbTextCompare) Then

            TempString = InStr(1, TempStringArray(a), "=", vbTextCompare) + 1
            TempString = Mid(TempStringArray(a), TempString, Len(TempStringArray(a)) - TempString + 1)
            
            ReDim Preserve TrackTitles(TrackCount) As String
            
            TrackTitles(TrackCount) = TempString
            
            TrackCount = TrackCount + 1
            
            Debug.Print "'" & TempString & "'"
            
            Winsock1.SendData "quit" & vbCrLf
            
            TempString = ""
        End If
    Next

End If


End Sub


Private Sub Winsock1_Close()

Winsock1.Close
RaiseEvent Disconnected
RaiseEvent CDInfo(Trim(AlbumName), Trim(ArtistName), UBound(TrackTitles))

End Sub
Public Property Get DriveLetter() As String
Attribute DriveLetter.VB_ProcData.VB_Invoke_Property = "PropertyPage1"

DriveLetter = sDriveLetter

End Property

Public Property Let DriveLetter(Letter As String)

sDriveLetter = Left(Letter, 1)

End Property

Public Function QueryCD() As String

Dim CDClass As CCd
Set CDClass = New CCd

CDClass.Init sDriveLetter & ":"

QueryString = CDClass.QueryString
DiskID = CDClass.DiscID

Set CDClass = Nothing

End Function

Public Function Connect(Server As CDDB_Server) As Boolean

On Error GoTo ErrorHand

Dim ServerAddr As String

Select Case Server
    Case 0: ServerAddr = "us.cddb.com"
    Case 1: ServerAddr = "sj.ca.us.cddb.com"
    Case 2: ServerAddr = "sc.ca.us.cddb.com"
End Select

QueryCD

If QueryString = "Drive not ready. Try again." Then
    MsgBox "Drive not ready. Try again."
    Exit Function
End If

Winsock1.LocalPort = 8880
Winsock1.RemotePort = 8880
Winsock1.Connect ServerAddr, 8880

Exit Function

ErrorHand:
Connect = False

End Function

Private Sub Winsock1_Connect()

RaiseEvent Connected
Winsock1.SendData "cddb hello " & CurrentMachineName & " " & Winsock1.LocalIP & " CDDB_ACTIVE_X_CONTROL_FREEWARE 1.0.0" & vbCrLf

End Sub



Public Function GetTrackName(Index As Long)
On Error Resume Next

GetTrackName = TrackTitles(Index)

End Function

Public Function GetTrackCount() As Long
On Error Resume Next

GetTrackCount = UBound(TrackTitles)

End Function

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    MsgBox "Version 1.0.0" & vbCrLf & "June 24, 2000", vbInformation, "CDDB ActiveX Control"
End Sub
