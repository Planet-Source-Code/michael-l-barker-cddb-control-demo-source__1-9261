Attribute VB_Name = "ExtraStuff"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Const MAX_COMPUTERNAME_LENGTH As Long = 15&

Public Function CurrentMachineName() As String

Dim lSize As Long
Dim sBuffer As String
sBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
lSize = Len(sBuffer)

If GetComputerName(sBuffer, lSize) Then
    CurrentMachineName = Left$(sBuffer, lSize)
End If

End Function

