Attribute VB_Name = "DateFormatHandler"
Option Explicit

'This code is just a revamp of several other VB programmers' code
'  which I used to change the short date format in a program that
'  uses MS SQL Server 2005. I thought it may be usefull to others.

Private Const LOCALE_SSHORTDATE = &H1F
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SYSTEM_DEFAULT = &H800

Declare Function GetLocaleInfo Lib "kernel32" Alias _
    "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String, ByVal cchData As Long) As Long

Declare Function SetLocaleInfo Lib "kernel32" Alias _
    "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String) As Boolean
    
Private Const LOCALE_SDATE = &H1F
Private Const LOCALE_STIMEFORMAT = &H1003
Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Function GetDateFormat() As String
Dim sBuff As String
Dim x As Long

sBuff = Space$(64)
x = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuff, Len(sBuff))

If x > 0 Then
  sBuff = Left$(sBuff, x - 1)
  GetDateFormat = Left$(sBuff, x - 1)
End If

End Function

Public Sub SetDateFormat(sDate As String)
Dim dwLCID As Long

dwLCID = GetSystemDefaultLCID()

If SetLocaleInfo(dwLCID, LOCALE_SDATE, sDate) = False Then
    Exit Sub
End If

PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0

End Sub

