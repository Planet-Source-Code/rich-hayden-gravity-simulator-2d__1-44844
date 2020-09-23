Attribute VB_Name = "mdlMain"
'Copyright (c) 2003 Richard Hayden. All Rights Reserved.
'
'You may use any code contained in this project for NON-commercial gain
'as long as credit is given to the author (Richard Hayden).
'
'Thanks for downloading my code!

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_USER = &H400&
Public Const TBM_GETTOOLTIPS = WM_USER + 30
Public Const TTM_ACTIVATE = WM_USER + 1

Sub Main()
    Load frmMain
    frmMain.Show
End Sub


