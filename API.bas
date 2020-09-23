Attribute VB_Name = "API"
'Used To Hide and SHow The Cursor
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Const SND_FILENAME = &H20000
Public Const SND_LOOP = &H8
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

