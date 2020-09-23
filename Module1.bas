Attribute VB_Name = "Module1"
Option Explicit

' ================================================================================
' MIDAR's Froggies Game
'
' If you modify the graphics, sounds or gameplay, send me an e-mail to explain the
' details - If your game is good enough, I'll put it on my web site for you.
'
' Peter Wilson
' MIDAR Pty Ltd
' E-Mail:  peter@midar.com.au
' Web   :    www.midar.com.au
' ================================================================================

' Used for measuring time down to the millisecond (there a 1000 milliseconds in a sec.)
Public Declare Function GetTickCount Lib "kernel32" () As Long

' Used for playing sounds.
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public g_DisplayName As String

Public Function AppTitle() As String

    AppTitle = App.Title & " v." & App.Major & "." & App.Minor & "." & App.Revision
    
End Function


