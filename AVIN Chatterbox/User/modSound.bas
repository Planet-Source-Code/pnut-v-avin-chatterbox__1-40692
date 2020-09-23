Attribute VB_Name = "mdlSound"
'  AVIN Chatterbox
'  Copyright 2002 by Andrew Vaughan
'
' Please vote for me on PSC!

Option Explicit ' Make sure we declare all of the variables (to prevent bugs)

' API call to play a sound
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' The IM Sound
Public Enum enuSound
    IMSound = 1
End Enum

Dim sndFileName As String

Public Sub PlaySound(sndEvent As enuSound)
    Dim Filename As String  ' Get the filename
    Dim a As Integer        ' Temporary integer
    
    Select Case sndEvent    ' Get the possible sounds
        ' If it is 1, play the im sound
        Case 1: Filename = "\im.wav"
    End Select
    
    ' Format the filename
    If Right(App.Path, 2) = ":\" Then
        Filename = App.Path & Filename
    Else
        Filename = App.Path & "\" & Filename
    End If
    
    ' If filename is blank, get an inserted one
    If Filename = "" Then Filename = sndFileName
    
    ' If there is an error, say so!
    On Error GoTo Err
    
    ' Play the sound
    sndPlaySound Filename, 3
    
    ' Close me
    Exit Sub

Err:
    ' Oh no!
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
End Sub
