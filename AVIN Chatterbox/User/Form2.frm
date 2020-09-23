VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font Color..."
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optYel 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton optGreen 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton optRed 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton optMag 
      BackColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton optCyan 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton optBlue 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton optBlack 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  AVIN Chatterbox
'  Copyright 2002 by Andrew Vaughan
'
'  Please vote for me on PSC!

' OK Button (Saves the color)
Private Sub Command1_Click()
    Dim a As ColorConstants     ' Color constants
    
    ' If black, make a = black
    If optBlack.Value = True Then
        a = vbBlack
    ' If blue, make a = blue
    ElseIf optBlue.Value = True Then
        a = vbBlue
    ' If cyan, make a = cyan
    ElseIf optCyan.Value = True Then
        a = vbCyan
    ' If green, make a = green
    ElseIf optGreen.Value = True Then
        a = vbGreen
    ' If magenta, make a = magenta
    ElseIf optMag.Value = True Then
        a = vbMagenta
    ' If red, make a = red
    ElseIf optRed.Value = True Then
        a = vbRed
    ' If yellow, make a = yellow
    ElseIf optYel.Value = True Then
        a = vbYellow
    End If
    
    ' Make the button for the font turn the color
    Form1.Command1.BackColor = a
    
    ' Make the text-box font turn that color
    Form1.txtIn.ForeColor = a
    
    ' Go back to the main screen
    Unload Form2
End Sub

' Load this dialog box
Private Sub Form_Load()
    ' Set the colors of the buttons
    optBlack.BackColor = vbBlack
    optBlue.BackColor = vbBlue
    optCyan.BackColor = vbCyan
    optMag.BackColor = vbMagenta
    optGreen.BackColor = vbGreen
    optRed.BackColor = vbRed
    optYel.BackColor = vbYellow
    
    ' See what color the button has and select the appropriate button
    Select Case Form1.Command1.BackColor
        Case vbBlue
            optBlue.Value = True
        Case vbCyan
            optCyan.Value = True
        Case vbGreen
            optGreen.Value = True
        Case vbMagenta
            optMag.Value = True
        Case vbRed
            optRed.Value = True
        Case vbYellow
            optYel.Value = True
        Case Else
            optBlack.Value = True
    End Select
End Sub
