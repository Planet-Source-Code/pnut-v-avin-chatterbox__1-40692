VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Picture..."
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "->"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton btnPrev 
      Caption         =   "<-"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  AVIN Chatterbox
'  Copyright 2002 by Andrew Vaughan
'
'  Please vote for me on PSC!

Dim imgNum As Integer   ' Current image number

' Preview the next image
Private Sub btnNext_Click()
    ' If we go above the maximum number of images, then stop!
    If imgNum < 105 Then imgNum = imgNum + 1
    
    ' Load the requested picture into the preview area
    Picture1.Picture = LoadPicture(App.Path & "\img" & imgNum & ".gif")
End Sub

' Preview the previous image
Private Sub btnPrev_Click()
    ' If we go below 1, then stop!
    If imgNum > 1 Then imgNum = imgNum - 1
    
    ' Load the requested picture into the preview area
    Picture1.Picture = LoadPicture(App.Path & "\img" & imgNum & ".gif")
End Sub

' OK Button...
Private Sub Command1_Click()
    ' Make sure we are connected...
    If Form1.WS.State = sckConnected Then
        ' Send the appropriate information for the image
        Form1.WS.SendData "IMG" & Chr$(175) & imgNum
    End If
    
    ' Go back to the screen
    Unload Me
End Sub

' Cancel button
Private Sub Command2_Click()
    ' Go back to the screen
    Unload Me
End Sub

' Load the picture
Private Sub Form_Load()
    btnNext_Click
    btnPrev_Click
End Sub
