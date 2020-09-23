VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font Name..."
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option7 
      Caption         =   "7"
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton Option6 
      Caption         =   "6"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton Option5 
      Caption         =   "5"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton Option4 
      Caption         =   "4"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.OptionButton Option3 
      Caption         =   "3"
      Height          =   255
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton Option2 
      Caption         =   "2"
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  AVIN Chatterbox
'  Copyright 2002 by Andrew Vaughan
'
'  Please vote for me on PSC!

'  NOTE - Fonts only work with web pages if they are on BOTH computer!

' Save button...
Private Sub Command1_Click()
    With Form1.txtIn.Font
        ' Set the appropriate font name
        .Name = Text1.Text
        
        ' Set the font size for the textbox to the appropriate size
        If Option1.Value = True Then
            .Size = 9
        ElseIf Option2.Value = True Then
            .Size = 10
        ElseIf Option3.Value = True Then
            .Size = 12
        ElseIf Option4.Value = True Then
            .Size = 14
        ElseIf Option5.Value = True Then
            .Size = 18
        ElseIf Option6.Value = True Then
            .Size = 24
        ElseIf Option7.Value = True Then
            .Size = 36
        End If
    End With
    
    ' Go back to the main screen
    Me.Hide
End Sub

Private Sub Form_Load()
    ' Get the current font name
    Text1.Text = Form1.txtIn.Font.Name
    
    ' Get the current font size and select the appropriate button
    Select Case Form1.txtIn.Font.Size
        Case 9
            Option1.Value = True
        Case 12
            Option3.Value = True
        Case 14
            Option4.Value = True
        Case 18
            Option5.Value = True
        Case 24
            Option6.Value = True
        Case 36
            Option7.Value = True
        Case Else
            Option2.Value = True
    End Select
End Sub
