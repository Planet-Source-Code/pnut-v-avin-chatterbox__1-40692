VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVIN Chatterbox"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      Height          =   315
      Left            =   7680
      TabIndex        =   13
      Top             =   4920
      Width           =   255
   End
   Begin SHDocVwCtl.WebBrowser Web 
      CausesValidation=   0   'False
      Height          =   4815
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   7935
      ExtentX         =   13996
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   7935
      Begin VB.CommandButton Command6 
         Caption         =   "Font"
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear Chat"
         Height          =   375
         Left            =   6480
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "PS"
         Height          =   375
         Left            =   7080
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   7200
         Top             =   360
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Picture..."
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton btnSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   7080
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton btnClr 
         Caption         =   "Clear"
         Height          =   255
         Left            =   7080
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtIn 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   720
         Width           =   6855
      End
      Begin VB.CheckBox ckStrike 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox ckUnd 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox ckItal 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox ckBold 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Input: (HTML Allowed)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "169.254.140.165"
   End
   Begin VB.Menu asfg 
      Caption         =   "Shortcuts"
      Begin VB.Menu mnuSend 
         Caption         =   "Send"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuPS 
         Caption         =   "Private Send"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Picture"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  AVIN Chatterbox
'  Copyright 2002 by Andrew Vaughan
'
'  Please vote for me on PSC!

Option Explicit ' Make it so I have to declare ALL variables (prevents bugs)

Dim SeperatorChar As String       ' The seperator for sending information
Dim Quote As String     ' The '"' character
Dim myName As String    ' Your name that you enter
Dim BKMK As Long        ' Keeps the place in the html file (so it always shows the last entry)
Dim Header As String    ' The HTML header
Dim ConIP As String     ' The IP the server is on

' Clears the input box text
Private Sub btnClr_Click()
    txtIn.Text = "" ' Clear text
End Sub

' Sends the information entered to the server
Private Sub btnSend_Click()
    Dim tempString As String    ' Temporary String
    Dim x As Long               ' Counter
    
    ' If the user clicked send without entering any text, send an error message
    If txtIn.Text = "" Then MsgBox "Please insert text!", vbOKOnly + vbCritical, "Error": Exit Sub
    
    ' Add styles to text...
    If ckBold.Value = 1 Then tempString = tempString & "<B>"    ' Add the HTML syntax for bold
    If ckItal.Value = 1 Then tempString = tempString & "<I>"    ' Add the HTML syntax for itallic
    If ckUnd.Value = 1 Then tempString = tempString & "<U>"     ' Add the HTML syntax for underline
    If ckStrike.Value = 1 Then tempString = tempString & "<S>"  ' Add the HTML syntax for strikthrough
    
    ' Check the possible text colors...
    Select Case Command1.BackColor
        Case vbBlue
            ' If blue, put the HTML syntax for a blue font
            tempString = tempString & "<FONT COLOR=" & Quote & "Blue" & Quote
        Case vbCyan
            ' If cyan, put the HTML syntax for a cyan font
            tempString = tempString & "<FONT COLOR=" & Quote & "Cyan" & Quote
        Case vbGreen
            ' If green, put the HTML syntax for a green font
            tempString = tempString & "<FONT COLOR=" & Quote & "Green" & Quote
        Case vbMagenta
            ' If magenta, put the HTML syntax for a magenta font
            tempString = tempString & "<FONT COLOR=" & Quote & "Magenta" & Quote
        Case vbRed
            ' If red, put the HTML syntax for a red font
            tempString = tempString & "<FONT COLOR=" & Quote & "Red" & Quote
        Case vbYellow
            ' If yellow, put the HTML syntax for a yellow font
            tempString = tempString & "<FONT COLOR=" & Quote & "Yellow" & Quote
        Case Else
            ' If the color is anything else, put the HTML syntax for a black font
            tempString = tempString & "<FONT COLOR=" & Quote & "Black" & Quote
    End Select
    
    ' Put in the HTML syntax for the selected font...
    tempString = tempString & " Face=" & Quote & txtIn.Font.Name & Quote
    
    ' Check the possible font sizes... (HTML only has 7 sizes)
    Select Case txtIn.Font.Size
        Case 9
            ' HTML Font Size    = 1
            ' Windows Font Size = 9
            tempString = tempString & " Size=1>"
        Case 12
            ' HTML Font Size    = 3
            ' Windows Font Size = 12
            tempString = tempString & " Size=3>"
        Case 14.25
            ' HTML Font Size    = 4
            ' Windows Font Size = 14.25
            tempString = tempString & " Size=4>"
        Case 18
            ' HTML Font Size    = 5
            ' Windows Font Size = 18
            tempString = tempString & " Size=5>"
        Case 24
            ' HTML Font Size    = 6
            ' Windows Font Size = 24
            tempString = tempString & " Size=6>"
        Case 36
            ' HTML Font Size    = 7
            ' Windows Font Size = 36
            tempString = tempString & " Size=7>"
        Case Else
            ' If 2 is selected or there is an error...
            
            ' HTML Font Size    = 2
            ' Windows Font Size = 10
            tempString = tempString & " Size=2>"
    End Select
    
    ' Process the text and add it to the temporary string...
    
    ' PS (Browsers can't see return signals (enter), so you must enter the
    ' HTML syntax for a return (<BR>) so that it puts cairrage returns...
    For x = 1 To Len(txtIn.Text)
        ' For every return in the text, place a <BR>
        If Asc(Mid$(txtIn.Text, x, 1)) <= 10 Then tempString = tempString & "<BR>" Else: tempString = tempString & Mid$(txtIn.Text, x, 1)
    Next x
    
    ' Reset all of the settings so they don't interfere with other messages
    If ckBold.Value = 1 Then tempString = tempString & "</B>"   ' HTML syntax for no bold
    If ckItal.Value = 1 Then tempString = tempString & "</I>"   ' HTML syntax for no itallics
    If ckUnd.Value = 1 Then tempString = tempString & "</U>"    ' HTML syntax for no underline
    If ckStrike.Value = 1 Then tempString = tempString & "</S>" ' HTML syntax for no strikethrough
    tempString = tempString & "</Font>" ' HTML syntax for no more font changes (color, size and face)
    
    ' Format the string so that it is sent to everybody on the server (PUBL)
    tempString = "PUBL" & SeperatorChar & myName & SeperatorChar & tempString
    
    ' Make sure that we are connected to the network
    If WS.State = sckConnected Then
        ' If we are, send the data
        WS.SendData tempString
    End If
    
    ' Clear the textbox and put the focus on it so you don't have to keep
    ' clicking over to it after you type a message
    txtIn.Text = ""
    txtIn.SetFocus
End Sub

Private Sub ckBold_Click()
    ' Make the text bold
    If ckBold.Value = 0 Then txtIn.Font.Bold = False Else: txtIn.Font.Bold = True
End Sub

Private Sub ckItal_Click()
    ' Make the text itallics
    If ckItal.Value = 0 Then txtIn.Font.Italic = False Else: txtIn.Font.Italic = True
End Sub

Private Sub ckStrike_Click()
    ' Make the text strikethrough
    If ckStrike.Value = 0 Then txtIn.Font.Strikethrough = False Else: txtIn.Font.Strikethrough = True
End Sub

Private Sub ckUnd_Click()
    ' Make the text underlined
    If ckUnd.Value = 0 Then txtIn.Font.Underline = False Else: txtIn.Font.Underline = True
End Sub

Private Sub Command1_Click()
    ' Show the font color dialog box
    Form2.Show vbModal
End Sub

Private Sub Command2_Click()
    ' Show the insert picture dialog box
    Form3.Show vbModal
End Sub

' This allows for private sending over the network!  (So only a certain user
' can see it)
Private Sub Command3_Click()
    Dim toName As String    ' The person you're sending it to
    Dim tempString As String, sa As String, sb As String    ' Temporary String
    Dim x As Long           ' Counter
    
    ' If the user clicked send without entering any text, send an error message
    If txtIn.Text = "" Then MsgBox "Please insert text!", vbOKOnly + vbCritical, "Error": Exit Sub
    
    ' Get the username of the person to send it to
    toName = InputBox("Private send to which name?", "Private Send")
    
    ' Add styles to text...
    If ckBold.Value = 1 Then tempString = tempString & "<B>"    ' Add the HTML syntax for bold
    If ckItal.Value = 1 Then tempString = tempString & "<I>"    ' Add the HTML syntax for itallic
    If ckUnd.Value = 1 Then tempString = tempString & "<U>"     ' Add the HTML syntax for underline
    If ckStrike.Value = 1 Then tempString = tempString & "<S>"  ' Add the HTML syntax for strikthrough
    
    ' Check the possible text colors...
    Select Case Command1.BackColor
        Case vbBlue
            ' If blue, put the HTML syntax for a blue font
            tempString = tempString & "<FONT COLOR=" & Quote & "Blue" & Quote
        Case vbCyan
            ' If cyan, put the HTML syntax for a cyan font
            tempString = tempString & "<FONT COLOR=" & Quote & "Cyan" & Quote
        Case vbGreen
            ' If green, put the HTML syntax for a green font
            tempString = tempString & "<FONT COLOR=" & Quote & "Green" & Quote
        Case vbMagenta
            ' If magenta, put the HTML syntax for a magenta font
            tempString = tempString & "<FONT COLOR=" & Quote & "Magenta" & Quote
        Case vbRed
            ' If red, put the HTML syntax for a red font
            tempString = tempString & "<FONT COLOR=" & Quote & "Red" & Quote
        Case vbYellow
            ' If yellow, put the HTML syntax for a yellow font
            tempString = tempString & "<FONT COLOR=" & Quote & "Yellow" & Quote
        Case Else
            ' If the color is anything else, put the HTML syntax for a black font
            tempString = tempString & "<FONT COLOR=" & Quote & "Black" & Quote
    End Select
    
    ' Put in the HTML syntax for the selected font...
    tempString = tempString & " Face=" & Quote & txtIn.Font.Name & Quote
    
    ' Check the possible font sizes... (HTML only has 7 sizes)
    Select Case txtIn.Font.Size
        Case 9
            ' HTML Font Size    = 1
            ' Windows Font Size = 9
            tempString = tempString & " Size=1>"
        Case 12
            ' HTML Font Size    = 3
            ' Windows Font Size = 12
            tempString = tempString & " Size=3>"
        Case 14.25
            ' HTML Font Size    = 4
            ' Windows Font Size = 14.25
            tempString = tempString & " Size=4>"
        Case 18
            ' HTML Font Size    = 5
            ' Windows Font Size = 18
            tempString = tempString & " Size=5>"
        Case 24
            ' HTML Font Size    = 6
            ' Windows Font Size = 24
            tempString = tempString & " Size=6>"
        Case 36
            ' HTML Font Size    = 7
            ' Windows Font Size = 36
            tempString = tempString & " Size=7>"
        Case Else
            ' If 2 is selected or there is an error...
            
            ' HTML Font Size    = 2
            ' Windows Font Size = 10
            tempString = tempString & " Size=2>"
    End Select
    
    ' Process the text and add it to the temporary string...
    
    ' PS (Browsers can't see return signals (enter), so you must enter the
    ' HTML syntax for a return (<BR>) so that it puts cairrage returns...
    For x = 1 To Len(txtIn.Text)
        ' For every return in the text, place a <BR>
        If Asc(Mid$(txtIn.Text, x, 1)) <= 10 Then tempString = tempString & "<BR>" Else: tempString = tempString & Mid$(txtIn.Text, x, 1)
    Next x
    
    ' Reset all of the settings so they don't interfere with other messages
    If ckBold.Value = 1 Then tempString = tempString & "</B>"   ' HTML syntax for no bold
    If ckItal.Value = 1 Then tempString = tempString & "</I>"   ' HTML syntax for no itallics
    If ckUnd.Value = 1 Then tempString = tempString & "</U>"    ' HTML syntax for no underline
    If ckStrike.Value = 1 Then tempString = tempString & "</S>" ' HTML syntax for no strikethrough
    tempString = tempString & "</Font>" ' HTML syntax for no more font changes (color, size and face)
    
    ' Format the string so that it is sent to ONLY the specified person (PRIV)
    sa = "PRIV" & SeperatorChar & LCase$(toName) & SeperatorChar & myName & SeperatorChar & tempString
    
    ' Send it privately to yourself too, so you can see it too
    DoData "PRIV" & SeperatorChar & myName & " to " & toName & SeperatorChar & myName & SeperatorChar & tempString
    
    ' Make sure that we are connected to the network
    If WS.State = sckConnected Then
        ' If we are, send the data
        WS.SendData sa
    End If
    
    ' Clear the textbox and put the focus on it so you don't have to keep
    ' clicking over to it after you type a message
    txtIn.Text = ""
    txtIn.SetFocus
End Sub

' Exit the program and tell the server we're leaving
Private Sub Command4_Click()
    Dim a As Integer    ' Temporary integer
    
    ' See what font size there is
    Select Case txtIn.Font.Size
        Case 9: a = 1
        Case 12: a = 3
        Case 14: a = 4
        Case 18: a = 5
        Case 24: a = 6
        Case 36: a = 7
        Case Else: a = 2
    End Select
    
    ' Save the current settings in the registry for later use
    SaveSetting "Chatterbox", "Font", "Color", Command1.BackColor
    SaveSetting "Chatterbox", "Font", "Face", txtIn.Font.Name
    SaveSetting "Chatterbox", "Font", "Bold", ckBold.Value
    SaveSetting "Chatterbox", "Font", "Ital", ckItal.Value
    SaveSetting "Chatterbox", "Font", "Und", ckUnd.Value
    SaveSetting "Chatterbox", "Font", "Strike", ckStrike.Value
    SaveSetting "Chatterbox", "Font", "Size", Str(a)
    
    DoEvents    ' Let Windows do it's thing
    
    ' Make sure we're connected
    If WS.State = 7 Then
        ' Tell the server we're leaving
        WS.SendData "PUBL" & SeperatorChar & " " & SeperatorChar & myName & " has left the session"
    End If
    
    DoEvents    ' Let Windows do it's thing again
    
    WS.Close    ' Close the connection to the server
    
    End ' Close me
End Sub

' Clear the chat (but keep the header)
Private Sub Command5_Click()
    ' Clear the header
    Header = ""
    
    ' Rewrite the header
    ' This baisically just says:
        ' Make the Caption of the window say AVIN Chatterbox Chat
        ' Make the bookmark = 0 (the top of the page)
    Header = "<HTML><HEAD><TITLE>AVIN Chatterbox Chat</TITLE>" & vbNewLine & _
             "<SCRIPT language=" & Quote & "vbscript" & Quote & ">" & vbNewLine & _
             "sub tobot()" & vbNewLine & _
             "  bot0.click" & vbNewLine & _
             "end sub" & vbNewLine & _
             "</SCRIPT></HEAD>"
    
    ' Set the bookmark to 0 (the top of the page)
    BKMK = 0
    
    ' Save the bookmark number
    SaveSetting "Chatterbox", "Bkmk", "Bkmk", BKMK
    
    DoEvents    ' Let Windows do it's thing
    
    ' Open the temporary file where the html file is stored
    Open "C:\Windows\Temp\avchft.html" For Output As #1
        ' Write the header (above)
        Print #1, Header
        
        ' Make it so that it's a white background with black text and set to the first bookmark
        Print #1, "<BODY BGCOLOR=" & Quote & "WHITE" & Quote & " TEXT=" & Quote & "BLACK" & Quote & " onLoad=" & Quote & "tobot" & Quote & ">"
        ' Set the font for the title
        Print #1, "<FONT FACE=" & Quote & "arial" & Quote & " SIZE=" & Quote & "4" & Quote & " COLOR=" & Quote & "BLACK" & Quote & ">"
        ' Write the title, followed by 2 returns
        Print #1, "Welcome to AVIN Chatterbox<BR><HR>"
        ' Set the font for the info
        Print #1, "<FONT FACE=" & Quote & "arial" & Quote & " SIZE=" & Quote & "1" & Quote & " COLOR=" & Quote & "BLACK" & Quote & ">"
        ' Write the info
        Print #1, "<I>(If your text does not appear then the server is not up.)</I><BR><BR>"
        ' Set the first bookmark (the top of the page)
        Print #1, "<A id=" & Quote & "bot0" & Quote & " name=" & Quote & "bot0" & Quote & " href=" & Quote & "#bot1" & Quote & "></a></FONT>"
        ' Reset the font for the first message
        Print #1, "<FONT FACE=" & Quote & "arial" & Quote & " size=" & Quote & "2" & Quote & ">"
    Close #1    ' Close the file
    
    DoEvents    ' Let Windows do it's thing
    
    ' Tell the browser to load the file we just made
    Web.Navigate "C:\Windows\Temp\avchft.html"
End Sub

' Show the font dialog
Private Sub Command6_Click()
    Form4.Show vbModal
End Sub

' Load the form
Private Sub Form_Load()
    Dim b As Integer    ' Temporary Integer
    
    ' Check to see if Chatterbox is already open so we don't open 2 at the same time
    If App.PrevInstance Then MsgBox "AVIN Chatterbox already open!", vbOKOnly + vbInformation, "Error": End
    
    ' Get the font color and set the appropriate button to uphold it
    Command1.BackColor = Val(GetSetting("Chatterbox", "Font", "Color", "0"))
    
    txtIn.ForeColor = Command1.BackColor    ' Set the font color for the text box
    txtIn.Font.Name = GetSetting("Chatterbox", "Font", "Face", "arial") ' Set the font name for the text box
    b = Val(GetSetting("Chatterbox", "Font", "Size", "3"))  ' Get the font size
    ckBold.Value = Val(GetSetting("Chatterbox", "Font", "Bold", "0"))   ' Get the bold
    ckItal.Value = Val(GetSetting("Chatterbox", "Font", "Ital", "0"))   ' Get the itallics
    ckUnd.Value = Val(GetSetting("Chatterbox", "Font", "Und", "0")) ' Get the underline
    ckStrike.Value = Val(GetSetting("Chatterbox", "Font", "Strike", "0"))   ' Get the strikethrough
    
    DoEvents    'Let Windows do it's thing
    
    ' Transfer the HTML style font to Windows format
    Select Case b
        Case 1
            txtIn.Font.Size = 9
        Case 3
            txtIn.Font.Size = 12
        Case 4
            txtIn.Font.Size = 14
        Case 5
            txtIn.Font.Size = 18
        Case 6
            txtIn.Font.Size = 24
        Case 7
            txtIn.Font.Size = 36
        Case Else   ' If there's an error, set the default
            txtIn.Font.Size = 10
    End Select
    
    ' Set the bold buttons to the appropriate level (on/off)
    ckBold_Click
    ckItal_Click
    ckUnd_Click
    ckStrike_Click
    
    ' Set the constants
    Quote = Chr$(34)
    SeperatorChar = Chr$(175)
    
    ' Set it to the first bookmark
    BKMK = 0
    
    ' Get the user's username
1   myName = InputBox("Please Type in Your Name:", "Type In Name...")
    If myName = "" Then GoTo 1  ' If they didn't type in anything, do it again
    
    DoEvents    ' Let Windows do it's thing
    
    ' IP = ###.###.###.###
    ' Computer Name = ANYTHING (Mine is 'Andrew')
    
    ' Get the IP or Computer Name of the server
24  ConIP = InputBox("Please enter name of computer or IP address of the server:", "255.255.255.255")
    If ConIP = "" Then GoTo 24
    
    DoEvents    ' Let Windows do it's thing again
    
    WS.Close    ' Close Winsock (incase it's open)
    WS.Connect ConIP, 420   ' Connect to the port given above
    
    DoEvents    ' Let Windows do it's thing AGAIN
    
    ' Write the starting HTML file...
22  SaveSetting "Chatterbox", "Bkmk", "Bkmk", "0"   ' Save the bookmark number
    
    Header = "" ' Clear the header
    
    ' Write the original header (see the clear chat button)
    Header = "<HTML><HEAD><TITLE>AVIN Chatterbox Chat</TITLE>" & vbNewLine & _
             "<SCRIPT language=" & Quote & "vbscript" & Quote & ">" & vbNewLine & _
             "sub tobot()" & vbNewLine & _
             "  bot0.click" & vbNewLine & _
             "end sub" & vbNewLine & _
             "</SCRIPT></HEAD>"
    
    ' Write the original file (see the clear chat button)
    Open "C:\Windows\Temp\avchft.html" For Output As #1
        Print #1, Header
        Print #1, "<BODY BGCOLOR=" & Quote & "WHITE" & Quote & " TEXT=" & Quote & "BLACK" & Quote & " onLoad=" & Quote & "tobot" & Quote & ">"
        Print #1, "<FONT FACE=" & Quote & "arial" & Quote & " SIZE=" & Quote & "4" & Quote & " COLOR=" & Quote & "BLACK" & Quote & ">"
        Print #1, "Welcome to AVIN Chatterbox<BR><HR>"
        Print #1, "<FONT FACE=" & Quote & "arial" & Quote & " SIZE=" & Quote & "1" & Quote & " COLOR=" & Quote & "BLACK" & Quote & ">"
        Print #1, "<I>(If your text does not appear then the server is not up.)</I><BR><BR>"
        Print #1, "<A id=" & Quote & "bot0" & Quote & " name=" & Quote & "bot0" & Quote & " href=" & Quote & "#bot1" & Quote & "></a></FONT>"
        Print #1, "<FONT FACE=" & Quote & "arial" & Quote & " size=" & Quote & "2" & Quote & ">"
    Close #1
    
    ' Locate and load the HTML file we just made
23  Web.Navigate "file://C:\Windows\Temp\avchft.html"
    
    ' Make sure we are connected to the server...
    If WS.State = 7 Then
        ' Tell everyone we are here
        WS.SendData "PUBL" & SeperatorChar & " " & SeperatorChar & "</I>" & myName & " has joined the session</I>"
    End If
    Exit Sub    ' Exit the sub

' ERROR CONTROL - If the server cannot be found
3   MsgBox "The Server is currently not available.", vbOKOnly + vbInformation, "Server not Available..."
    End ' Exit
End Sub

' Open the insert picture dialog
Private Sub mnuPic_Click()
    Command2_Click
End Sub

' Private send a message
Private Sub mnuPS_Click()
    Command3_Click
End Sub

' Public send a message
Private Sub mnuSend_Click()
    btnSend_Click
End Sub

' See's if we are connected to the server
Private Sub Timer1_Timer()
    ' Oh no!
    MsgBox "Connection timed out", vbOKOnly, "Error"
    End
End Sub

' Make sure the user doesn't insert the seperator character into the textbox
' (This shouldn't really happen since this character isn't on the keyboard,
'  but there are ways it can be inputed, so just to be safe...)
Private Sub txtIn_Change()
    With txtIn
        ' See if the last character entered was it
        ' If so, delete it
        If Right$(.Text, 1) = SeperatorChar Then .Text = Left$(.Text, Len(.Text) - 1)
    End With
End Sub

' See's if the server has lost connection (the server closed)
Private Sub WS_Close()
    ' Oh no!
    MsgBox "The Server has disconnected!", vbOKOnly + vbInformation, "Bye!"
    End
End Sub

' Gets messages from the server, splits them apart and posts them
Private Sub WS_DataArrival(ByVal bytesTotal As Long)
    Dim tempString As String    ' Temporary string
    
    WS.GetData tempString   ' Get the message from the server
    DoData tempString       ' Split it up and post it
End Sub

' Winsock error handler
Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' If the server is not up, end the program
    If Description = "Connection is forcefully rejected" Then Exit Sub
    
    ' If it's something else, tell the user what happend
    MsgBox Description, vbCritical, "Client Winsock Error: "
End Sub

' Split up the incoming data
Public Sub DoData(Data As String)
    Dim tempString As String    ' Temporary String
    Dim kickoff As Boolean      ' See's if the server wants to kick us off
    Dim aA() As String          ' Puts all of the messages in here
    Dim toPage As String        ' HTML syntax to go to the page
    Dim newHead As String       ' New header
    Dim newBody As String       ' New body (white background, black text)
    
    kickoff = False     ' Make sure it doesn't accidentally kick us off
    
    tempString = Data   ' Get the data
    
    aA = Split(tempString, SeperatorChar)   ' Split up the data with the seporator and put it into
                                            ' different slots in the array
    
    BKMK = BKMK + 1 ' Add 1 to the bookmark numbers (so that it will go there on the page)
    
    SaveSetting "Chatterbox", "Bkmk", "Bkmk", BKMK  ' Save the bookmark number
    
    DoEvents    ' Let Windows do it's thing

    ' Process the text
    Select Case aA(0)
        ' If it's a public message...
        Case "PUBL"
            ' If I sent it, make the name blue, otherwise make it red
            If aA(1) = myName Then toPage = toPage & "<B><Font Color=" & Quote & "Blue" & Quote & ">" & myName & ": </B></Font>" Else toPage = toPage & "<B><Font Color=" & Quote & "Red" & Quote & ">" & aA(1) & ": </B></FONT>"
            
            ' Put the text in and add a line
            toPage = toPage & aA(2) & "<BR>"
            
            ' Bookmark it!
            toPage = toPage & "<A id=bot" & BKMK & " name=bot" & BKMK & " href=" & Quote & "#bot" & BKMK & Quote & "></A>"
        
        ' If it's a private message...
        Case "PRIV"
            ' If the name it's being sent to is mine...
            If aA(1) = LCase$(myName) Then
                ' If it's from me, to me, make the name blue
                If aA(2) = myName Then
                    toPage = "<B><Font Color=" & Quote & "Blue" & Quote & ">" & myName & " (private): </B></Font>"
                ' Otherwise, make it red
                Else
                    toPage = "<B><Font Color=" & Quote & "Red" & Quote & ">" & aA(2) & " (private): </B></FONT>"
                End If
                
                ' Make sure it's on the next line & put the text in
                toPage = toPage & aA(3) & "<BR>"
                
                ' Bookmark it!
                toPage = toPage & "<A id=bot" & BKMK & " name=bot" & BKMK & " href=" & Quote & "#bot" & BKMK & Quote & "></A>"
            
            ' If it's not to me, don't put anything!
            Else
                Exit Sub
            End If
        
        ' If it's an image...
        Case "IMG"
            ' Find the source name and insert it
            
            ' NOTE - The actual image is not sent over the network, just a
            ' number representing one of the images that is (supposed to be)
            ' stored in the application path
            toPage = "<IMG SRC=" & Quote & App.Path & "\img" & aA(1) & ".gif" & Quote & "><BR>"
            
            ' Bookmark it!
            toPage = toPage & "<A id=bot" & BKMK & " name=bot" & BKMK & " href=" & Quote & "#bot" & BKMK & Quote & "></A>"
        
        ' If the server wants to kick us off :(
        Case "KICKOFF"
            
            ' See if the IP it wants to kick off is the same as ours
            If aA(1) = WS.LocalIP Then
                ' Tell us what happened
                MsgBox "You have been kicked off by the server" & vbNewLine & "for various reasons.", vbOKOnly + vbCritical, "Bye!"
                ' Awww
                kickoff = True
            End If
            ' Skip all the posting stuff
            GoTo 50
    End Select
    
    ' Rewrite the header with the new bookmark
    newHead = "<HTML><HEAD><TITLE>AVIN Chatterbox Chat</TITLE>" & vbNewLine & _
              "<SCRIPT language=" & Quote & "vbscript" & Quote & ">" & vbNewLine & _
              "sub tobot()" & vbNewLine & _
              "  bot" & BKMK & ".click" & vbNewLine & _
              "end sub" & vbNewLine & _
              "</SCRIPT></HEAD>"
    
    Dim temp As String  ' Temporary String
    
    ' Open the HTML file for reading
    Open "C:\Windows\Temp\avchft.html" For Input As #1
        ' Get through the stuff I don't want (the header)
        Line Input #1, temp
        Line Input #1, temp
        Line Input #1, temp
        Line Input #1, temp
        Line Input #1, temp
        Line Input #1, temp
        
        ' Copy the current conversation
        Do
            Line Input #1, temp ' Get a message
            newBody = newBody & temp & vbNewLine    ' Add it to the list
        Loop Until EOF(1) ' Go to the end of the file
        
        newBody = newBody & toPage  ' Add all of the messages & the new message
    Close #1
    
    DoEvents    ' Let Windows do it's thing
    
    ' Open the HTML file for writing
    Open "C:\Windows\Temp\avchft.html" For Output As #1
        ' Print the header
        Print #1, newHead
        ' Pring the body (white background & black text)
        Print #1, newBody
    Close #1
    
    DoEvents    ' Let Windows do it's thing again
    
    ' See if the server wants to kick me off, if so click the X button
50  If kickoff = True Then Command4_Click
    
    ' Otherwise, find the HTML file and load it
    Web.Navigate "C:\windows\temp\avchft.html"
    
    ' Play a ding sound
    PlaySound 1
End Sub
