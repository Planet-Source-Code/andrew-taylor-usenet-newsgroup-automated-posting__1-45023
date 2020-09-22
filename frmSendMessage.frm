VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSendMessage 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sending Message"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Cancel Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      Height          =   1725
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   6135
   End
   Begin MSWinsockLib.Winsock WS1 
      Left            =   30
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'*              Andrew's USENET Poster            *
'**************************************************
'*                                                *
'* This application is a fully functional, but    *
'* not entirely bug free.  It has been provided   *
'* free of charge to you for the purpose of       *
'* of education.  Please use it wisely.  I do ask *
'* that you please do not attempt to market the   *
'* program as is.  If you wish to change it       *
'* drastically, and then market it, by all means  *
'* go for it.  But please make sure that it won't *
'* look or act the same.  This comes from a       *
'* commercial application that I am developing    *
'* and I don't want to get in trouble. :)         *
'*                                                *
'* Should you have any questions about the app    *
'* you can feel free to contact me at:            *
'* lazrbrain@hotmail.com                          *
'*                                                *
'* Your comments are greatly appreciated.         *
'*                                                *
'* (c) 2003 Andrew Taylor                         *
'*                                                *
'**************************************************

Public ServerAddress As String, Username As String, Password As String, ServerPort As Integer, Post As Integer
Public Sender As String, SenderEmail As String
Public Subject As String, Message As String
Private WSMessage As String

Public Function SendMessage(PostToGroups As String, FileAttachments As String) As Integer
    
    Sending = 1
    SendMessage = 0

    If ServerAddress = "" Then
        A = MsgBox("No server selected, please select a server for the message you wish to send.  If you have not created a server listing, do so under the File Menu.", vbOKOnly, "No Server Selected")
        Sending = 0
        Unload Me
    End If

    Dim NewsGroups() As String
    Dim Attachments() As String

    NewsGroups = Split(PostToGroups, "|")
    Attachments = Split(FileAttachments, "|")
    
    txtStatus.Text = ""
    
    If ConnectToServer = 0 Then
        
        MsgBox "Unable to connect to server."
        Unload Me
        Exit Function
    
    End If

    PostingTo = 0
    
    txtStatus.Text = txtStatus.Text & "Encoding Attachments" & vbCrLf
    
    DoEvents
   
    AttachCode = ""
    
    For T = 0 To UBound(Attachments)
    
        If Attachments(T) <> "" Then
        
            AttachCode = AttachCode & vbCrLf & vbCrLf & UUEncodeFile(Attachments(T))
            DoEvents
            
        Else
        
            Exit For
            
        End If
        
    Next T
    
    FinishedPosting = 0
    
    Do While FinishedPosting = 0
    
        WSMessage = ""
    
        WS1.SendData "post" & vbCrLf
        
        Do Until WSMessage <> ""
        
            DoEvents
            
        Loop
        
        WSMessage = ""
    
        WS1.SendData "from: " & SenderEmail & " (" & Sender & ")" & vbCrLf
        
        Groups = ""
        
        For T = PostingTo To (PostingTo + Post) - 1
        
            If T > UBound(NewsGroups) Then
                
                FinishedPosting = 1
                Exit For
            
            End If
        
            If NewsGroups(T) <> "" Then
            
                If Groups <> "" Then Groups = Groups & ", "
                Groups = Groups & NewsGroups(T)
                
            Else
            
                FinishedPosting = 1
            
            End If
            
        Next T
            
        PostingTo = T
            
        Debug.Print Groups
        If Len(txtStatus.Text) + Len("Posting to: " & Groups & vbCrLf) > 50000 Then
            txtStatus.Text = Right(txtStatus.Text, Len(txtStatus.Text) - Len("Posting to: " & Groups & vbCrLf)) & "Posting to: " & Groups & vbCrLf
        Else
            txtStatus.Text = txtStatus.Text & "Posting to: " & Groups & vbCrLf
        End If
        
        txtStatus.SelStart = Len(txtStatus.Text)
        
        DoEvents
    
        WS1.SendData "newsgroups: " & Groups & vbCrLf
        
        DoEvents
        
        WS1.SendData "subject: " & Subject & " (" & PostingTo & ")" & vbCrLf
        
        DoEvents
        
        WS1.SendData vbCrLf
        
        DoEvents
        
        WS1.SendData Message & vbCrLf & vbCrLf
        
        DoEvents
        
        WS1.SendData AttachCode & vbCrLf
        
        DoEvents
        
        WS1.SendData vbCrLf
        
        DoEvents
        
        'WMMessage = ""
        
        WS1.SendData "."
        
        DoEvents
        
        WS1.SendData vbCrLf
        
        DoEvents
        
        Do Until WSMessage <> ""
        
            DoEvents
            
        Loop
        
        
    
    Loop
    
    WS1.Close
    
    Sending = 0
    SendMessage = 1
    
    Unload Me

End Function

Private Function ConnectToServer() As Integer

    WSMessage = ""

    WS1.RemoteHost = ServerAddress
    WS1.RemotePort = ServerPort
    
    txtStatus.Text = txtStatus.Text & "Connecting to " & ServerAddress & vbCrLf
    
    DoEvents
    
    ConnectState = WS1.State
    
    WS1.Connect
    
    Do Until WS1.State = 7 Or WS1.State = 9 Or WS1.State = 0
    
        DoEvents
    
        If ConnectState <> WS1.State Then
        
            Select Case WS1.State
            
                Case 1
                
                    txtStatus.Text = txtStatus.Text & "Opening Port" & vbCrLf
                
                Case 3
                    
                    txtStatus.Text = txtStatus.Text & "Connection Pending" & vbCrLf
                
                Case 4
                
                    txtStatus.Text = txtStatus.Text & "Resolving Host" & vbCrLf
                
                Case 5
                
                    txtStatus.Text = txtStatus.Text & "Host Resolved" & vbCrLf
                
                Case 6
                
                    txtStatus.Text = txtStatus.Text & "Connecting..." & vbCrLf
                
                Case 7
                
                    txtStatus.Text = txtStatus.Text & "Connection Established" & vbCrLf
                
                Case 8
                
                    txtStatus.Text = txtStatus.Text & "Peer Closing Connection" & vbCrLf
                    WS1.Close
                    ConnectToServer = 0
                    Exit Function
                
                Case 9
                
                    txtStatus.Text = txtStatus.Text & "Error Establishing Connection" & vbCrLf
                    WS1.Close
                    ConnectToServer = 0
                    Exit Function
                
            End Select
            
            ConnectState = WS1.State
            
            DoEvents
        
        End If

    Loop
    
    If WS1.State = 7 Then
        
        'WS1.SendData vbCrLf

        Do Until WSMessage <> ""
    
            DoEvents
            
        Loop
    
        If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "200" Then
        
            txtStatus.Text = txtStatus.Text & "Authenticating..." & vbCrLf
            
        Else
        
            txtStatus.Text = txtStatus.Text & "Server Not Responding..." & vbCrLf
            WS1.Close
            
        End If
    
        WSMessage = ""
        
        StartTimer = Time
        
        Do While DateDiff("s", StartTime, Time) < 3
            DoEvents
        Loop
        
        WS1.SendData "post" & vbCrLf
    
        Do While WSMessage = ""
        
            DoEvents
        
        Loop
        
        If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "480" Then
        
            WSMessage = ""
        
            WS1.SendData "authinfo user " & Username & vbCrLf
            
            Do While WSMessage = ""
            
                DoEvents
            
            Loop
            
            WSMessage = ""
        
            WS1.SendData "authinfo pass " & Password & vbCrLf
            
            Do While WSMessage = ""
            
                DoEvents
            
            Loop
        
            If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "502" Then
                
                txtStatus.Text = txtStatus.Text & "Invalid Username/Password" & vbCrLf
                ConnectToServer = 0
                WS1.Close
                Exit Function
            
            ElseIf Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "281" Then
            
                txtStatus.Text = txtStatus.Text & "Authentication Successful" & vbCrLf
                ConnectToServer = 1
                'WS1.Close
                Exit Function
            
            Else
        
                txtStatus.Text = txtStatus.Text & "Server Error" & vbCrLf
                ConnectToServer = 0
                WS1.Close
                Exit Function
        
            End If
        
        Else
    
            ConnectToServer = 1
        
            WS1.SendData vbCrLf & "." & vbCrLf
        
            'WS1.Close
        
            txtStatus.Text = txtStatus.Text & "Connected Successfully" & vbCrLf
    
        End If

    Else
    
        WS1.Close
        ConnectToServer = 0
        Sending = 0
        txtStatus.Text = txtStatus.Text & "Unable to connect" & vbCrLf
        Unload Me
        Exit Function
        
    End If


End Function


Private Sub cmdCancel_Click()

    WS1.Close
    Sending = 0
    Unload Me

End Sub

Private Sub Form_Load()

    Me.Top = (frmMain.Height / 2) - (Me.Height / 2)
    Me.Left = (frmMain.Width / 2) - (Me.Width / 2)

End Sub

Public Function UUEncodeFile(strFilePath As String) As String

Dim intFile As Integer     'file handler
Dim intTempFile As Integer 'temp file
Dim lFileSize As Long      'size of the file
Dim strFileName As String  'name of the file
Dim strFileData As String  'file data chunk
Dim lEncodedLines As Long  'number of encoded lines
Dim strTempLine As String  'temporary string
Dim i As Long              'loop counter
Dim j As Integer           'loop counter

Dim strResult As String
'
'Get file name
strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
'
'Insert first marker: "begin 664 ..."
strResult = "begin 664 " + strFileName + vbLf
'
'Get file size
lFileSize = FileLen(strFilePath)
lEncodedLines = lFileSize \ 45 + 1
'
'Prepare buffer to retrieve data from
'the file by 45 symbols chunks
strFileData = Space(45)
'
intFile = FreeFile
'
Open strFilePath For Binary As intFile
For i = 1 To lEncodedLines
    'Read file data by 45-bytes cnunks
    '
    If i = lEncodedLines Then
        'Last line of encoded data often is not
        'equal to 45, therefore we need to change
        'size of the buffer
        strFileData = Space(lFileSize Mod 45)
    End If
    'Retrieve data chunk from file to the buffer
Get intFile, , strFileData
    'Add first symbol to encoded string that informs
    'about quantity of symbols in encoded string.
    'More often "M" symbol is used.
strTempLine = Chr(Len(strFileData) + 32)
    '
    If i = lEncodedLines And (Len(strFileData) Mod 3) Then
        'If the last line is processed and length of
        'source data is not a number divisible by 3,
        'add one or two blankspace symbols
        strFileData = strFileData + Space(3 - _
            (Len(strFileData) Mod 3))
    End If

    For j = 1 To Len(strFileData) Step 3
        'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
        '
        '1 byte
        strTempLine = strTempLine + _
            Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
        '2 byte
        strTempLine = strTempLine + _
            Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
            + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
        '3 byte
        strTempLine = strTempLine + _
            Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
            + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
        '4 byte
        strTempLine = strTempLine + _
            Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
    Next j
    'add encoded line to result buffer
    strResult = strResult + strTempLine + vbLf
    'reset line buffer
    strTempLine = ""
Next i
Close intFile
'add the end marker
strResult = strResult & "'" & vbLf + "end" + vbLf
'asign return value
UUEncodeFile = strResult

End Function

Private Sub Form_Unload(Cancel As Integer)

If WS1.State <> 0 Then
    WS1.Close
End If

End Sub

Private Sub WS1_DataArrival(ByVal bytesTotal As Long)

Call WS1.GetData(Incoming, vbString, bytesTotal)

If Len(txtStatus.Text) + Len(Incoming) > 50000 Then
    txtStatus.Text = Right(txtStatus.Text, Len(txtStatus.Text) - Len(Incoming)) & Incoming
Else
    txtStatus.Text = txtStatus.Text & Incoming
End If

txtStatus.SelStart = Len(txtStatus)

WSMessage = Incoming

End Sub

