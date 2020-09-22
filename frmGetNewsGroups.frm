VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmGetNewsGroups 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Server Newsgroup Listing"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock WS1 
      Left            =   30
      Top             =   1410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtStatus 
      Height          =   1305
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   5505
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1545
   End
End
Attribute VB_Name = "frmGetNewsGroups"
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


Public ServerName As String, ServerAddress As String, ServerPort As Integer, Username As String, Password As String

Private RawGroups As String, GettingListing As String
Private WSMessage As String
Private ListCountEstimate As Long
Private Groups() As String

Private Sub Heap_Sort()
    
    'I didn't write this code, but thanks to the
    'person that did.  It's the fastest one I could
    'find.  Sorting here saves lots of time when
    'loading to a list box.
    
    Dim j, tmp
    Dim Last1 As Long
    
    Last1 = UBound(Groups) - LBound(Groups)
    Max = UBound(Groups)
    Min = LBound(Groups)

    'Note: j = MAX - (MAX - MIN) \ 2
    j = Last1 - Last1 \ 2
    
    Do Until j >= Last1
        SiftUp Last1, j, 0
        j = j + 1
    Loop

    For j = 0 To Last1 - 1
        If Groups(Last1) < Groups(j) Then
            tmp = Groups(Last1)
            Groups(Last1) = Groups(j)
            Groups(j) = tmp
            
            SiftUp Last1, Last1 - 1, j
        End If
    Next
End Sub
Private Sub SiftUp(first, ByVal midl, last)
    
    'Used by Heap_sort
    
    Dim k, k1, m1, tmp
    
    k = (midl - first) * 2 + first
    Do While k >= last
        If k > last Then
            k1 = k + 1
            If Groups(k) < Groups(k1) Then
                k = k - 1
            End If
        End If
        
        k1 = k + 1
        m1 = midl + 1
        
        If Groups(k1) < Groups(m1) Then
            tmp = Groups(k1)
            Groups(k1) = Groups(m1)
            Groups(m1) = tmp
        Else
            Exit Do
        End If
        midl = k
        k = (midl - first) * 2 + first
    Loop
End Sub


Public Function GetNewsgroupListing() As Integer

    GettingListing = 0
    RawGroups = ""
    WSMessage = ""
    ListCountEstimate = 0

    txtStatus.Text = ""
    
    'If we aren't connected, connect now. :)
    
    If ConnectToServer = 0 Then
        
        MsgBox "Unable to connect to server."
        GetNewsgroupListing = 0
        Unload Me
        Exit Function
    
    End If
    
    txtStatus.Text = "Getting List of Newsgroups"
    
    GettingListing = 1
    
    WS1.SendData "list" & vbCrLf
    
    'Wait until the end of the list, designated by the CRLF.CRLF
    
    Do While InStr(1, RawGroups, vbCrLf & "." & vbCrLf, vbBinaryCompare) = 0
    
        DoEvents 'Twiddle my thumbs?
        
    Loop
    
    GettingListing = 0
    
    'Clean up the listing a little, the server will send
    'some junk back that we don't want.
    RawGroups = Right(RawGroups, Len(RawGroups) - (InStr(1, RawGroups, vbCrLf) + 1))
    RawGroups = Left(RawGroups, Len(RawGroups) - 5)
    
    'Put the listing into an array.
    Groups = Split(RawGroups, vbCrLf)
    
    'Sort the listing alphabetically
    Call Heap_Sort
    
    'Save The Listing
    Open App.Path & "\servers\" & ServerName & ".lst" For Output As #1
    
        For T = 0 To UBound(Groups)
        
            Print #1, Groups(T)
            
        Next T
        
    
    Close #1
    
    GetNewsgroupListing = 1
    
    txtStatus.Text = "List Complete"

End Function


Private Function ConnectToServer() As Integer

    WSMessage = ""

    'Globals from basGeneral
    WS1.RemoteHost = ServerAddress
    WS1.RemotePort = ServerPort
    
    txtStatus.Text = txtStatus.Text & "Connecting to " & ServerAddress & vbCrLf
    
    DoEvents
    
    ConnectState = WS1.State
    
    WS1.Connect
    
    '7 = Connected, 8 = Server Disconnect, 9 = Error, 0 = Closed
    Do Until WS1.State = 7 Or WS1.State = 9 Or WS1.State = 0
    
        DoEvents
    
        'Let the user know what's going on.
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
    
    'If we connected successfully
    
    If WS1.State = 7 Then
        
        'Get the servers attention.  Some servers need this
        'others don't.
        WS1.SendData vbCrLf

        Do Until WSMessage <> ""
    
            DoEvents
            
        Loop
    
        'Did Server say hello?
        If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "200" Then
        
            txtStatus.Text = txtStatus.Text & "Authenticating..." & vbCrLf
            
        Else
        
            txtStatus.Text = txtStatus.Text & "Server Not Responding..." & vbCrLf
            WS1.Close
            
        End If
    
        WSMessage = ""
        
        'Send bogus info to server, we want to see if
        'it wants a username or not.
        WS1.SendData "Help" & vbCrLf
    
        Do While WSMessage = ""
        
            DoEvents
        
        Loop
    
        'If it want's a user name...
        If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "480" Then
        
            WSMessage = ""
        
            WS1.SendData "authinfo user " & Username & vbCrLf
            
            'Wait for it to register the username
            Do While WSMessage = ""
            
                DoEvents
            
            Loop
            
            WSMessage = ""
        
            WS1.SendData "authinfo pass " & Password & vbCrLf
            
            Do While WSMessage = ""
            
                DoEvents
            
            Loop
        
            'Bad username/password
            If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "502" Then
                
                txtStatus.Text = txtStatus.Text & "Invalid Username/Password" & vbCrLf
                ConnectToServer = 0
                WS1.Close
                Exit Function
            
            'Where in!
            ElseIf Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "281" Then
            
                txtStatus.Text = txtStatus.Text & "Authentication Successful"
                ConnectToServer = 1
                'WS1.Close
                Exit Function
            
            Else
        
                'Ack!  What do I do?
                txtStatus.Text = txtStatus.Text & "Server Error" & vbCrLf
                ConnectToServer = 0
                WS1.Close
                Exit Function
        
            End If
        
        Else
    
            ConnectToServer = 1
        
            'WS1.Close
        
            txtStatus.Text = txtStatus.Text & "Connected Successfully" & vbCrLf
    
        End If

    Else
    
        WS1.Close
        ConnectToServer = 0
        txtStatus.Text = txtStatus.Text & "Unable to connect" & vbCrLf
        Exit Function
        
    End If


End Function


Private Sub Form_Load()

'Center the form
Me.Top = (frmMain.Height / 2) - (Me.Height / 2)
Me.Left = (frmMain.Width / 2) - (Me.Width / 2)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Make sure the WS is closed, not doing so can cause vb to hang.
If WS1.State <> 0 Then
    WS1.Close
End If

End Sub

Private Sub WS1_DataArrival(ByVal bytesTotal As Long)

If GettingListing = 0 Then
    
    'Put the incoming data in a variable so my loops
    'can read it.
    Call WS1.GetData(WSMessage, vbString, bytesTotal)
    
ElseIf GettingListing = 1 Then

    Dim ListCount() As String

    Call WS1.GetData(IncomingList, vbString, bytesTotal)
    
    RawGroups = RawGroups & IncomingList
    
    'Attempt to get a rough count of the incoming
    'listing.
    ListCount = Split(IncomingList, vbCrLf)
    
    ListCountEstimate = ListCountEstimate + UBound(ListCount)
    
    'This lets the user know they aren't just
    'sitting there.
    txtStatus.Text = "Getting Newsgroups...  Found " & ListCountEstimate

    DoEvents

End If

End Sub

