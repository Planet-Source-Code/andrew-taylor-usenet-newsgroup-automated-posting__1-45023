VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmManageServers 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Servers"
   ClientHeight    =   4245
   ClientLeft      =   2610
   ClientTop       =   2865
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6930
   Begin VB.TextBox txtPort 
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   15
      Top             =   2220
      Width           =   1935
   End
   Begin VB.TextBox txtPost 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2190
      TabIndex        =   14
      Top             =   2220
      Width           =   1935
   End
   Begin VB.TextBox txtStatus 
      Height          =   1065
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3120
      Width           =   4065
   End
   Begin MSWinsockLib.Winsock WS1 
      Left            =   3720
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEditServer 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Edit Server"
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
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2670
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdDeleteServer 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Delete"
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
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2670
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveChanges 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Save"
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
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2670
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdAddServer 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Add Server"
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2670
      Width           =   1275
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2190
      TabIndex        =   7
      Top             =   1590
      Width           =   1935
   End
   Begin VB.TextBox txtUsername 
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   1590
      Width           =   1935
   End
   Begin VB.TextBox txtServerAddress 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   930
      Width           =   4065
   End
   Begin VB.TextBox txtServerName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   270
      Width           =   4065
   End
   Begin VB.ListBox lstServers 
      Height          =   4155
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   2685
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   2010
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Max Cross Posts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2190
      TabIndex        =   16
      Top             =   2010
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2190
      TabIndex        =   8
      Top             =   1380
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1170
   End
End
Attribute VB_Name = "frmManageServers"
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


Private intEditStatus As Integer

Private ServerNames() As String
Private ModifyServer As String
Private WSMessage As String

Private Function CheckConnection(ServerAddress As String, Username As String, Password As String, ServerPort As Integer) As Integer

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
                    CheckConnection = 0
                    Exit Function
                
                Case 9
                
                    txtStatus.Text = txtStatus.Text & "Error Establishing Connection" & vbCrLf
                    WS1.Close
                    CheckConnection = 0
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
        
        WS1.SendData "Post" & vbCrLf
    
        Do While WSMessage = ""
        
            DoEvents
        
        Loop
    
        If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "480" Then
        
            WSMessage = ""
        
            WS1.SendData "authinfo user " & txtUsername.Text & vbCrLf
            
            Do While WSMessage = ""
            
                DoEvents
            
            Loop
            
            WSMessage = ""
        
            WS1.SendData "authinfo pass " & txtPassword.Text & vbCrLf
            
            Do While WSMessage = ""
            
                DoEvents
            
            Loop
        
            If Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "502" Then
                
                txtStatus.Text = txtStatus.Text & "Invalid Username/Password" & vbCrLf
                CheckConnection = 0
                WS1.Close
                Exit Function
            
            ElseIf Left(WSMessage, InStr(1, WSMessage, " ", vbTextCompare) - 1) = "281" Then
            
                txtStatus.Text = txtStatus.Text & "Authentication Successful"
                CheckConnection = 1
                WS1.Close
                Exit Function
            
            Else
        
                txtStatus.Text = txtStatus.Text & "Server Error" & vbCrLf
                CheckConnection = 0
                WS1.Close
                Exit Function
        
            End If
        
        Else
    
            CheckConnection = 1
        
            WS1.Close
        
            txtStatus.Text = txtStatus.Text & "Connected Successfully" & vbCrLf
    
        End If

    Else
    
        WS1.Close
        CheckConnection = 0
        txtStatus.Text = txtStatus.Text & "Unable to connect" & vbCrLf
        Exit Function
        
    End If





End Function

Private Sub SaveServerList()

    Open App.Path & "\servers.dat" For Output As #1
    
        For T = 0 To UBound(ServerNames)
        
            If ServerNames(T, 1) <> "" Then
            
                Print #1, ServerNames(T, 1) & "|" & ServerNames(T, 2) & "|" & ServerNames(T, 3) & "|" & ServerNames(T, 4) & "|" & ServerNames(T, 5) & "|" & ServerNames(T, 6)
                
            End If

        Next T
        
    Close #1

End Sub


Private Sub RefreshServerList()

   ' On Error GoTo RefreshServerListError

    Dim ServerNameList(65000, 6) As String
    Dim ServerInfo() As String
    
    lstServers.Clear
    
    Open App.Path & "\servers.dat" For Input As #1
    
    ServerCount = 0
    
    Do While Not EOF(1)
    
        Line Input #1, ServerLine
        
        ServerInfo = Split(ServerLine, "|", , vbBinaryCompare)
        
        ServerNameList(ServerCount, 1) = ServerInfo(0)
        ServerNameList(ServerCount, 2) = ServerInfo(1)
        ServerNameList(ServerCount, 3) = ServerInfo(2)
        ServerNameList(ServerCount, 4) = ServerInfo(3)
        ServerNameList(ServerCount, 5) = ServerInfo(4)
        ServerNameList(ServerCount, 6) = ServerInfo(5)
        
        ServerCount = ServerCount + 1
        
    Loop
    
    Close #1
    
    ReDim ServerNames(ServerCount, 6) As String
    
    For T = 0 To ServerCount
    
        ServerNames(T, 1) = ServerNameList(T, 1)
        ServerNames(T, 2) = ServerNameList(T, 2)
        ServerNames(T, 3) = ServerNameList(T, 3)
        ServerNames(T, 4) = ServerNameList(T, 4)
        ServerNames(T, 5) = ServerNameList(T, 5)
        ServerNames(T, 6) = ServerNameList(T, 6)
        
        
    Next T
    
    For T = 0 To UBound(ServerNames)
        
        If ServerNames(T, 1) <> "" Then
        
            lstServers.AddItem ServerNames(T, 1)
            
        End If
        
    Next T
    
    Exit Sub
    
RefreshServerListError:

    Select Case Err.Number
    
        Case 53
    
            Open App.Path & "\server.dat" For Output As #1
            Close #1
            ReDim ServerNames(0, 6) As String
            Exit Sub
    
        Case Else
        
            A = MsgBox("Unable to load server listing." & vbCrLf & vbCrLf & "Error: " & Err.Number & " | " & Err.Description, vbOKOnly, "Manage Servers")
            
            Exit Sub
            
    End Select
    
End Sub


Private Sub cmdAddServer_Click()

    intEditStatus = 1
    
    txtServerName.Enabled = True
    txtServerAddress.Enabled = True
    txtUsername.Enabled = True
    txtPassword.Enabled = True
    txtPort.Enabled = True
    txtPost.Enabled = True
    txtServerName.Text = ""
    txtServerAddress.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtPort.Text = "119"
    txtPost.Text = "10"
    cmdAddServer.Visible = False
    cmdEditServer.Visible = False
    cmdDeleteServer.Visible = False
    cmdSaveChanges.Visible = True
    lstServers.Enabled = True
    txtServerName.SetFocus

End Sub

Private Sub cmdDeleteServer_Click()

    intEditStatus = 0

    'Find Server In List
    For T = 0 To UBound(ServerNames)
    
        If ServerNames(T, 1) = txtServerName.Text Then
            
            ServerNames(T, 1) = ""
            ServerNames(T, 2) = ""
            ServerNames(T, 3) = ""
            ServerNames(T, 4) = ""
            ServerNames(T, 5) = ""
            ServerNames(T, 6) = ""
            
            
            Exit For
            
        End If
        
    Next T
    
    Open App.Path & "\servers\" & txtServerName.Text & ".lst" For Output As #1
        Print #1, vbCrLf
    Close #1
    
    Kill App.Path & "\servers\" & txtServerName.Text & ".lst"
    
    Call SaveServerList
    Call RefreshServerList
    
    txtServerName.Enabled = False
    txtServerAddress.Enabled = False
    txtUsername.Enabled = False
    txtPassword.Enabled = False
    txtPort.Enabled = False
    txtPost.Enabled = False
    txtServerName.Text = ""
    txtServerAddress.Text = ""
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtPost.Text = ""
    txtPort.Text = ""
    cmdAddServer.Visible = True
    cmdEditServer.Visible = False
    cmdDeleteServer.Visible = False
    cmdSaveChanges.Visible = False
    lstServers.Enabled = True
    lstServers.SetFocus
    
    
End Sub

Private Sub cmdEditServer_Click()

    intEditStatus = 2

    ModifyServer = txtServerName.Text

    txtServerName.Enabled = True
    txtServerAddress.Enabled = True
    txtUsername.Enabled = True
    txtPassword.Enabled = True
    txtPort.Enabled = True
    txtPost.Enabled = True
    cmdAddServer.Visible = False
    cmdSaveChanges.Visible = True
    cmdEditServer.Visible = False
    cmdDeleteServer.Visible = False
    lstServers.Enabled = True
    txtServerName.SetFocus

End Sub

Private Sub cmdSaveChanges_Click()

Select Case intEditStatus

    Case 1 'Adding Server
    
        'Check to make sure all information has been input
        If txtServerName.Text = "" Then
            
            txtServerName.Text = txtServerAddress.Text
        
        End If
        
        'If no server address, inform user
        If txtServerAddress.Text = "" Then
            
            A = MsgBox("You must enter an address for the server you wish to connect to, generally these look like 'news.myserver.com' or '1.2.3.4'.", vbInformation, "Missing Server Address")
    
        End If
        
        'Check Server List For Existing Server With This Name
        
        ServerExist = 0
        
        For T = 0 To UBound(ServerNames)
        
            If UCase(ServerNames(T, 1)) = UCase(txtServerName.Text) Then
                ServerExist = 1
                Exit For
            End If
        
        Next T
        
        'Check to see if you can connect
        
        txtStatus.Text = ""
        txtStatus.Visible = True
        
        ValidServer = CheckConnection(txtServerAddress.Text, txtUsername.Text, txtPassword.Text, txtPort.Text)
        
        
        If ValidServer = 0 Then Exit Sub
        
        
        'If the server does exist, add it.
        
        If ServerExist = 0 Then
        
            Open App.Path & "\servers.dat" For Append As #1
            
            Print #1, txtServerName.Text & "|" & txtServerAddress.Text & "|" & txtUsername.Text & "|" & txtPassword.Text & "|" & txtPort.Text & "|" & txtPost.Text
    
            Close #1
            
        End If
    
        intEditStatus = 0
       
        frmGetNewsGroups.ServerAddress = txtServerAddress.Text
        frmGetNewsGroups.ServerName = txtServerName.Text
        frmGetNewsGroups.Username = txtUsername.Text
        frmGetNewsGroups.Password = txtPassword.Text
        frmGetNewsGroups.ServerPort = txtPort.Text
        
        frmGetNewsGroups.Show
        frmGetNewsGroups.GetNewsgroupListing
        Unload frmGetNewsGroups
     
        txtServerName.Enabled = False
        txtServerAddress.Enabled = False
        txtUsername.Enabled = False
        txtPassword.Enabled = False
        txtPort.Enabled = False
        txtPost.Enabled = False
        txtServerName.Text = ""
        txtServerAddress.Text = ""
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtPort.Text = ""
        txtPost.Text = ""
        cmdAddServer.Visible = True
        cmdEditServer.Visible = False
        cmdDeleteServer.Visible = False
        cmdSaveChanges.Visible = False
        lstServers.Enabled = True
        lstServers.SetFocus
        
        Call RefreshServerList
        
    
    Case 2
    
        'Check to make sure all information has been input
        If txtServerName.Text = "" Then
            
            txtServerName.Text = txtServerAddress.Text
        
        End If
        
        'If no server address, inform user
        If txtServerAddress.Text = "" Then
            
            A = MsgBox("You must enter an address for the server you wish to connect to, generally these look like 'news.myserver.com' or '1.2.3.4'.", vbInformation, "Missing Server Address")
    
        End If
        
        For T = 0 To UBound(ServerNames)
        
            If ServerNames(T, 1) = ModifyServer Then
                
                ServerNames(T, 1) = ""
                ServerNames(T, 2) = ""
                ServerNames(T, 3) = ""
                ServerNames(T, 4) = ""
                ServerNames(T, 5) = ""
                ServerNames(T, 6) = ""
                
                
                Exit For
                
            End If
            
        Next T
        
        'Call cmdSaveServer_click
        
        'Check Server List For Existing Server With This Name
        
        ServerExist = 0
        
        For T = 0 To UBound(ServerNames)
        
            If UCase(ServerNames(T, 1)) = UCase(txtServerName.Text) Then
                ServerExist = 1
                Exit For
            End If
        
        Next T
        
        'Check to see if you can connect
        
        txtStatus.Text = ""
        txtStatus.Visible = True
        
        ValidServer = CheckConnection(txtServerAddress.Text, txtUsername.Text, txtPassword.Text, txtPort.Text)
        
        
        If ValidServer = 0 Then Exit Sub
        
        
        'If the server does exist, add it.
        
        If ServerExist = 0 Then
        
            Open App.Path & "\servers.dat" For Append As #1
            
            Print #1, txtServerName.Text & "|" & txtServerAddress.Text & "|" & txtUsername.Text & "|" & txtPassword.Text & "|" & txtPort.Text & "|" & txtPost.Text
    
            Close #1
            
        End If
    
        intEditStatus = 0
       
        txtServerName.Enabled = False
        txtServerAddress.Enabled = False
        txtUsername.Enabled = False
        txtPassword.Enabled = False
        txtPort.Enabled = False
        txtPost.Enabled = False
        txtServerName.Text = ""
        txtServerAddress.Text = ""
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtPort.Text = ""
        txtPost.Text = ""
        cmdAddServer.Visible = True
        cmdEditServer.Visible = False
        cmdDeleteServer.Visible = False
        cmdSaveChanges.Visible = False
        lstServers.Enabled = True
        lstServers.SetFocus
        
        Call RefreshServerList
    
    
End Select

End Sub

Private Sub Form_Load()

    intEditStatus = 0

    Call RefreshServerList

End Sub

Private Sub Form_Unload(Cancel As Integer)

    WS1.Close

End Sub

Private Sub lstServers_DblClick()

    intEditStatus = 0

    'Find Server In List
    For T = 0 To UBound(ServerNames)
    
        If ServerNames(T, 1) = lstServers.Text Then
            Exit For
        End If
        
    Next T

    txtServerName.Enabled = False
    txtServerAddress.Enabled = False
    txtUsername.Enabled = False
    txtPassword.Enabled = False
    txtPort.Enabled = False
    txtPost.Enabled = False
    txtServerName.Text = ServerNames(T, 1)
    txtServerAddress.Text = ServerNames(T, 2)
    txtUsername.Text = ServerNames(T, 3)
    txtPassword.Text = ServerNames(T, 4)
    txtPort.Text = ServerNames(T, 5)
    txtPost.Text = ServerNames(T, 6)
    cmdAddServer.Visible = True
    cmdEditServer.Visible = True
    cmdDeleteServer.Visible = True
    cmdSaveChanges.Visible = False
    txtStatus.Visible = False
    lstServers.Enabled = True
    lstServers.SetFocus

End Sub

Private Sub WS1_DataArrival(ByVal bytesTotal As Long)

    Call WS1.GetData(WSMessage, vbString, bytesTotal)
    
    txtStatus.Text = txtStatus.Text & WSMessage & vbCrLf

End Sub

Private Sub WS1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    A = MsgBox("Unable to connect to newsgroup server for the following reason: " & Description, vbOKOnly, "Unable To Connect")

    WS1.Close

End Sub
