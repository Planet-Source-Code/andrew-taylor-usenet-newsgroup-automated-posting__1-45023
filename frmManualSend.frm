VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmManualSend 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manually Post Message"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9120
   Begin VB.ComboBox cmbGroupListing 
      Height          =   315
      Left            =   4590
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   270
      Width           =   4455
   End
   Begin VB.TextBox txtSendereMail 
      Height          =   315
      Left            =   4590
      TabIndex        =   12
      Top             =   900
      Width           =   4455
   End
   Begin VB.TextBox txtSender 
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   900
      Width           =   4485
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2010
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Send Message"
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
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1515
   End
   Begin VB.CommandButton cmdAttach 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Attach File"
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
      TabIndex        =   8
      Top             =   5280
      Width           =   1515
   End
   Begin VB.ListBox lstAttachments 
      Height          =   960
      IntegralHeight  =   0   'False
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   4260
      Width           =   8985
   End
   Begin VB.TextBox txtMessage 
      Height          =   1845
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   2130
      Width           =   8985
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1530
      Width           =   8985
   End
   Begin VB.ComboBox cmbServerList 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   4485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Post To Group Listing:"
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
      Left            =   4590
      TabIndex        =   15
      Top             =   60
      Width           =   1920
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender e-Mail Address:"
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
      Left            =   4590
      TabIndex        =   13
      Top             =   690
      Width           =   1980
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Name:"
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
      TabIndex        =   11
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attachments:"
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
      Top             =   4050
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
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
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
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
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server To Post To:"
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
      TabIndex        =   1
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "frmManualSend"
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


Private ServerNames() As String

Private Sub GetGroupList()

'Get a directory listing of the saved groups files.

cmbGroupListing.Clear

Group = Dir(App.Path & "\savedlist\*.lst", vbNormal)

Do While Group <> ""

    cmbGroupListing.AddItem Left(Group, Len(Group) - 4)
    
    Group = Dir

Loop



End Sub

Private Sub RefreshServerList()

    'Some errors are handled special
    On Error GoTo RefreshServerListError

    'Temp holding places.
    Dim ServerNameList(65000, 6) As String
    Dim ServerInfo() As String
    
    'Open up our server listing
    Open App.Path & "\servers.dat" For Input As #1
    
    ServerCount = 0
    
    'Populate the array of server names.
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
    
    'Transfer server names to form level variable
    ReDim ServerNames(ServerCount, 6) As String
    
    For T = 0 To ServerCount
    
        ServerNames(T, 1) = ServerNameList(T, 1)
        ServerNames(T, 2) = ServerNameList(T, 2)
        ServerNames(T, 3) = ServerNameList(T, 3)
        ServerNames(T, 4) = ServerNameList(T, 4)
        ServerNames(T, 5) = ServerNameList(T, 5)
        ServerNames(T, 6) = ServerNameList(T, 6)
        
        
    Next T
    
    'Populate the server names combo box
    cmbServerList.Clear
    
    For T = 0 To UBound(ServerNames)
        
        If ServerNames(T, 1) <> "" Then
        
            cmbServerList.AddItem ServerNames(T, 1)
            
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

Private Sub cmdAttach_Click()

    'Open a common dialog to get the file they wish to
    'attach.
    CD1.DialogTitle = "Attach File"
    CD1.DefaultExt = "*.*"
    CD1.Filter = "All Files|*.*"
    CD1.Action = 1
    
    'Make sure the file isn't already attached.
    FileExist = 0
    
    For T = 0 To lstAttachments.ListCount - 1
    
        lstAttachments.ListIndex = T
        
        If lstAttachments.Text = CD1.FileName Then FileExist = 1
        
    Next T
    
    If FileExist = 0 Then lstAttachments.AddItem CD1.FileName

End Sub


Private Sub cmdSend_Click()

Dim Attachments As String
Dim NewsGroups As String

Attachments = ""
NewsGroups = ""

'Now we need to make a list of groups to send it to.
Open App.Path & "\savedlist\" & cmbGroupListing.Text & ".lst" For Input As #1

'We can't pass an array to a function, and I didn't want to make a global,
'so instead we pass it a string and then split it at the other end.
Do While Not EOF(1)

    Line Input #1, NewsGroup

    If NewsGroups <> "" Then NewsGroups = NewsGroups & "|"
    NewsGroups = NewsGroups & NewsGroup

Loop

Close #1

'Do the same with the list of attachments.
For T = 0 To lstAttachments.ListCount - 1

    lstAttachments.ListIndex = T
    
    If Attachments <> "" Then Attachments = Attachments & "|"
    Attachments = Attachments & lstAttachments.Text
    
Next T

If NewsGroups = "" Then

    MsgBox "You must select some newsgroups to post to."
    
    Exit Sub

End If

If txtSender.Text = "" Then

    A = MsgBox("You must enter a senders name before posting a message.", vbOKOnly, "Sender Required")
    Exit Sub
    
End If

If txtSendereMail.Text = "" Then

    A = MsgBox("You must enter a senders e-mail address before posting a message.", vbOKOnly, "Sender Required")
    Exit Sub
    
End If

If txtSubject.Text = "" Then

    A = MsgBox("You must enter a subject for your post.", vbOKOnly, "Subject Required")
    Exit Sub
    
End If

If txtMessage.Text = "" And Attachments = "" Then

    A = MsgBox("You can't send a message without a message, please enter a message.", vbOKOnly, "Message Required")
    Exit Sub

End If
    
'Find the server we have selected in the server array.
For T = 0 To UBound(ServerNames)

    If ServerNames(T, 1) = cmbServerList.Text Then
        Exit For
    End If
    
Next T

If ServerNames(T, 1) <> "" Then

frmSendMessage.ServerAddress = ServerNames(T, 2)
frmSendMessage.Username = ServerNames(T, 3)
frmSendMessage.Password = ServerNames(T, 4)
frmSendMessage.ServerPort = ServerNames(T, 5)
frmSendMessage.Post = ServerNames(T, 6)

frmSendMessage.Sender = txtSender.Text
frmSendMessage.SenderEmail = txtSendereMail.Text
frmSendMessage.Subject = txtSubject.Text
frmSendMessage.Message = txtMessage.Text

frmSendMessage.Show

Call frmSendMessage.SendMessage(NewsGroups, Attachments)

Else

    A = MsgBox("You must select a server to post to, if no servers are listed, go to the manage server menu item under the file menu.", vbOKOnly, "No Server Selected")

End If

End Sub

Private Sub Form_Load()

    'Populate The Server List
    RefreshServerList
    GetGroupList

End Sub

Private Sub txtListSort_Change()

End Sub


Private Sub lstAttachments_DblClick()

lstAttachments.RemoveItem lstAttachments.ListIndex

End Sub

