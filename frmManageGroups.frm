VERSION 5.00
Begin VB.Form frmManageGroups 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Posting Groups"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdGetGroups 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Update Server Groups"
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
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2730
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Delete List"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2730
      Width           =   1515
   End
   Begin VB.CommandButton cmdSaveList 
      BackColor       =   &H00ADDDFA&
      Caption         =   "Save List"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2730
      Width           =   1515
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00ADDDFA&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   210
      Width           =   495
   End
   Begin VB.ComboBox cmbGroupListing 
      Height          =   315
      Left            =   4530
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   210
      Width           =   3885
   End
   Begin VB.ComboBox cmbServerList 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   210
      Width           =   4485
   End
   Begin VB.ListBox lstGroups 
      Height          =   1830
      IntegralHeight  =   0   'False
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   810
      Width           =   4485
   End
   Begin VB.ListBox lstSelected 
      Height          =   1830
      IntegralHeight  =   0   'False
      Left            =   4530
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   810
      Width           =   4455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Listing Name:"
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
      Left            =   4530
      TabIndex        =   9
      Top             =   0
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server To Select Groups From:"
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
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Groups:  (Double Click to add)"
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
      Left            =   30
      TabIndex        =   6
      Top             =   600
      Width           =   3435
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Newsgroup Listing..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   30
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.Label lblNewsgroupsLoaded 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Newsgroups Loaded: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   30
      TabIndex        =   4
      Top             =   1740
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Groups:"
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
      Left            =   4530
      TabIndex        =   3
      Top             =   600
      Width           =   1485
   End
End
Attribute VB_Name = "frmManageGroups"
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
Private Groups() As String

Private Sub GetGroupList()

'Get a list of all the saved group files.
cmbGroupListing.Clear

Group = Dir(App.Path & "\savedlist\*.lst", vbNormal)

Do While Group <> ""

    'We don't want the .lst on there...
    cmbGroupListing.AddItem Left(Group, Len(Group) - 4)
    
    Group = Dir

Loop



End Sub

Private Sub Heap_Sort()
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


Private Sub LoadGroupList()

    'Get the group listing from the server...
    Dim tempGroups(2000000) As String

    lstGroups.Clear
    
    If cmbServerList.Text <> "" Then
    
        Open App.Path & "\servers\" & cmbServerList.Text & ".lst" For Input As #1
        
            PeriodCount = 0
            GroupCount = 0
        
            lstGroups.Visible = False
            lblNewsgroupsLoaded.Visible = True
            lblLoading.Visible = True
            
            DoEvents
        
            'Put all the groups into a list box.
            Do While Not EOF(1)
            
                Line Input #1, NewsGroup
            
                'lstGroups.AddItem Left(NewsGroup, InStr(1, NewsGroup, " ") - 1)
                tempGroups(GroupCount) = LCase(Left(NewsGroup, InStr(1, NewsGroup, " ") - 1))
                GroupCount = GroupCount + 1
                
                    PeriodCount = PeriodCount + 1
                    
                    'Slows it down a lot if we update the status for each line
                    'so instead we will do it for every 1000.  On my PIII 800
                    'it loads so fast it would be pointless to update any more
                    'often.
                    If PeriodCount = 1000 Then
                    
                        lblLoading.Caption = "Loading Newsgroup Listing"
                        lblNewsgroupsLoaded.Caption = "Newsgroups Found: " & GroupCount
                        DoEvents
                        PeriodCount = 0
                    
                    End If
       
            Loop
                        
            Close #1
            
            
                ReDim Groups(GroupCount)
                
                'Put them into a form level array.
                For T = 0 To GroupCount
                
                    Groups(T) = tempGroups(T)
                
                Next T
            
                'Alphabetize them.
                lblLoading.Caption = "Sorting Newsgroup Listing"
                Call Heap_Sort
                
                PeriodCount = 0
            
            
            
                'Put them in the list box.
                For T = 0 To UBound(Groups)
            
                    lstGroups.AddItem Groups(T)
            
                    PeriodCount = PeriodCount + 1
                    
                    If PeriodCount = 1000 Then
                    
                        lblLoading.Caption = "Sorting Newsgroup Listing"
                        lblNewsgroupsLoaded.Caption = "Newsgroups Sorted: " & T
                        PeriodCount = 0
                        DoEvents
                    
                    End If
                
                Next T
            
            lstGroups.Visible = True
            lblLoading.Visible = False
            lblNewsgroupsLoaded.Visible = False
            
            
        Close #1
        
    End If


End Sub

Private Sub GetGroupRecords()

    'Open a saved group file and show me what groups
    'are in that list.
    Open App.Path & "\savedlist\" & cmbGroupListing.Text & ".lst" For Input As #1
    
    lstSelected.Clear
    
    Do While Not EOF(1)
    
        Line Input #1, NewsGroup
    
        lstSelected.AddItem NewsGroup

    Loop
    
    Close #1
    

End Sub
Private Sub cmbGroupListing_Click()

    GetGroupRecords

End Sub

Private Sub cmbServerList_Click()
    Call LoadGroupList
End Sub

Private Sub cmdDelete_Click()
    
    'Delete the group listing file
    A = MsgBox("You are about to delete the " & cmbGroupListing.Text & " group listing.  Are you sure you want to do this?", vbOKCancel, "Delete Group Listing?")

    If A = 1 Then

        Kill App.Path & "\savedlist\" & cmbGroupListing & ".lst"
        lstSelected.Clear
        Call GetGroupList
    
    End If

End Sub

Private Sub cmdGetGroups_Click()
        
'See the frmGetNewsGroups for more comments

For T = 0 To UBound(ServerNames)

    If ServerNames(T, 1) = cmbServerList.Text Then
        Exit For
    End If
    
Next T

frmGetNewsGroups.ServerAddress = ServerNames(T, 2)
frmGetNewsGroups.Username = ServerNames(T, 3)
frmGetNewsGroups.Password = ServerNames(T, 4)
frmGetNewsGroups.ServerPort = ServerNames(T, 5)
frmGetNewsGroups.ServerName = ServerNames(T, 1)

frmGetNewsGroups.Show
frmGetNewsGroups.GetNewsgroupListing
Unload frmGetNewsGroups

Call LoadGroupList

End Sub

Private Sub cmdNew_Click()

    'Get a name for the new group
    A = InputBox("What would you like to call your new group listing?", "New Group Creation")
    
    If A <> "" Then
        
        'Does the file already exist?  We don't want
        'to overwrite it.
        FileExist = Dir(App.Path & "\savedlist\" & A & ".lst", vbNormal)
        
        If FileExist = "" Then
        
            Open App.Path & "\savedlist\" & A & ".lst" For Output As #1
            Close #1
            
        End If
        
        Call GetGroupList
        
        cmbGroupListing.Text = A
        
    End If

End Sub



Private Sub SaveGroupRecords()
    
    'Save the selected groups to a file.
    Open App.Path & "\savedlist\" & cmbGroupListing.Text & ".lst" For Output As #1
    
        For T = 0 To lstSelected.ListCount - 1
        
            lstSelected.ListIndex = T
            
            Print #1, lstSelected.Text
            
        Next T
        
    Close #1

End Sub
Private Sub RefreshServerList()

    'See frmManualSend for documentation
   ' On Error GoTo RefreshServerListError

    Dim ServerNameList(65000, 6) As String
    Dim ServerInfo() As String
    
    cmbServerList.Clear
    
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

Private Sub cmdSaveList_Click()

    Call SaveGroupRecords

End Sub

Private Sub Form_Load()

    Call RefreshServerList
    Call GetGroupList

End Sub

Private Sub lstGroups_DblClick()

Exist = 0

For T = 0 To lstSelected.ListCount - 1
    
    lstSelected.ListIndex = T
    
    If lstSelected.Text = lstGroups.Text Then
        Exist = 1
    End If

Next T

    If Exist = 0 Then lstSelected.AddItem lstGroups.Text

End Sub

Private Sub lstSelected_DblClick()

lstSelected.RemoveItem (lstSelected.ListIndex)

End Sub

