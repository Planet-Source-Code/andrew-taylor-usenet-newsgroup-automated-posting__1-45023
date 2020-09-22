VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H00400000&
   Caption         =   "Andrew's USENET Poster"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11850
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMessageQue 
      Interval        =   60000
      Left            =   30
      Top             =   60
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuManageServers 
         Caption         =   "Manage Servers"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuManageGroups 
         Caption         =   "Manage Posting Groups"
         Shortcut        =   ^G
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMessenger 
      Caption         =   "&Messenger"
      Begin VB.Menu mnuManualSend 
         Caption         =   "Create & Send Messages"
         Shortcut        =   ^T
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSchedules 
         Caption         =   "Schedules"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub MDIForm_Load()

End Sub

Private Sub mnuExit_Click()

    End

End Sub

Private Sub mnuManageGroups_Click()

    frmManageGroups.Show

End Sub

Private Sub mnuManageServers_Click()

    frmManageServers.Show

End Sub

Private Sub mnuManualSend_Click()

    frmManualSend.Show

End Sub

