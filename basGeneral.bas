Attribute VB_Name = "basGeneral"
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
'**************************************************


Public Sub Main()

Call InitializeInstall

frmMain.Show

End Sub


Private Sub InitializeInstall()

    DirListing = Dir(App.Path & "\", vbDirectory)
    
    FoundSaves = 0
    
    Do While DirListing <> ""
          
        If (GetAttr(App.Path & "\" & DirListing) And vbDirectory) = vbDirectory Then
            If LCase(DirListing) = "savedlist" Then FoundSaves = 1
        End If
    
        DirListing = Dir
    
    Loop
    
    If FoundSaves = 0 Then
        
        MkDir App.Path & "\savedlist"
        
    End If
    
End Sub


