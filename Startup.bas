Attribute VB_Name = "Startup"
Sub Main()


    Load DebugData ' Form is now in memory but invisible
    DebugData.Visible = False

    ' Now show your real startup form
    Quest.Show
    
End Sub
