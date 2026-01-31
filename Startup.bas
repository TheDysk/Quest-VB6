Attribute VB_Name = "Startup"
Sub Main()


    Load DebugData ' Form is now in memory but invisible
    DebugData.Visible = False ' Setting visibility to false to ensure form is not visible (incase of incorrect form visibility setting during testing and copilation)
    
    ' Now show your real startup form
    Quest.Show
    
End Sub
