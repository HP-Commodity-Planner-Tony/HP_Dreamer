Attribute VB_Name = "Run_Python"
Sub Run_Python()
    
    'Run python script https://stackoverflow.com/questions/15951837/wait-for-shell-command-to-complete
    Dim wsh As Object, path As String
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    path = ThisWorkbook.path & "\KB_Demand_WF 5.0.exe"
    wsh.Run Chr(34) & path & Chr(34), windowStyle, waitOnReturn
    
End Sub

