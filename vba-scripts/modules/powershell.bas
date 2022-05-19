' Executes a powershell command in a powershell window
Public Function PS_Execute(ByVal sPSCmd As String)
    'Setup the powershell command properly
    sPSCmd = "powershell -noprofile -nologo -command " & sPSCmd
    'Execute and capture the returned value
    CreateObject("WScript.Shell").Exec (sPSCmd)
End Sub

' Silently executes a powershell command
Public Function PS_Run(ByVal sPSCmd As String)
    'Setup the powershell command properly
    sPSCmd = "powershell -noprofile -nologo  -command " & sPSCmd
    'Execute and capture the returned value
    CreateObject("WScript.Shell").Run sPSCmd, 0, True
End Sub

' Silently executes a powershell command and returns the output
Public Function PS_GetOutput(ByVal sPSCmd As String) As String
    'Setup the powershell command properly
    sPSCmd = "powershell -noprofile -nologo -command " & sPSCmd
    'Execute and capture the returned value
    PS_GetOutput = CreateObject("WScript.Shell").Exec(sPSCmd).StdOut.ReadAll
End Function