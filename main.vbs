Function GetInput()
    Dim userInput
    Dim scriptPath

    userInput = InputBox("1. FPS UNLOCKER" & vbCrLf & "2. MODTOOLS" & vbCrLf & "3. LOAD SAVED SETTINGS" & vbCrLf & "INFO. INFORMATION", "Roblox Modtools")

    ' Convert user input to uppercase for case-insensitive comparison
    userInput = UCase(userInput)

    ' Check if user input is valid and set the script path accordingly
    Select Case userInput
        Case "1"
            scriptPath = "fps_unlocker.vbs"
        Case "2"
            scriptPath = "modtools.bat"
        Case "3"
            scriptPath = "load_settings.vbs"
        Case "INFO"
            scriptPath = "info.txt"
        Case Else
            ' Exit the script
            WScript.Quit
    End Select

    ' Run the selected script
    RunScript scriptPath
End Function

Sub RunScript(scriptPath)
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run scriptPath, 1, True
End Sub

' Call the function to get user input
GetInput
