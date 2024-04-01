Function CheckGameVersion(gameVersion)
    Dim versionPath
    Dim robloxPlayerPath
    Dim clientSettingsPath
    Dim clientAppSettingsPath
    Dim objFSO
    Dim result
    
    ' Construct the path to the game version folder
    versionPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Roblox\versions\version-" & gameVersion
    
    ' Check if the folder exists
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(versionPath) Then
        ' Check if RobloxPlayerBeta.dll exists in the folder
        robloxPlayerPath = versionPath & "\RobloxPlayerBeta.dll"
        If objFSO.FileExists(robloxPlayerPath) Then
            ' Check if ClientSettings folder exists
            clientSettingsPath = versionPath & "\ClientSettings"
            If Not objFSO.FolderExists(clientSettingsPath) Then
                ' Create ClientSettings folder if it doesn't exist
                objFSO.CreateFolder clientSettingsPath
            End If
            
            ' Check if ClientAppSettings.json file exists inside ClientSettings folder
            clientAppSettingsPath = clientSettingsPath & "\ClientAppSettings.json"
            If Not objFSO.FileExists(clientAppSettingsPath) Then
                ' Create ClientAppSettings.json file if it doesn't exist
                Dim objFile
                Set objFile = objFSO.CreateTextFile(clientAppSettingsPath)
                objFile.WriteLine("{""DFIntTaskSchedulerTargetFps"": 60}")
                objFile.Close
            End If
            
            ' If everything is found, set result to True
            result = "True"
        Else
            result = "False"
        End If
    Else
        result = "False"
    End If
    
    CheckGameVersion = result
End Function

' Function to get game version from the user
Function GetGameVersion()
    Dim gameVersion
    gameVersion = InputBox("Enter the game version (e.g., f573c8cc796e4c97):", "FPS Unlocker")
    GetGameVersion = gameVersion
Set fso = CreateObject("Scripting.FileSystemObject")
username = WScript.CreateObject("WScript.Network").UserName
folderPath = "C:\Users\" & username & "\AppData\Local\Roblox\Versions\version-" & gameVersion
dllFilePath = folderPath & "\RobloxPlayerBeta.dll"

If gameVersion = "" Or Not fso.FolderExists(folderPath) Then
    WScript.Quit
End If

If Not fso.FileExists(dllFilePath) Then
    WScript.Quit
End If
End Function

' Function to get desired FPS from the user
Function GetDesiredFPS()
    Dim fps
    fps = InputBox("Set FPS to (leave empty for default):", "FPS Unlocker")
    If fps = "" Then
        fps = 0
    End If
    GetDesiredFPS = fps
End Function

' Main function
Sub Main()
    Dim gameVersion
    Dim isFileExists
    
    ' Get game version from the user
    gameVersion = GetGameVersion()
    
    ' Check if RobloxPlayerBeta.dll exists and create ClientSettings folder if needed
    isFileExists = CheckGameVersion(gameVersion)
        
    ' Get desired FPS from the user
    Dim desiredFPS
    desiredFPS = GetDesiredFPS()
    
    ' Update ClientAppSettings.json with the desired FPS
    Dim clientSettingsPath
    Dim clientAppSettingsPath
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    clientSettingsPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Roblox\versions\version-" & gameVersion & "\ClientSettings"
    clientAppSettingsPath = clientSettingsPath & "\ClientAppSettings.json"
    If objFSO.FileExists(clientAppSettingsPath) Then
        Dim objFile
        Set objFile = objFSO.OpenTextFile(clientAppSettingsPath, 2)
        objFile.WriteLine("{""DFIntTaskSchedulerTargetFps"": " & desiredFPS & "}")
        objFile.Close
    End If
    
    ' Save the desired FPS value and game version in settings file
    Dim scriptPath
    Dim settingsFile
    Dim settingsContent
    scriptPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
    settingsFile = scriptPath & "settings"
    If objFSO.FileExists(settingsFile) Then
        ' Read existing settings file
        Set objFile = objFSO.OpenTextFile(settingsFile, 1)
        settingsContent = objFile.ReadAll
        objFile.Close
        
        ' Replace existing FPS and version lines or append them if not found
        Dim arrLines
        arrLines = Split(settingsContent, vbCrLf)
        Dim updatedContent
        Dim fpsFound
        Dim versionFound
        fpsFound = False
        versionFound = False
        For Each line In arrLines
            If InStr(line, "FPS = ") = 1 Then
                ' Replace existing FPS line
                updatedContent = updatedContent & "FPS = " & desiredFPS & vbCrLf
                fpsFound = True
            ElseIf InStr(line, "VERSION = ") = 1 Then
                ' Replace existing version line
                updatedContent = updatedContent & "VERSION = " & gameVersion & vbCrLf
                versionFound = True
            Else
                ' Keep other lines unchanged
                updatedContent = updatedContent & line & vbCrLf
            End If
        Next
        ' If FPS line not found, append it
        If Not fpsFound Then
            updatedContent = updatedContent & "FPS = " & desiredFPS & vbCrLf
        End If
        ' If version line not found, append it
        If Not versionFound Then
            updatedContent = updatedContent & "VERSION = " & gameVersion & vbCrLf
        End If
        
        ' Write updated content back to settings file
        Set objFile = objFSO.OpenTextFile(settingsFile, 2)
        objFile.Write updatedContent
        objFile.Close
    Else
        ' Create new settings file
        Set objFile = objFSO.CreateTextFile(settingsFile, True)
        objFile.WriteLine("FPS = " & desiredFPS)
        objFile.WriteLine("VERSION = " & gameVersion)
        objFile.Close
    End If
End Sub

' Call the main function
Main
