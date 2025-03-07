' file_utils.vbs - Utilities for file operations
' ------------------------------------------------------------

Option Explicit

' Save data to a file
Sub SaveDataToFile(data, filePath)
    Dim file
    
    ' Create directory if it doesn't exist
    Dim directory
    directory = Left(filePath, InStrRev(filePath, "\") - 1)
    
    If Not FSO.FolderExists(directory) Then
        FSO.CreateFolder(directory)
    End If
    
    ' Open file for writing (create if not exists, overwrite if exists)
    Set file = FSO.OpenTextFile(filePath, 2, True)
    
    ' Write data to file
    file.Write data
    
    ' Close file
    file.Close
    
    If DEBUG_MODE Then
        WScript.Echo "Data saved to " & filePath
    End If
End Sub

' Read data from a file
Function ReadDataFromFile(filePath)
    Dim file, content
    
    ' Check if file exists
    If Not FSO.FileExists(filePath) Then
        LogError "File not found: " & filePath
        ReadDataFromFile = ""
        Exit Function
    End If
    
    ' Open file for reading
    Set file = FSO.OpenTextFile(filePath, 1)
    
    ' Read entire file
    content = file.ReadAll
    
    ' Close file
    file.Close
    
    ReadDataFromFile = content
End Function

' Append data to an existing file
Sub AppendDataToFile(data, filePath)
    Dim file
    
    ' Create file if it doesn't exist
    If Not FSO.FileExists(filePath) Then
        SaveDataToFile data, filePath
        Exit Sub
    End If
    
    ' Open file for appending
    Set file = FSO.OpenTextFile(filePath, 8, True)
    
    ' Append data to file
    file.WriteLine data
    
    ' Close file
    file.Close
    
    If DEBUG_MODE Then
        WScript.Echo "Data appended to " & filePath
    End If
End Sub

' Create a log entry
Sub LogMessage(message, logType)
    Dim logFile, logFileName, timestamp
    
    ' Create timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
    
    ' Determine log file name based on type
    Select Case UCase(logType)
        Case "ERROR"
            logFileName = LOG_DIRECTORY & "\error.log"
        Case "INFO"
            logFileName = LOG_DIRECTORY & "\info.log"
        Case "DEBUG"
            If Not DEBUG_MODE Then Exit Sub
            logFileName = LOG_DIRECTORY & "\debug.log"
        Case Else
            logFileName = LOG_DIRECTORY & "\application.log"
    End Select
    
    ' Ensure log directory exists
    If Not FSO.FolderExists(LOG_DIRECTORY) Then
        FSO.CreateFolder(LOG_DIRECTORY)
    End If
    
    ' Format log message
    message = timestamp & " - " & UCase(logType) & ": " & message
    
    ' Write to log file
    AppendDataToFile message, logFileName
    
    ' Echo to console if in debug mode
    If DEBUG_MODE Then
        WScript.Echo message
    End If
End Sub

' Check if directory exists and create if not
Function EnsureDirectoryExists(directoryPath)
    If Not FSO.FolderExists(directoryPath) Then
        On Error Resume Next
        FSO.CreateFolder directoryPath
        
        If Err.Number <> 0 Then
            LogMessage "Error creating directory " & directoryPath & ": " & Err.Description, "ERROR"
            EnsureDirectoryExists = False
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    EnsureDirectoryExists = True
End Function

' Delete a file if it exists
Function DeleteFileIfExists(filePath)
    On Error Resume Next
    
    If FSO.FileExists(filePath) Then
        FSO.DeleteFile filePath, True  ' True = force delete
        
        If Err.Number <> 0 Then
            LogMessage "Error deleting file " & filePath & ": " & Err.Description, "ERROR"
            DeleteFileIfExists = False
            Exit Function
        End If
    End If
    
    On Error GoTo 0
    DeleteFileIfExists = True
End Function

' Copy a file
Function CopyFile(sourceFile, destinationFile, overwrite)
    On Error Resume Next
    
    If Not FSO.FileExists(sourceFile) Then
        LogMessage "Source file does not exist: " & sourceFile, "ERROR"
        CopyFile = False
        Exit Function
    End If
    
    FSO.CopyFile sourceFile, destinationFile, overwrite
    
    If Err.Number <> 0 Then
        LogMessage "Error copying file from " & sourceFile & " to " & destinationFile & ": " & Err.Description, "ERROR"
        CopyFile = False
        Exit Function
    End If
    
    On Error GoTo 0
    CopyFile = True
End Function

' Get list of files in a directory
Function GetFileList(directoryPath, filePattern)
    Dim files, file, fileList
    
    If Not FSO.FolderExists(directoryPath) Then
        LogMessage "Directory does not exist: " & directoryPath, "ERROR"
        GetFileList = Array()
        Exit Function
    End If
    
    Set files = FSO.GetFolder(directoryPath).Files
    fileList = ""
    
    For Each file In files
        If filePattern = "" Or InStr(file.Name, filePattern) > 0 Then
            If fileList <> "" Then fileList = fileList & ","
            fileList = fileList & file.Path
        End If
    Next
    
    If fileList = "" Then
        GetFileList = Array()
    Else
        GetFileList = Split(fileList, ",")
    End If
End Function
