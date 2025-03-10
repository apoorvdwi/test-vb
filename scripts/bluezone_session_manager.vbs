' bluezone_session_manager.vbs
' Utility script for managing BluZone sessions
' ------------------------------------------------------------

Option Explicit

' Include configuration file
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("config/settings.vbs", 1).ReadAll

' Main function to demonstrate usage
Sub Main()
    Dim choice, sessionName
    
    ' Display menu
    WScript.Echo "BluZone Session Manager"
    WScript.Echo "1. Connect to BluZone session"
    WScript.Echo "2. Disconnect from BluZone session"
    WScript.Echo "3. Check BluZone session status"
    WScript.Echo "4. Exit"
    
    choice = InputBox("Enter your choice (1-4):", "BluZone Session Manager")
    
    Select Case choice
        Case "1"
            sessionName = InputBox("Enter session name (or leave blank for default):", "Connect to BluZone")
            If sessionName = "" Then sessionName = BLUEZONE_SESSION_CONFIG
            ConnectToBluZoneSession sessionName
            
        Case "2"
            DisconnectFromBluZoneSession
            
        Case "3"
            If CheckBluZoneSessionStatus() Then
                WScript.Echo "BluZone session is currently active."
            Else
                WScript.Echo "No active BluZone session found."
            End If
            
        Case "4"
            WScript.Echo "Exiting..."
            
        Case Else
            WScript.Echo "Invalid choice. Please try again."
    End Select
End Sub

' Connect to a BluZone session
Function ConnectToBluZoneSession(sessionConfig)
    On Error Resume Next
    
    Dim session
    WScript.Echo "Attempting to connect to BluZone session: " & sessionConfig
    
    ' Create BlueZone session object
    Set session = CreateObject("BlueZone.AutomationManager")
    
    If Err.Number <> 0 Then
        WScript.Echo "Error creating BlueZone automation manager: " & Err.Description
        ConnectToBluZoneSession = False
        Exit Function
    End If
    
    ' Connect to BlueZone
    session.Connect
    
    If Err.Number <> 0 Then
        WScript.Echo "Error connecting to BlueZone: " & Err.Description
        ConnectToBluZoneSession = False
        Exit Function
    End If
    
    ' Initialize session with the configuration
    session.OpenSession sessionConfig
    
    If Err.Number <> 0 Then
        WScript.Echo "Error opening BlueZone session: " & Err.Description
        session.Disconnect
        Set session = Nothing
        ConnectToBluZoneSession = False
        Exit Function
    End If
    
    ' Store session in a global object for later use
    Set GLOBAL_BLUEZONE_SESSION = session
    
    WScript.Echo "Successfully connected to BluZone session."
    On Error GoTo 0
    ConnectToBluZoneSession = True
End Function

' Disconnect from a BluZone session
Function DisconnectFromBluZoneSession()
    On Error Resume Next
    
    ' Check if we have an active session
    If Not CheckBluZoneSessionStatus() Then
        WScript.Echo "No active BluZone session to disconnect."
        DisconnectFromBluZoneSession = False
        Exit Function
    End If
    
    Dim session
    Set session = GLOBAL_BLUEZONE_SESSION
    
    ' Close the session
    session.CloseSession
    
    If Err.Number <> 0 Then
        WScript.Echo "Warning: Error closing BluZone session: " & Err.Description
        Err.Clear
    End If
    
    ' Disconnect from BlueZone
    session.Disconnect
    
    If Err.Number <> 0 Then
        WScript.Echo "Warning: Error disconnecting from BluZone: " & Err.Description
        DisconnectFromBluZoneSession = False
        Exit Function
    End If
    
    ' Clean up
    Set GLOBAL_BLUEZONE_SESSION = Nothing
    Set session = Nothing
    
    WScript.Echo "Successfully disconnected from BluZone session."
    On Error GoTo 0
    DisconnectFromBluZoneSession = True
End Function

' Check if a BluZone session is active
Function CheckBluZoneSessionStatus()
    On Error Resume Next
    
    ' Check if global session object exists and is valid
    If IsObject(GLOBAL_BLUEZONE_SESSION) Then
        ' Try to access a property to confirm it's still valid
        Dim dummy
        dummy = GLOBAL_BLUEZONE_SESSION.IsConnected
        
        If Err.Number = 0 Then
            CheckBluZoneSessionStatus = True
            Exit Function
        End If
        
        Err.Clear
    End If
    
    On Error GoTo 0
    CheckBluZoneSessionStatus = False
End Function

' Define a global variable to hold the session object
Dim GLOBAL_BLUEZONE_SESSION
Set GLOBAL_BLUEZONE_SESSION = Nothing

' If this script is run directly (not included), execute the Main procedure
If WScript.ScriptName = "bluezone_session_manager.vbs" Then
    Call Main()
End If
