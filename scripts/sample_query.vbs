' sample_query.vbs - Example script for querying mainframe data
' ------------------------------------------------------------

Option Explicit

' Include necessary libraries and configuration
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../config/settings.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/api_utils.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/data_utils.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/file_utils.vbs", 1).ReadAll

' Main script execution
Sub Main()
    Dim session
    Dim accountNumber
    
    ' Get account number from command line or use default
    If WScript.Arguments.Count > 0 Then
        accountNumber = WScript.Arguments(0)
    Else
        accountNumber = "12345678"  ' Default account number for testing
    End If
    
    ' Display startup message
    WScript.Echo "Starting account query for account: " & accountNumber
    
    ' Initialize BlueZone session
    If InitializeBlueZone(session) Then
        ' Login to mainframe
        If LoginToMainframe(session) Then
            ' Navigate to customer inquiry screen
            If NavigateToScreen(session, "CUSTOMER_INQUIRY") Then
                ' Enter account number and submit query
                session.SetCursorPos 10, 25
                session.SendString accountNumber
                session.SendKeys "<Enter>"
                
                ' Wait for response
                session.WaitForString "ACCOUNT DETAILS", 5
                
                ' Extract account data
                Dim screenData, accountData
                screenData = GetScreenData(session)
                
                ' Process the screen data into structured format
                accountData = ProcessData(screenData)
                
                ' Save the data to file
                SaveDataToFile accountData, OUTPUT_DIRECTORY & "\account_" & accountNumber & ".json"
                
                ' Display success message
                WScript.Echo "Account data retrieved successfully and saved to file."
                
                ' Optional: Display formatted data
                WScript.Echo "Account Information:"
                WScript.Echo FormatData(ParseJson(accountData), "TABLE")
            End If
            
            ' Logout from mainframe
            LogoutFromMainframe session
        End If
        
        ' Disconnect session
        DisconnectBlueZone session
    End If
    
    WScript.Echo "Script execution completed."
End Sub

' Extract specific account information from screen data
Function ExtractAccountInfo(screenText)
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Extract key account information - customize these field names based on your mainframe screens
    dict.Add "AccountNumber", ExtractField(screenText, "Account Number")
    dict.Add "CustomerName", ExtractField(screenText, "Customer Name")
    dict.Add "Balance", ExtractField(screenText, "Balance")
    dict.Add "AccountType", ExtractField(screenText, "Account Type")
    dict.Add "OpenDate", ExtractField(screenText, "Open Date")
    dict.Add "LastTransaction", ExtractField(screenText, "Last Transaction")
    
    Set ExtractAccountInfo = dict
End Function

' Start the main procedure
Call Main()
