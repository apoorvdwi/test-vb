' auth_token_example.vbs - Example of using auth tokens with API calls
' ------------------------------------------------------------

Option Explicit

' Include necessary libraries and configuration
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../config/settings.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/api_utils.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/data_utils.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/file_utils.vbs", 1).ReadAll

' Main script execution
Sub Main()
    Dim endpoint, response, data
    
    ' Display startup message
    WScript.Echo "Starting Auth Token API Example"
    
    ' Log current authentication settings
    WScript.Echo "Authentication type: " & AUTH_TYPE
    WScript.Echo "Using token for authentication" & vbCrLf
    
    ' Example: Secured endpoint that requires authentication
    endpoint = API_BASE_URL & "/secured/users"
    WScript.Echo "Making authenticated GET request to " & endpoint
    
    response = HttpGet(endpoint)
    
    If response <> "" Then
        WScript.Echo "API call successful!"
        WScript.Echo "Response: " & response
        
        ' Save response to file
        SaveDataToFile response, OUTPUT_DIRECTORY & "\secured_api_response.json"
    Else
        WScript.Echo "API call failed!"
    End If
    
    ' Example: POST request with authentication
    endpoint = API_BASE_URL & "/secured/data"
    
    ' Create sample data
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "name", "John Doe"
    data.Add "email", "john.doe@example.com"
    data.Add "department", "Finance"
    
    ' Convert to JSON
    Dim jsonData
    jsonData = ConvertToJson(data)
    
    WScript.Echo vbCrLf & "Making authenticated POST request to " & endpoint
    WScript.Echo "Data: " & jsonData
    
    ' Make the API call with authentication and retry logic
    response = CallAPIWithRetry(endpoint, "POST", jsonData)
    
    If response <> "" Then
        WScript.Echo "API call successful!"
        WScript.Echo "Response: " & response
        
        ' Parse the response
        Dim responseData
        Set responseData = ParseJson(response)
        
        ' Display the parsed data
        If responseData.Count > 0 Then
            WScript.Echo vbCrLf & "Parsed Response Data:"
            Dim key
            For Each key In responseData.Keys
                WScript.Echo key & ": " & responseData(key)
            Next
        End If
    Else
        WScript.Echo "API call failed!"
    End If
    
    WScript.Echo vbCrLf & "Script execution completed."
End Sub

' Example of manually setting auth token at runtime
Sub SetAuthTokenManually(newToken, newType)
    ' Note: In a real implementation, these would modify global constants
    ' For demonstration purposes only
    WScript.Echo "Changing auth token from " & AUTH_TOKEN & " to " & newToken
    WScript.Echo "Changing auth type from " & AUTH_TYPE & " to " & newType
    
    ' In a real scenario, you would need to create a different approach
    ' since constants cannot be modified at runtime
    ' This is just for demonstration purposes
End Sub

' Start the main procedure
Call Main()
