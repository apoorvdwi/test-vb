' sample_api_call.vbs - Example script for making external API calls
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
    WScript.Echo "Starting API call example"
    
    ' Example 1: Simple GET request
    endpoint = API_BASE_URL & "/users"
    WScript.Echo "Making GET request to " & endpoint
    
    response = HttpGet(endpoint)
    
    If response <> "" Then
        WScript.Echo "API call successful!"
        WScript.Echo "Response: " & response
        
        ' Save response to file
        SaveDataToFile response, OUTPUT_DIRECTORY & "\api_users.json"
    Else
        WScript.Echo "API call failed!"
    End If
    
    ' Example 2: POST request with data
    endpoint = API_BASE_URL & "/data"
    
    ' Create sample data
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "name", "John Doe"
    data.Add "email", "john.doe@example.com"
    data.Add "status", "active"
    
    ' Convert to JSON
    Dim jsonData
    jsonData = ConvertToJson(data)
    
    WScript.Echo vbCrLf & "Making POST request to " & endpoint
    WScript.Echo "Data: " & jsonData
    
    ' Make the API call with retry logic
    response = CallAPIWithRetry(endpoint, "POST", jsonData)
    
    If response <> "" Then
        WScript.Echo "API call successful!"
        WScript.Echo "Response: " & response
        
        ' Save response to file
        SaveDataToFile response, OUTPUT_DIRECTORY & "\api_post_response.json"
        
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

' Create a sample customer record
Function CreateCustomerRecord(customerId, name, email)
    Dim customer
    Set customer = CreateObject("Scripting.Dictionary")
    
    customer.Add "id", customerId
    customer.Add "name", name
    customer.Add "email", email
    customer.Add "createdAt", Now
    
    Set CreateCustomerRecord = customer
End Function

' Start the main procedure
Call Main()
