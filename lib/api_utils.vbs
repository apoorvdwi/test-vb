' api_utils.vbs - Utilities for making HTTP API calls
' ------------------------------------------------------------

Option Explicit

' Make an HTTP API call with the specified method and data
Function MakeAPICall(url, method, data)
    Dim http, response
    
    ' Create HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Initialize the HTTP request
    http.Open method, url, False
    
    ' Set request headers
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' Set authorization header with token if available
    If Len(AUTH_TOKEN) > 0 Then
        http.SetRequestHeader "Authorization", AUTH_TYPE & " " & AUTH_TOKEN
    ElseIf Len(API_KEY) > 0 Then
        ' Fall back to API key if token is not available
        http.SetRequestHeader "Authorization", "Bearer " & API_KEY
    End If
    
    ' Send the request with data if provided
    If method = "GET" Or IsEmpty(data) Then
        http.Send
    Else
        http.Send data
    End If
    
    ' Check for successful response
    If http.Status >= 200 And http.Status < 300 Then
        ' Return the response text
        MakeAPICall = http.ResponseText
    Else
        ' Log error
        LogError "API call failed: " & http.Status & " - " & http.StatusText & " - " & http.ResponseText
        MakeAPICall = ""
    End If
    
    ' Clean up
    Set http = Nothing
End Function

' Make an HTTP GET request to the specified URL
Function HttpGet(url)
    HttpGet = MakeAPICall(url, "GET", Empty)
End Function

' Make an HTTP POST request with JSON data
Function HttpPost(url, jsonData)
    HttpPost = MakeAPICall(url, "POST", jsonData)
End Function

' Make an HTTP PUT request with JSON data
Function HttpPut(url, jsonData)
    HttpPut = MakeAPICall(url, "PUT", jsonData)
End Function

' Make an HTTP DELETE request
Function HttpDelete(url)
    HttpDelete = MakeAPICall(url, "DELETE", Empty)
End Function

' Convert a dictionary or array to JSON string
Function ConvertToJson(data)
    Dim jsonString, key
    
    If TypeName(data) = "Dictionary" Then
        ' Convert dictionary to JSON
        jsonString = "{"
        For Each key In data.Keys
            If Len(jsonString) > 1 Then jsonString = jsonString & ","
            jsonString = jsonString & """" & key & """:""" & data(key) & """"
        Next
        jsonString = jsonString & "}"
        ConvertToJson = jsonString
    ElseIf IsArray(data) Then
        ' Convert array to JSON
        jsonString = "["
        Dim i
        For i = LBound(data) To UBound(data)
            If i > LBound(data) Then jsonString = jsonString & ","
            jsonString = jsonString & """" & data(i) & """"
        Next
        jsonString = jsonString & "]"
        ConvertToJson = jsonString
    Else
        ' Return the data as is (assuming it's already a JSON string)
        ConvertToJson = data
    End If
End Function

' Parse a JSON string into a Dictionary
Function ParseJson(jsonString)
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Implement simple JSON parsing logic
    ' Note: For production use, consider using a more robust JSON parser
    ' This is a simplified implementation for basic JSON responses
    
    If Left(Trim(jsonString), 1) = "{" And Right(Trim(jsonString), 1) = "}" Then
        ' Strip curly braces
        jsonString = Mid(jsonString, 2, Len(jsonString) - 2)
        
        ' Split by commas outside of quotes
        Dim pairs, pair
        pairs = Split(jsonString, ",")
        
        For Each pair In pairs
            ' Split key and value
            Dim keyValue, key, value
            keyValue = Split(pair, ":", 2)
            
            If UBound(keyValue) >= 1 Then
                ' Clean up key and value
                key = Trim(keyValue(0))
                value = Trim(keyValue(1))
                
                ' Remove surrounding quotes if present
                If Left(key, 1) = """" And Right(key, 1) = """" Then
                    key = Mid(key, 2, Len(key) - 2)
                End If
                
                If Left(value, 1) = """" And Right(value, 1) = """" Then
                    value = Mid(value, 2, Len(value) - 2)
                End If
                
                ' Add to dictionary
                dict.Add key, value
            End If
        Next
    End If
    
    Set ParseJson = dict
End Function

' Handle API response with retry logic
Function CallAPIWithRetry(url, method, data)
    Dim response
    Dim attempt
    
    For attempt = 1 To MAX_RETRY_ATTEMPTS
        response = MakeAPICall(url, method, data)
        
        ' Check if the call was successful
        If response <> "" Then
            CallAPIWithRetry = response
            Exit Function
        End If
        
        ' Log retry attempt
        If DEBUG_MODE Then
            WScript.Echo "API call failed. Retry attempt " & attempt & " of " & MAX_RETRY_ATTEMPTS
        End If
        
        ' Wait before retry
        WScript.Sleep RETRY_DELAY
    Next
    
    ' All retry attempts failed
    LogError "API call failed after " & MAX_RETRY_ATTEMPTS & " retry attempts"
    CallAPIWithRetry = ""
End Function

' Log an error message
Sub LogError(message)
    Dim logFile
    Dim timestamp
    
    ' Create timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)
    
    ' Log to console if in debug mode
    If DEBUG_MODE Then
        WScript.Echo timestamp & " - ERROR: " & message
    End If
    
    ' Log to file
    Set logFile = FSO.OpenTextFile(LOG_DIRECTORY & "\api_error.log", 8, True)
    logFile.WriteLine timestamp & " - " & message
    logFile.Close
End Sub
