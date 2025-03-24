' modify_and_update_data.vbs - Script for fetching, modifying, and updating data via API
' ------------------------------------------------------------

Option Explicit

' Include necessary libraries and configuration
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../config/settings.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/api_utils.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/data_utils.vbs", 1).ReadAll
ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("../lib/file_utils.vbs", 1).ReadAll

' Main script execution
Sub Main()
    Dim endpoint, response, jsonData, modifiedData

    ' Display startup message
    WScript.Echo "Starting data modification and update process"
    
    ' Step 1: Make GET request to fetch the data
    endpoint = API_BASE_URL & "/data"
    WScript.Echo "Fetching data from " & endpoint
    
    response = HttpGet(endpoint)
    
    If response = "" Then
        WScript.Echo "API call failed! Could not retrieve data."
        Exit Sub
    End If
    
    WScript.Echo "Data retrieved successfully!"
    
    ' Step 2: Parse the JSON response
    Set jsonData = ParseJson(response)
    
    ' Step 3: Modify the data
    Set modifiedData = ModifyArrayData(jsonData)
    
    ' Step 4: Convert modified data back to JSON string
    Dim jsonString
    jsonString = ConvertToJson(modifiedData)
    
    ' Step 5: Send PUT request to update the data
    WScript.Echo "Updating data at " & endpoint
    response = HttpPut(endpoint, jsonString)
    
    If response <> "" Then
        WScript.Echo "Data updated successfully!"
        WScript.Echo "Response: " & response
    Else
        WScript.Echo "Failed to update data!"
    End If
    
    WScript.Echo "Process completed."
End Sub

' Function to modify array data
' This function handles the array of arrays structure
Function ModifyArrayData(jsonData)
    Dim modifiedData, i, j, item
    
    ' Create a copy of the original data structure
    Set modifiedData = CloneObject(jsonData)
    
    ' Check if we have an array of arrays
    If IsArray(modifiedData) Then
        WScript.Echo "Processing array data..."
        
        ' Loop through the outer array
        For i = 0 To UBound(modifiedData)
            ' Check if this element is also an array
            If IsArray(modifiedData(i)) Then
                ' Loop through the inner array
                For j = 0 To UBound(modifiedData(i))
                    ' If the inner element is a dictionary (object)
                    If TypeName(modifiedData(i)(j)) = "Dictionary" Then
                        ' Modify the specific field(s) you need to change
                        ' Example: modifying a field called "status" to "updated"
                        If modifiedData(i)(j).Exists("status") Then
                            modifiedData(i)(j)("status") = "updated"
                        End If
                        
                        ' You can add more field modifications here as needed
                    End If
                Next
            End If
        Next
    ElseIf TypeName(modifiedData) = "Dictionary" Then
        ' If the root is a dictionary object
        WScript.Echo "Processing dictionary data..."
        
        ' Handle nested arrays in dictionary properties
        For Each item In modifiedData.Keys
            If IsArray(modifiedData(item)) Then
                ' Process arrays within dictionary
                modifiedData(item) = ProcessNestedArray(modifiedData(item))
            End If
        Next
    End If
    
    Set ModifyArrayData = modifiedData
End Function

' Helper function to process nested arrays
Function ProcessNestedArray(arrData)
    Dim result, i, j, item
    
    ' Create a copy of the array
    result = arrData
    
    ' Process each item in the array
    For i = 0 To UBound(result)
        If IsArray(result(i)) Then
            ' Process nested array
            For j = 0 To UBound(result(i))
                If TypeName(result(i)(j)) = "Dictionary" Then
                    ' Modify specific fields in dictionary
                    ' Example: modifying a field called "status" to "updated"
                    If result(i)(j).Exists("status") Then
                        result(i)(j)("status") = "updated"
                    End If
                End If
            Next
        ElseIf TypeName(result(i)) = "Dictionary" Then
            ' Modify dictionary directly
            If result(i).Exists("status") Then
                result(i)("status") = "updated"
            End If
        End If
    Next
    
    ProcessNestedArray = result
End Function

' Function to create a deep copy of an object (array or dictionary)
Function CloneObject(obj)
    Dim result, key, i
    
    If IsArray(obj) Then
        ' Clone array
        ReDim result(UBound(obj))
        For i = 0 To UBound(obj)
            If IsObject(obj(i)) Then
                Set result(i) = CloneObject(obj(i))
            ElseIf IsArray(obj(i)) Then
                result(i) = CloneObject(obj(i))
            Else
                result(i) = obj(i)
            End If
        Next
        CloneObject = result
    ElseIf TypeName(obj) = "Dictionary" Then
        ' Clone dictionary
        Set result = CreateObject("Scripting.Dictionary")
        For Each key In obj.Keys
            If IsObject(obj(key)) Then
                Set result(key) = CloneObject(obj(key))
            ElseIf IsArray(obj(key)) Then
                result(key) = CloneObject(obj(key))
            Else
                result(key) = obj(key)
            End If
        Next
        Set CloneObject = result
    ElseIf IsObject(obj) Then
        ' For other objects, just set reference (can't deep clone arbitrary objects)
        Set CloneObject = obj
    Else
        ' For simple values
        CloneObject = obj
    End If
End Function

' HTTP PUT Function
Function HttpPut(url, data)
    Dim http
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    On Error Resume Next
    http.Open "PUT", url, False
    http.setRequestHeader "Content-Type", "application/json"
    
    ' Add authorization header if needed
    If API_AUTH_REQUIRED Then
        http.setRequestHeader "Authorization", "Bearer " & API_AUTH_TOKEN
    End If
    
    http.send data
    
    If Err.Number <> 0 Then
        WScript.Echo "Error in HttpPut: " & Err.Description
        HttpPut = ""
    Else
        If http.status >= 200 And http.status < 300 Then
            HttpPut = http.responseText
        Else
            WScript.Echo "HTTP Error: " & http.status & " - " & http.statusText
            HttpPut = ""
        End If
    End If
    
    On Error GoTo 0
    Set http = Nothing
End Function

' Start the main procedure
Call Main()
