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
    Dim jsonString
    
    Select Case TypeName(data)
        Case "Dictionary"
            ' Convert dictionary to JSON
            jsonString = "{"
            Dim dictKey, dictValue, needComma
            needComma = False
            
            For Each dictKey In data.Keys
                If needComma Then jsonString = jsonString & ","
                
                ' Handle key
                jsonString = jsonString & """" & Replace(dictKey, """", "\""") & """:"
                
                ' Handle value based on its type
                Select Case TypeName(data(dictKey))
                    Case "Dictionary", "Variant()"
                        ' Recursively convert nested objects/arrays
                        jsonString = jsonString & ConvertToJson(data(dictKey))
                    Case "String"
                        jsonString = jsonString & """" & Replace(data(dictKey), """", "\""") & """"
                    Case "Boolean"
                        If data(dictKey) Then
                            jsonString = jsonString & "true"
                        Else
                            jsonString = jsonString & "false"
                        End If
                    Case "Null"
                        jsonString = jsonString & "null"
                    Case "Integer", "Long", "Double", "Single"
                        jsonString = jsonString & data(dictKey)
                    Case Else
                        ' Default to string representation for other types
                        jsonString = jsonString & """" & Replace(CStr(data(dictKey)), """", "\""") & """"
                End Select
                
                needComma = True
            Next
            
            jsonString = jsonString & "}"
            
        Case "Variant()", "Byte()"
            ' Convert array to JSON array
            Dim i, arrValue
            jsonString = "["
            needComma = False
            
            For i = LBound(data) To UBound(data)
                If needComma Then jsonString = jsonString & ","
                
                ' Handle value based on its type
                Select Case TypeName(data(i))
                    Case "Dictionary", "Variant()"
                        ' Recursively convert nested objects/arrays
                        jsonString = jsonString & ConvertToJson(data(i))
                    Case "String"
                        jsonString = jsonString & """" & Replace(data(i), """", "\""") & """"
                    Case "Boolean"
                        If data(i) Then
                            jsonString = jsonString & "true"
                        Else
                            jsonString = jsonString & "false"
                        End If
                    Case "Null"
                        jsonString = jsonString & "null"
                    Case "Integer", "Long", "Double", "Single"
                        jsonString = jsonString & data(i)
                    Case Else
                        ' Default to string representation for other types
                        jsonString = jsonString & """" & Replace(CStr(data(i)), """", "\""") & """"
                End Select
                
                needComma = True
            Next
            
            jsonString = jsonString & "]"
            
        Case "Boolean"
            If data Then
                jsonString = "true"
            Else
                jsonString = "false"
            End If
            
        Case "Null"
            jsonString = "null"
            
        Case "Integer", "Long", "Double", "Single"
            jsonString = data
            
        Case Else
            ' Default to string for all other types
            If IsNull(data) Then
                jsonString = "null"
            ElseIf IsEmpty(data) Then
                jsonString = "null"
            Else
                jsonString = """" & Replace(CStr(data), """", "\""") & """"
            End If
    End Select
    
    ConvertToJson = jsonString
End Function

' Parse a JSON string into appropriate VBScript objects (Dictionary for objects, Array for arrays)
Function ParseJson(jsonStr)
    Dim jsonText, charIndex, char
    
    ' Initialize
    jsonText = Trim(jsonStr)
    charIndex = 1
    
    ' Call the recursive parsing function
    Set ParseJson = ParseValue(jsonText, charIndex)
End Function

' Helper function to parse a JSON value at the current position
Function ParseValue(jsonText, ByRef charIndex)
    Dim char
    
    ' Skip whitespace
    SkipWhitespace jsonText, charIndex
    
    ' Get current character
    char = Mid(jsonText, charIndex, 1)
    
    Select Case char
        Case "{"  ' Object
            Set ParseValue = ParseObject(jsonText, charIndex)
            
        Case "["  ' Array
            ParseValue = ParseArray(jsonText, charIndex)
            
        Case """"  ' String
            ParseValue = ParseString(jsonText, charIndex)
            
        Case "t", "f"  ' Boolean (true/false)
            ParseValue = ParseBoolean(jsonText, charIndex)
            
        Case "n"  ' Null
            ParseValue = ParseNull(jsonText, charIndex)
            
        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"  ' Number
            ParseValue = ParseNumber(jsonText, charIndex)
            
        Case Else
            ' Unexpected character
            Err.Raise 9001, "ParseJson", "Unexpected character in JSON: " & char & " at position " & charIndex
    End Select
End Function

' Helper function to parse a JSON object (dictionary)
Function ParseObject(jsonText, ByRef charIndex)
    Dim obj, char, key, expectComma
    
    ' Create dictionary object
    Set obj = CreateObject("Scripting.Dictionary")
    
    ' Move past the opening brace
    charIndex = charIndex + 1
    
    ' Skip whitespace
    SkipWhitespace jsonText, charIndex
    
    ' Check for empty object
    If Mid(jsonText, charIndex, 1) = "}" Then
        charIndex = charIndex + 1
        Set ParseObject = obj
        Exit Function
    End If
    
    ' Parse key-value pairs
    expectComma = False
    Do
        ' If expecting a comma, check for it
        If expectComma Then
            SkipWhitespace jsonText, charIndex
            char = Mid(jsonText, charIndex, 1)
            If char <> "," Then
                Err.Raise 9001, "ParseJson", "Expected comma in object at position " & charIndex
            End If
            charIndex = charIndex + 1
        End If
        
        ' Skip whitespace
        SkipWhitespace jsonText, charIndex
        
        ' Check for closing brace (end of object)
        If Mid(jsonText, charIndex, 1) = "}" Then
            charIndex = charIndex + 1
            Exit Do
        End If
        
        ' Parse key (must be a string)
        key = ParseString(jsonText, charIndex)
        
        ' Skip whitespace
        SkipWhitespace jsonText, charIndex
        
        ' Check for colon
        If Mid(jsonText, charIndex, 1) <> ":" Then
            Err.Raise 9001, "ParseJson", "Expected colon after key in object at position " & charIndex
        End If
        charIndex = charIndex + 1
        
        ' Parse value
        If key <> "" Then
            obj.Add key, ParseValue(jsonText, charIndex)
        End If
        
        ' Expect comma for next key-value pair
        expectComma = True
    Loop
    
    ' Return object
    Set ParseObject = obj
End Function

' Helper function to parse a JSON array
Function ParseArray(jsonText, ByRef charIndex)
    Dim arr, items, i, char, expectComma
    
    ' Create temporary collection to hold items
    Set items = CreateObject("System.Collections.ArrayList")
    
    ' Move past the opening bracket
    charIndex = charIndex + 1
    
    ' Skip whitespace
    SkipWhitespace jsonText, charIndex
    
    ' Check for empty array
    If Mid(jsonText, charIndex, 1) = "]" Then
        charIndex = charIndex + 1
        ' Create empty array
        arr = Array()
        ParseArray = arr
        Exit Function
    End If
    
    ' Parse array items
    expectComma = False
    Do
        ' If expecting a comma, check for it
        If expectComma Then
            SkipWhitespace jsonText, charIndex
            char = Mid(jsonText, charIndex, 1)
            If char <> "," Then
                Err.Raise 9001, "ParseJson", "Expected comma in array at position " & charIndex
            End If
            charIndex = charIndex + 1
        End If
        
        ' Skip whitespace
        SkipWhitespace jsonText, charIndex
        
        ' Check for closing bracket (end of array)
        If Mid(jsonText, charIndex, 1) = "]" Then
            charIndex = charIndex + 1
            Exit Do
        End If
        
        ' Parse value and add to collection
        items.Add ParseValue(jsonText, charIndex)
        
        ' Expect comma for next item
        expectComma = True
    Loop
    
    ' Convert collection to array
    ReDim arr(items.Count - 1)
    For i = 0 To items.Count - 1
        If IsObject(items(i)) Then
            Set arr(i) = items(i)
        Else
            arr(i) = items(i)
        End If
    Next
    
    ' Return array
    ParseArray = arr
End Function

' Helper function to parse a JSON string
Function ParseString(jsonText, ByRef charIndex)
    Dim startPos, endPos, result
    
    ' Move past the opening quote
    charIndex = charIndex + 1
    startPos = charIndex
    
    ' Find the closing quote
    Do While charIndex <= Len(jsonText)
        ' Check for escape sequence
        If Mid(jsonText, charIndex, 1) = "\" Then
            charIndex = charIndex + 2
        ' Check for closing quote
        ElseIf Mid(jsonText, charIndex, 1) = """" Then
            Exit Do
        Else
            charIndex = charIndex + 1
        End If
    Loop
    
    ' Extract string value
    endPos = charIndex - 1
    result = Mid(jsonText, startPos, endPos - startPos + 1)
    
    ' Handle escape sequences
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    result = Replace(result, "\/", "/")
    result = Replace(result, "\b", Chr(8))
    result = Replace(result, "\f", Chr(12))
    result = Replace(result, "\n", vbLf)
    result = Replace(result, "\r", vbCr)
    result = Replace(result, "\t", vbTab)
    
    ' Process Unicode escape sequences (\uXXXX)
    Dim reUnicode, matches, match
    Set reUnicode = New RegExp
    reUnicode.Global = True
    reUnicode.Pattern = "\\u([0-9a-fA-F]{4})"
    
    Set matches = reUnicode.Execute(result)
    For Each match In matches
        Dim hexCode, unicodeChar
        hexCode = match.SubMatches(0)
        unicodeChar = ChrW(CLng("&H" & hexCode))
        result = Replace(result, match.Value, unicodeChar)
    Next
    
    ' Move past the closing quote
    charIndex = charIndex + 1
    
    ' Return the parsed string
    ParseString = result
End Function

' Helper function to parse a JSON boolean value
Function ParseBoolean(jsonText, ByRef charIndex)
    Dim value
    
    ' Check for "true"
    If Mid(jsonText, charIndex, 4) = "true" Then
        value = True
        charIndex = charIndex + 4
    ' Check for "false"
    ElseIf Mid(jsonText, charIndex, 5) = "false" Then
        value = False
        charIndex = charIndex + 5
    Else
        Err.Raise 9001, "ParseJson", "Invalid boolean value at position " & charIndex
    End If
    
    ' Return the parsed boolean
    ParseBoolean = value
End Function

' Helper function to parse a JSON null value
Function ParseNull(jsonText, ByRef charIndex)
    ' Check for "null"
    If Mid(jsonText, charIndex, 4) = "null" Then
        charIndex = charIndex + 4
        ParseNull = Null
    Else
        Err.Raise 9001, "ParseJson", "Invalid null value at position " & charIndex
    End If
End Function

' Helper function to parse a JSON number
Function ParseNumber(jsonText, ByRef charIndex)
    Dim startPos, char, inFraction, inExponent, hasSign
    
    startPos = charIndex
    inFraction = False
    inExponent = False
    hasSign = False
    
    ' Parse characters of the number
    Do While charIndex <= Len(jsonText)
        char = Mid(jsonText, charIndex, 1)
        
        Select Case char
            Case "-", "+"
                ' Sign is only allowed at the start or after 'e' or 'E'
                If charIndex > startPos And Not inExponent Then
                    Exit Do
                End If
                If inExponent And hasSign Then
                    Exit Do
                End If
                If inExponent Then hasSign = True
                
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                ' Digits are always allowed
                
            Case "."
                ' Decimal point is only allowed once and not in exponent
                If inFraction Or inExponent Then
                    Exit Do
                End If
                inFraction = True
                
            Case "e", "E"
                ' Exponent marker is only allowed once and not at the start
                If inExponent Or charIndex = startPos Then
                    Exit Do
                End If
                inExponent = True
                hasSign = False  ' Reset sign flag for the exponent
                
            Case Else
                ' Any other character means we've reached the end of the number
                Exit Do
        End Select
        
        charIndex = charIndex + 1
    Loop
    
    ' Extract the number string
    Dim numStr
    numStr = Mid(jsonText, startPos, charIndex - startPos)
    
    ' Convert to appropriate numeric type
    If InStr(numStr, ".") > 0 Or InStr(LCase(numStr), "e") > 0 Then
        ' Float/double
        ParseNumber = CDbl(numStr)
    Else
        ' Integer
        ParseNumber = CLng(numStr)
    End If
End Function

' Helper function to skip whitespace characters
Sub SkipWhitespace(jsonText, ByRef charIndex)
    Dim char
    
    Do While charIndex <= Len(jsonText)
        char = Mid(jsonText, charIndex, 1)
        If char = " " Or char = vbTab Or char = vbCr Or char = vbLf Then
            charIndex = charIndex + 1
        Else
            Exit Do
        End If
    Loop
End Sub

' Create RegExp object
Function New_RegExp()
    Dim regExp
    Set regExp = CreateObject("VBScript.RegExp")
    Set New_RegExp = regExp
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
