' data_utils.vbs - Utilities for data handling and transformation
' ------------------------------------------------------------

Option Explicit

' Process data from mainframe screen
Function ProcessData(rawData)
    Dim result
    
    ' Remove any trailing spaces and format data
    result = Trim(rawData)
    
    ' Parse screen data into structured format
    result = ParseScreenData(result)
    
    ' Convert to JSON format for API calls
    result = ConvertToJson(result)
    
    ProcessData = result
End Function

' Parse screen data into a structured format
Function ParseScreenData(screenText)
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Example parsing logic - customize for your specific screen format
    Dim lines, line, i
    lines = Split(screenText, vbCrLf)
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Skip empty lines
        If Len(line) > 0 Then
            ' Look for key-value pairs (e.g., "Field Name: Value")
            Dim colonPos, key, value
            colonPos = InStr(line, ":")
            
            If colonPos > 0 Then
                key = Trim(Left(line, colonPos - 1))
                value = Trim(Mid(line, colonPos + 1))
                
                ' Add to dictionary if not empty
                If Len(key) > 0 And Len(value) > 0 Then
                    dict.Add key, value
                End If
            End If
        End If
    Next
    
    Set ParseScreenData = dict
End Function

' Extract specific data field from screen text
Function ExtractField(screenText, fieldName)
    Dim lines, line, i
    lines = Split(screenText, vbCrLf)
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Check if line contains field
        If InStr(line, fieldName) > 0 Then
            Dim parts
            parts = Split(line, ":")
            
            If UBound(parts) >= 1 Then
                ExtractField = Trim(parts(1))
                Exit Function
            End If
        End If
    Next
    
    ' Field not found
    ExtractField = ""
End Function

' Format data for display or output
Function FormatData(data, formatType)
    Select Case UCase(formatType)
        Case "CSV"
            FormatData = ConvertToCSV(data)
        Case "JSON"
            FormatData = ConvertToJson(data)
        Case "XML"
            FormatData = ConvertToXML(data)
        Case "TABLE"
            FormatData = ConvertToTable(data)
        Case Else
            FormatData = data
    End Select
End Function

' Convert dictionary to CSV format
Function ConvertToCSV(data)
    Dim result, key
    
    ' Create header row
    For Each key In data.Keys
        If Len(result) > 0 Then result = result & ","
        result = result & """" & key & """"
    Next
    result = result & vbCrLf
    
    ' Create data row
    Dim dataRow
    For Each key In data.Keys
        If Len(dataRow) > 0 Then dataRow = dataRow & ","
        dataRow = dataRow & """" & data(key) & """"
    Next
    result = result & dataRow
    
    ConvertToCSV = result
End Function

' Convert dictionary to XML format
Function ConvertToXML(data)
    Dim result, key
    
    result = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    result = result & "<data>" & vbCrLf
    
    For Each key In data.Keys
        result = result & "  <" & key & ">" & data(key) & "</" & key & ">" & vbCrLf
    Next
    
    result = result & "</data>"
    
    ConvertToXML = result
End Function

' Convert dictionary to formatted table
Function ConvertToTable(data)
    Dim result, key, maxKeyLength, padding
    
    ' Find the maximum key length for padding
    maxKeyLength = 0
    For Each key In data.Keys
        If Len(key) > maxKeyLength Then maxKeyLength = Len(key)
    Next
    
    ' Add padding for alignment
    maxKeyLength = maxKeyLength + 2
    
    ' Create table header
    result = String(maxKeyLength + 20, "-") & vbCrLf
    
    ' Create table rows
    For Each key In data.Keys
        padding = String(maxKeyLength - Len(key), " ")
        result = result & key & ":" & padding & data(key) & vbCrLf
    Next
    
    ' Add table footer
    result = result & String(maxKeyLength + 20, "-")
    
    ConvertToTable = result
End Function

' Validate data against expected format/rules
Function ValidateData(data, rules)
    Dim isValid, key
    isValid = True
    
    ' Simple validation example
    For Each key In rules.Keys
        ' Check if required field exists
        If rules(key) = "required" Then
            If Not data.Exists(key) Or Trim(data(key)) = "" Then
                isValid = False
                Exit For
            End If
        End If
    Next
    
    ValidateData = isValid
End Function

' Clean and sanitize input data
Function SanitizeData(input)
    Dim result
    
    ' Remove potentially dangerous characters
    result = input
    
    ' Remove script tags
    result = Replace(result, "<script", "")
    result = Replace(result, "</script>", "")
    
    ' Remove special characters
    result = Replace(result, "'", "''")  ' Escape single quotes for SQL
    
    SanitizeData = result
End Function
