Option Explicit

' API key for Google Gemini API
Private Const API_KEY As String = "AIzaSyCWC8fiA6mUX590yF4tSchHxLA7iQBe3BY"
Private Const API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

' References needed:
' Microsoft Scripting Runtime
' Microsoft XML, v6.0 (or available version)

' Main procedure to analyze contract clauses
Public Sub AnalyzeContractClauses()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim contractText As String
    Dim result As String
    Dim tagValue As String
    Dim rawResponse As String
    Dim quotaExceeded As Boolean
    Dim retryCount As Integer
    Dim waitTimeMinutes As Integer
    Dim queuedLimitReached As Boolean
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Get active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Ask user for row range to process
    startRow = InputBox("Enter the starting row number (2 or greater):", "Row Range", 2)
    If startRow < 2 Then startRow = 2
    
    endRow = InputBox("Enter the ending row number (up to " & lastRow & "):", "Row Range", lastRow)
    If endRow > lastRow Then endRow = lastRow
    If endRow < startRow Then endRow = startRow
    
    ' Display progress information
    Application.StatusBar = "Analyzing contract clauses..."
    Application.ScreenUpdating = False
    
    ' Add headers if they don't exist
    If ws.Cells(1, 1).Value = "" Then ws.Cells(1, 1).Value = "Contract Text"
    If ws.Cells(1, 2).Value = "" Then ws.Cells(1, 2).Value = "Analysis Result"
    If ws.Cells(1, 3).Value = "" Then ws.Cells(1, 3).Value = "Determination"
    If ws.Cells(1, 4).Value = "" Then ws.Cells(1, 4).Value = "Raw JSON Response"
    
    ' Loop through each row with content in column A
    quotaExceeded = False
    queuedLimitReached = False
    i = startRow
    
    Do While i <= endRow
        ' Get contract text from column A
        contractText = ws.Cells(i, 1).Value
        
        ' Only process if there's text in the cell
        If Len(Trim(contractText)) > 0 Then
            ' Update status bar
            Application.StatusBar = "Analyzing row " & i & " of " & endRow & "..."
            
            ' Call Gemini API to analyze the text and get both result and raw response
            result = AnalyzeWithGemini(contractText, rawResponse)
            
            ' Check if quota exceeded
            If (InStr(1, result, "quota", vbTextCompare) > 0 And _
               (InStr(1, result, "exceeded", vbTextCompare) > 0 Or _
                InStr(1, result, "exhausted", vbTextCompare) > 0)) Or _
               InStr(1, result, "Rate limit", vbTextCompare) > 0 Then
                
                quotaExceeded = True
                retryCount = 0
                waitTimeMinutes = 1
                
                ' Mark the cell to show we hit a quota limit
                ws.Cells(i, 2).Value = "API QUOTA EXCEEDED - Automatically waiting to retry..."
                ws.Cells(i, 4).Value = rawResponse
                
                ' Auto-retry loop with increasing wait times
                Do While quotaExceeded And retryCount < 5  ' Try up to 5 times with increasing wait
                    ' Update status bar to show waiting
                    Application.StatusBar = "API quota exceeded. Waiting for " & waitTimeMinutes & " minute(s) before retry " & (retryCount + 1) & " of 5..."
                    
                    ' Log the wait to debug
                    Debug.Print "Quota exceeded at row " & i & ". Waiting " & waitTimeMinutes & " minute(s) before retry " & (retryCount + 1) & " of 5..."
                    
                    ' Wait for increasing amounts of time
                    Application.Wait (Now + TimeValue("0:0" & waitTimeMinutes & ":00"))
                    
                    ' Try again
                    result = AnalyzeWithGemini(contractText, rawResponse)
                    
                    ' Check if we're still quota limited
                    If (InStr(1, result, "quota", vbTextCompare) > 0 And _
                       (InStr(1, result, "exceeded", vbTextCompare) > 0 Or _
                        InStr(1, result, "exhausted", vbTextCompare) > 0)) Or _
                       InStr(1, result, "Rate limit", vbTextCompare) > 0 Then
                        
                        ' Still limited, increase wait time for next try
                        retryCount = retryCount + 1
                        waitTimeMinutes = waitTimeMinutes * 2  ' Exponential backoff
                        
                        ' Special case: if we detect daily quota issues
                        If InStr(1, result, "daily quota", vbTextCompare) > 0 Then
                            queuedLimitReached = True
                            Application.StatusBar = "Daily quota reached. Processing will stop."
                            Exit Do
                        End If
                    Else
                        ' Success! We got past the quota limit
                        quotaExceeded = False
                    End If
                Loop
                
                ' If we still have quota issues after all retries, skip this row
                If quotaExceeded Then
                    If queuedLimitReached Then
                        MsgBox "Daily API quota has been reached. Processing will stop at row " & i & "." & vbCrLf & _
                               "Please try again tomorrow starting from this row.", vbExclamation
                        Exit Do
                    Else
                        ws.Cells(i, 2).Value = "API QUOTA EXCEEDED - Skipped after 5 retries"
                        ws.Cells(i, 4).Value = rawResponse
                        i = i + 1
                        GoTo NextRowInMainLoop  ' Skip to next iteration
                    End If
                End If
            End If
            
            ' Write result to column B
            ws.Cells(i, 2).Value = result
            
            ' Store raw response in column D for debugging if there's an error
            If InStr(1, result, "Error", vbTextCompare) > 0 Then
                ws.Cells(i, 4).Value = rawResponse
            End If
            
            ' Extract the determination part
            If InStr(1, result, "UNCAPPED LIABILITY FOUND (Primary):") > 0 Then
                tagValue = "UNCAPPED LIABILITY FOUND (Primary)"
            ElseIf InStr(1, result, "UNCAPPED LIABILITY FOUND:") > 0 Then
                tagValue = "UNCAPPED LIABILITY FOUND"
            ElseIf InStr(1, result, "CAPPED LIABILITY:") > 0 Then
                tagValue = "CAPPED LIABILITY"
            ElseIf InStr(1, result, "UNCERTAIN:") > 0 Then
                tagValue = "UNCERTAIN"
            Else
                ' Look for the response formats with asterisks
                If InStr(1, result, "**UNCAPPED LIABILITY FOUND (Primary)**") > 0 Then
                    tagValue = "UNCAPPED LIABILITY FOUND (Primary)"
                ElseIf InStr(1, result, "**UNCAPPED LIABILITY FOUND**") > 0 Then
                    tagValue = "UNCAPPED LIABILITY FOUND"
                ElseIf InStr(1, result, "**CAPPED LIABILITY**") > 0 Then
                    tagValue = "CAPPED LIABILITY"
                ElseIf InStr(1, result, "**UNCERTAIN**") > 0 Then
                    tagValue = "UNCERTAIN"
                Else
                    ' If there's no standard format found, use the beginning of the text
                    If InStr(1, result, "Error", vbTextCompare) > 0 Then
                        tagValue = "Error"
                    Else
                        ' Try to extract a useful part
                        If Len(result) > 30 Then
                            tagValue = Left(result, 30) & "..."
                        Else
                            tagValue = result
                        End If
                    End If
                End If
            End If
            
            ' For debugging
            Debug.Print "Row " & i & ": Result prefix = " & Left(result, 30) & ", Tag = " & tagValue
            
            ' Write tag to column C
            ws.Cells(i, 3).Value = tagValue
            
            ' Add a small delay between API calls to avoid hitting rate limits
            Application.Wait (Now + TimeValue("0:00:02"))
        End If
        
        i = i + 1
NextRowInMainLoop:
    Loop
    
    ' Reset status bar and screen updating
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ' Show completion message
    If queuedLimitReached Then
        MsgBox "Analysis stopped due to daily API quota limitations. Last processed row: " & (i - 1) & vbCrLf & _
               "Please try again tomorrow starting from row " & i & ".", vbInformation
    Else
        MsgBox "Analysis complete! Processed rows " & startRow & " to " & endRow & ".", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description & " at row " & i, vbCritical
End Sub

' Function to analyze text using Gemini API
Private Function AnalyzeWithGemini(textToAnalyze As String, Optional ByRef rawResponse As String = "") As String
    Dim http As Object
    Dim jsonResponse As String
    Dim requestBody As String
    Dim promptPart1 As String
    Dim promptPart2 As String
    Dim promptPart3 As String
    Dim promptPart4 As String
    Dim fullUrl As String
    Dim responseText As String
    Dim statusCode As Long
    
    ' Set up HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    fullUrl = API_URL & "?key=" & API_KEY
    
    ' Break up the prompt into smaller parts to avoid too many line continuations
    promptPart1 = "You are a highly experienced contract analysis expert specializing in liability clauses. Your task is to analyze the following contract text and strictly determine if it contains provisions that create UNCAPPED liability applying to the ENTIRE contract, *considering the agreement as a whole*. "
    
    promptPart2 = "Ignore standard and customary exceptions for specific situations like fraud, gross negligence, willful misconduct, death, personal injury, statutory violations, or situations where limitations are legally unenforceable. These exceptions are common and DO NOT automatically make a contract 'uncapped'. Focus on whether, WITHOUT THESE STANDARD EXCEPTIONS, the core agreement still lacks a definitive, overarching monetary limit on potential liability."
    
    promptPart3 = "You MUST choose ONE of the following options, adhering to the specified format exactly:" & _
                  "\n\n1. **UNCAPPED LIABILITY FOUND (Primary):** If, after excluding standard exceptions, the agreement LACKS a clear, overarching monetary limit applicable to the majority of potential liabilities. Provide a brief explanation of why the core agreement is uncapped, citing the relevant clauses." & _
                  "\n\n2. **UNCAPPED LIABILITY FOUND:** If standard exceptions are so pervasive and broad in scope that they swallow the general liability cap, rendering it meaningless for a significant portion of potential liabilities. Provide a brief explanation." & _
                  "\n\n3. **CAPPED LIABILITY:** If the agreement contains a clear, overarching monetary limit (either a fixed sum or a readily calculable amount) applicable to the majority of potential liabilities, even if standard exceptions exist. Provide a brief explanation of the primary cap and why, despite the exceptions, it remains a meaningful limit." & _
                  "\n\n4. **UNCERTAIN:** If the contract text is so ambiguous or incomplete that you cannot confidently determine whether it is capped or uncapped, even after your expert analysis. Explain the specific ambiguities and why a determination is impossible."
    
    promptPart4 = "\n\nImportant Notes:\n* Overemphasizing commonly-used exceptions (fraud, gross negligence) will be considered an incorrect output.\n* Ensure your output is JUST the chosen response.\n* Always justify your response with reference to specific clause numbers.\n\nHere is the contract text to analyze:\n\n" & JsonEscape(textToAnalyze)
    
    ' Combine the prompt parts
    Dim fullPrompt As String
    fullPrompt = promptPart1 & "\n\n" & promptPart2 & "\n\n" & promptPart3 & promptPart4
    
    ' Prepare the request body (JSON) with the prompt
    requestBody = "{" & _
                 """contents"": [" & _
                 "{""parts"":[{""text"": """ & fullPrompt & """}]}" & _
                 "]," & _
                 """safetySettings"": [" & _
                 "{""category"": ""HARM_CATEGORY_DANGEROUS_CONTENT"", ""threshold"": ""BLOCK_ONLY_HIGH""}" & _
                 "]" & _
                 "}"
    
    ' Send the request
    On Error Resume Next
    http.Open "POST", fullUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody
    
    ' Check for errors in the HTTP request
    If Err.Number <> 0 Then
        rawResponse = "HTTP Error: " & Err.Description
        AnalyzeWithGemini = "Error in HTTP request: " & Err.Description
        Exit Function
    End If
    
    ' Get the status code
    statusCode = http.Status
    
    ' Process the response
    responseText = http.responseText
    
    ' Store raw response for debugging
    rawResponse = responseText
    
    ' Check if the response is successful
    If statusCode <> 200 Then
        AnalyzeWithGemini = "API Error: Status Code " & statusCode & ". " & responseText
        Exit Function
    End If
    
    ' For debugging purposes, you can uncomment this to see the raw response
    Debug.Print "Raw Response (first 500 chars): " & Left(responseText, 500)
    
    ' Extract text from the JSON response
    Dim extractedText As String
    extractedText = ExtractTextFromGeminiResponse(responseText)
    
    If extractedText <> "" Then
        AnalyzeWithGemini = extractedText
    Else
        ' If extraction failed, return the error
        AnalyzeWithGemini = "Error parsing JSON response: Could not extract text"
    End If
    
    On Error GoTo 0
End Function

' Helper function to escape JSON strings
Private Function JsonEscape(text As String) As String
    Dim result As String
    
    result = Replace(text, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    
    JsonEscape = result
End Function

' Function to extract text from Gemini API response
Private Function ExtractTextFromGeminiResponse(jsonString As String) As String
    ' This improved extractor includes better error handling and debugging
    
    On Error Resume Next
    
    Dim result As String
    result = ""
    
    ' Check if the JSON response is empty or null
    If Len(Trim(jsonString)) = 0 Then
        Debug.Print "Error: Empty JSON response"
        ExtractTextFromGeminiResponse = ""
        Exit Function
    End If
    
    ' Check if the response contains an error message
    If InStr(1, jsonString, """error"":", vbTextCompare) > 0 Then
        ' Extract the error message if present
        Dim errorStart As Long
        Dim errorEnd As Long
        Dim errorMsg As String
        
        errorStart = InStr(1, jsonString, """message"": """) + 12
        If errorStart > 12 Then  ' Found a message field
            errorEnd = InStr(errorStart, jsonString, """")
            If errorEnd > 0 Then
                errorMsg = Mid(jsonString, errorStart, errorEnd - errorStart)
                Debug.Print "API Error: " & errorMsg
                ExtractTextFromGeminiResponse = "API Error: " & errorMsg
                Exit Function
            End If
        End If
    End If
    
    ' Look for the text field in the parts array
    Dim textPattern As String
    textPattern = """text"": """
    
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, jsonString, textPattern)
    If startPos > 0 Then
        ' Move past the pattern
        startPos = startPos + Len(textPattern)
        
        ' Find the closing quote, being careful about escaped quotes
        Dim currentPos As Long
        Dim foundEnd As Boolean
        
        currentPos = startPos
        foundEnd = False
        
        Do While Not foundEnd And currentPos <= Len(jsonString)
            ' Look for next quote
            endPos = InStr(currentPos, jsonString, """")
            
            If endPos = 0 Then
                ' No more quotes found
                foundEnd = True
                Debug.Print "Warning: Could not find closing quote in JSON response"
            Else
                ' Check if this quote is escaped
                If Mid(jsonString, endPos - 1, 1) = "\" Then
                    ' This is an escaped quote, continue searching
                    currentPos = endPos + 1
                Else
                    ' Found the closing quote
                    foundEnd = True
                    result = Mid(jsonString, startPos, endPos - startPos)
                End If
            End If
        Loop
    Else
        ' Could not find the text pattern, try an alternative approach
        Debug.Print "Could not find 'text' field in response, trying alternative extraction..."
        
        ' Try to extract from finishReason section which might indicate an error
        Dim finishReasonStart As Long
        finishReasonStart = InStr(1, jsonString, """finishReason"":")
        If finishReasonStart > 0 Then
            Debug.Print "Found finishReason in response: " & Mid(jsonString, finishReasonStart, 30)
        End If
        
        ' Look for any text content in a more generic way
        Dim contentStart As Long
        contentStart = InStr(1, jsonString, """content"":")
        If contentStart > 0 Then
            Debug.Print "Found content field at position: " & contentStart
        End If
    End If
    
    ' Unescape the JSON string if we found something
    If result <> "" Then
        ' Handle basic escaping
        result = Replace(result, "\""", """")
        result = Replace(result, "\\", "\")
        result = Replace(result, "\n", vbNewLine)
        result = Replace(result, "\r", vbCr)
        result = Replace(result, "\t", vbTab)
    End If
    
    ' Return result
    ExtractTextFromGeminiResponse = result
    
    On Error GoTo 0
End Function

' Batch processing function to handle API quota limitations
Public Sub AnalyzeContractClausesInBatches()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim batchSize As Long
    Dim pauseMinutes As Long
    Dim continueProcessing As Boolean
    Dim processingResult As Boolean
    
    ' Get active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Get batch parameters from user
    startRow = InputBox("Enter the starting row number (2 or greater):", "Batch Processing", 2)
    If startRow < 2 Then startRow = 2
    If startRow > lastRow Then
        MsgBox "Starting row is beyond the end of data. Operation canceled.", vbExclamation
        Exit Sub
    End If
    
    batchSize = InputBox("Enter the batch size (recommended 10-15 contracts per batch):", "Batch Processing", 10)
    If batchSize <= 0 Then batchSize = 10
    
    pauseMinutes = InputBox("Enter the pause time between batches in minutes (recommended 2-5 minutes):", "Batch Processing", 2)
    If pauseMinutes < 1 Then pauseMinutes = 1
    
    ' Add headers if they don't exist
    If ws.Cells(1, 1).Value = "" Then ws.Cells(1, 1).Value = "Contract Text"
    If ws.Cells(1, 2).Value = "" Then ws.Cells(1, 2).Value = "Analysis Result"
    If ws.Cells(1, 3).Value = "" Then ws.Cells(1, 3).Value = "Determination"
    If ws.Cells(1, 4).Value = "" Then ws.Cells(1, 4).Value = "Raw JSON Response"
    
    ' Process in batches
    continueProcessing = True
    endRow = startRow + batchSize - 1
    If endRow > lastRow Then endRow = lastRow
    
    Do While continueProcessing And startRow <= lastRow
        ' Process the current batch
        Application.StatusBar = "Processing batch from row " & startRow & " to " & endRow & "..."
        
        ' Call the batch processing function with the batch range
        processingResult = ProcessBatch(startRow, endRow)
        
        ' Check if processing was stopped due to severe quota limit
        If Not processingResult Then
            MsgBox "Batch processing stopped due to daily API quota limitations." & vbCrLf & _
                   "Please wait at least 24 hours before trying again from row " & startRow & ".", vbExclamation
            continueProcessing = False
        Else
            ' Update for next batch
            startRow = endRow + 1
            endRow = startRow + batchSize - 1
            If endRow > lastRow Then endRow = lastRow
            
            ' If we're at the end, we're done
            If startRow > lastRow Then
                continueProcessing = False
            Else
                ' Otherwise, wait and continue automatically
                Application.StatusBar = "Batch complete. Waiting " & pauseMinutes & " minutes before processing next batch (rows " & startRow & " to " & endRow & ")..."
                Application.Wait (Now + TimeValue("0:0" & pauseMinutes & ":00"))
            End If
        End If
    Loop
    
    Application.StatusBar = False
    
    If Not continueProcessing And startRow <= lastRow Then
        MsgBox "Batch processing stopped at row " & startRow & ".", vbInformation
    Else
        MsgBox "All batches processed successfully!", vbInformation
    End If
End Sub

' Function to process a specific range of rows
Private Function ProcessBatch(startRow As Long, endRow As Long) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim contractText As String
    Dim result As String
    Dim tagValue As String
    Dim rawResponse As String
    Dim quotaExceeded As Boolean
    Dim retryCount As Integer
    Dim waitTimeMinutes As Integer
    Dim dailyQuotaReached As Boolean
    
    ' Default to success
    ProcessBatch = True
    
    ' Get active worksheet
    Set ws = ActiveSheet
    
    ' Process each row in the batch
    For i = startRow To endRow
        ' Get contract text from column A
        contractText = ws.Cells(i, 1).Value
        
        ' Only process if there's text in the cell
        If Len(Trim(contractText)) > 0 Then
            ' Update status bar
            Application.StatusBar = "Analyzing row " & i & " of " & endRow & " (batch)..."
            
            ' Call Gemini API to analyze the text and get both result and raw response
            result = AnalyzeWithGemini(contractText, rawResponse)
            
            ' Check if quota exceeded
            If (InStr(1, result, "quota", vbTextCompare) > 0 And _
               (InStr(1, result, "exceeded", vbTextCompare) > 0 Or _
                InStr(1, result, "exhausted", vbTextCompare) > 0)) Or _
               InStr(1, result, "Rate limit", vbTextCompare) > 0 Then
                
                quotaExceeded = True
                retryCount = 0
                waitTimeMinutes = 1
                dailyQuotaReached = False
                
                ' Mark the cell to show we're retrying
                ws.Cells(i, 2).Value = "API QUOTA EXCEEDED - Automatically waiting to retry..."
                ws.Cells(i, 4).Value = rawResponse
                
                ' Auto-retry loop
                Do While quotaExceeded And retryCount < 5
                    ' Update status bar to show waiting
                    Application.StatusBar = "API quota exceeded. Waiting for " & waitTimeMinutes & " minute(s) before retry " & (retryCount + 1) & " of 5..."
                    
                    ' Log the wait
                    Debug.Print "Quota exceeded at row " & i & ". Waiting " & waitTimeMinutes & " minute(s) before retry " & (retryCount + 1) & " of 5..."
                    
                    ' Wait for the specified time
                    Application.Wait (Now + TimeValue("0:0" & waitTimeMinutes & ":00"))
                    
                    ' Try again
                    result = AnalyzeWithGemini(contractText, rawResponse)
                    
                    ' Check if still rate limited
                    If (InStr(1, result, "quota", vbTextCompare) > 0 And _
                       (InStr(1, result, "exceeded", vbTextCompare) > 0 Or _
                        InStr(1, result, "exhausted", vbTextCompare) > 0)) Or _
                       InStr(1, result, "Rate limit", vbTextCompare) > 0 Then
                        
                        ' Still limited, increase time for next retry
                        retryCount = retryCount + 1
                        waitTimeMinutes = waitTimeMinutes * 2  ' Exponential backoff
                        
                        ' Check if it's a daily quota issue
                        If InStr(1, result, "daily quota", vbTextCompare) > 0 Then
                            dailyQuotaReached = True
                            Application.StatusBar = "Daily quota reached. Batch processing will stop."
                            Exit Do
                        End If
                    Else
                        ' Success! We're past the quota limitation
                        quotaExceeded = False
                    End If
                Loop
                
                ' If we still have quota issues after all retries
                If quotaExceeded Then
                    If dailyQuotaReached Then
                        ' This is a serious daily quota issue
                        ws.Cells(i, 2).Value = "DAILY API QUOTA EXCEEDED - Processing stopped"
                        ws.Cells(i, 4).Value = rawResponse
                        ProcessBatch = False  ' Signal that we should stop all processing
                        Exit Function
                    Else
                        ' Skip this particular row after maximum retries
                        ws.Cells(i, 2).Value = "API QUOTA EXCEEDED - Skipped after 5 retries"
                        ws.Cells(i, 4).Value = rawResponse
                        GoTo NextIteration  ' Skip to next iteration
                    End If
                End If
            End If
            
            ' Write result to column B
            ws.Cells(i, 2).Value = result
            
            ' Store raw response in column D for debugging if there's an error
            If InStr(1, result, "Error", vbTextCompare) > 0 Then
                ws.Cells(i, 4).Value = rawResponse
            End If
            
            ' Extract the determination part
            If InStr(1, result, "UNCAPPED LIABILITY FOUND (Primary):") > 0 Then
                tagValue = "UNCAPPED LIABILITY FOUND (Primary)"
            ElseIf InStr(1, result, "UNCAPPED LIABILITY FOUND:") > 0 Then
                tagValue = "UNCAPPED LIABILITY FOUND"
            ElseIf InStr(1, result, "CAPPED LIABILITY:") > 0 Then
                tagValue = "CAPPED LIABILITY"
            ElseIf InStr(1, result, "UNCERTAIN:") > 0 Then
                tagValue = "UNCERTAIN"
            Else
                ' Look for the response formats with asterisks
                If InStr(1, result, "**UNCAPPED LIABILITY FOUND (Primary)**") > 0 Then
                    tagValue = "UNCAPPED LIABILITY FOUND (Primary)"
                ElseIf InStr(1, result, "**UNCAPPED LIABILITY FOUND**") > 0 Then
                    tagValue = "UNCAPPED LIABILITY FOUND"
                ElseIf InStr(1, result, "**CAPPED LIABILITY**") > 0 Then
                    tagValue = "CAPPED LIABILITY"
                ElseIf InStr(1, result, "**UNCERTAIN**") > 0 Then
                    tagValue = "UNCERTAIN"
                Else
                    ' If there's no standard format found, use the beginning of the text
                    If InStr(1, result, "Error", vbTextCompare) > 0 Then
                        tagValue = "Error"
                    Else
                        ' Try to extract a useful part
                        If Len(result) > 30 Then
                            tagValue = Left(result, 30) & "..."
                        Else
                            tagValue = result
                        End If
                    End If
                End If
            End If
            
            ' For debugging
            Debug.Print "Row " & i & ": Result prefix = " & Left(result, 30) & ", Tag = " & tagValue
            
            ' Write tag to column C
            ws.Cells(i, 3).Value = tagValue
            
            ' Add a small delay between API calls to avoid hitting rate limits
            Application.Wait (Now + TimeValue("0:00:02"))
        End If
NextIteration:
    Next i
    
    ' If we get here, the batch completed successfully
    ProcessBatch = True
End Function 