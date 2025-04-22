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
    Dim contractText As String
    Dim result As String
    Dim tagValue As String
    Dim rawResponse As String
    Dim colonPosition As Long
    Dim determinationText As String
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Get active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Display progress information
    Application.StatusBar = "Analyzing contract clauses..."
    Application.ScreenUpdating = False
    
    ' Add headers if they don't exist
    If ws.Cells(1, 1).Value = "" Then ws.Cells(1, 1).Value = "Contract Text"
    If ws.Cells(1, 2).Value = "" Then ws.Cells(1, 2).Value = "Analysis Result"
    If ws.Cells(1, 3).Value = "" Then ws.Cells(1, 3).Value = "Determination"
    If ws.Cells(1, 4).Value = "" Then ws.Cells(1, 4).Value = "Raw JSON Response"
    
    ' Loop through each row with content in column A
    For i = 2 To lastRow ' Assuming row 1 contains headers
        ' Get contract text from column A
        contractText = ws.Cells(i, 1).Value
        
        ' Only process if there's text in the cell
        If Len(Trim(contractText)) > 0 Then
            ' Update status bar
            Application.StatusBar = "Analyzing row " & i & " of " & lastRow & "..."
            
            ' Call Gemini API to analyze the text and get both result and raw response
            result = AnalyzeWithGemini(contractText, rawResponse)
            
            ' Write result to column B
            ws.Cells(i, 2).Value = result
            
            ' Store raw response in column D for debugging if there's an error
            If InStr(1, result, "Error", vbTextCompare) > 0 Then
                ws.Cells(i, 4).Value = rawResponse
            End If
            
            ' Extract the determination part (text before the colon)
            If InStr(1, result, ":") > 0 Then
                colonPosition = InStr(1, result, ":")
                determinationText = Trim(Left(result, colonPosition - 1))
                
                ' Store the exact determination text rather than Yes/No/Uncertain
                tagValue = determinationText
            Else
                ' If there's no colon, use the first 30 characters or "Error/Uncertain"
                If InStr(1, result, "Error", vbTextCompare) > 0 Then
                    tagValue = "Error"
                Else
                    ' Try to extract a useful part if there's no colon
                    If Len(result) > 30 Then
                        tagValue = Left(result, 30) & "..."
                    Else
                        tagValue = result
                    End If
                End If
            End If
            
            ' For debugging
            Debug.Print "Row " & i & ": Result prefix = " & Left(result, 30) & ", Tag = " & tagValue
            
            ' Write tag to column C
            ws.Cells(i, 3).Value = tagValue
        End If
    Next i
    
    ' Reset status bar and screen updating
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Analysis complete!", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Function to analyze text using Gemini API
Private Function AnalyzeWithGemini(textToAnalyze As String, Optional ByRef rawResponse As String = "") As String
    Dim http As Object
    Dim jsonResponse As String
    Dim requestBody As String
    Dim fullUrl As String
    Dim responseText As String
    Dim statusCode As Long
    
    ' Set up HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    fullUrl = API_URL & "?key=" & API_KEY
    
    ' Prepare the request body (JSON) with more specific instructions
    requestBody = "{" & _
                 """contents"": [" & _
                 "{""parts"":[{""text"": ""You are a contract analysis expert. Please analyze the following contract clause and determine whether it creates an uncapped limitation of liability that applies to the entire contract as a whole. Ignore exceptions that apply only to specific situations (e.g., fraud, death, or statutory requirements). I need to know if there is any language that results in the absence of a clear overall liability cap for the entire agreement.\n\nRespond with EXACTLY one of these formats:\n\n1. If you find uncapped liability: 'UNCAPPED LIABILITY FOUND: [brief explanation]'\n2. If liability is capped: 'No uncapped liability found: [brief explanation]'\n3. If you cannot determine: 'UNCERTAIN: [explanation why you cannot determine]'\n\nHere is the contract text to analyze:\n\n" & JsonEscape(textToAnalyze) & """}]}" & _
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

' Creates a UI for the tool
Public Sub ShowContractAnalyzerUI()
    ' Create a simple form
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = UserForm
    
    ' Set form properties
    With frm
        .Properties("Caption") = "Contract Clause Analyzer"
        .Properties("Width") = 300
        .Properties("Height") = 150
    End With
    
    ' Add a label
    Dim lblDescription As Object
    Set lblDescription = frm.Designer.Controls.Add("Forms.Label.1")
    With lblDescription
        .Left = 20
        .Top = 20
        .Width = 260
        .Height = 40
        .Caption = "This tool will analyze contract clauses in column A for uncapped liability and output results in column B."
    End With
    
    ' Add a button
    Dim btnStart As Object
    Set btnStart = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnStart
        .Left = 100
        .Top = 80
        .Width = 100
        .Height = 30
        .Caption = "Start Analysis"
        .Name = "btnStart"
    End With
    
    ' Add code to the button
    Dim buttonCode As String
    buttonCode = "Private Sub btnStart_Click()" & vbCrLf & _
                 "    Unload Me" & vbCrLf & _
                 "    Call AnalyzeContractClauses" & vbCrLf & _
                 "End Sub"
    
    frm.CodeModule.AddFromString buttonCode
    
    ' Show the form
    frm.Name = "frmContractAnalyzer"
    
    ' Create a procedure to show the form
    Dim showFormCode As String
    showFormCode = "Sub LaunchContractAnalyzer()" & vbCrLf & _
                  "    frmContractAnalyzer.Show" & vbCrLf & _
                  "End Sub"
    
    ThisWorkbook.VBProject.VBComponents.Add(1).CodeModule.AddFromString showFormCode
End Sub

' Test function to verify JSON extraction
Public Sub TestJsonExtraction()
    Dim sampleJson As String
    Dim extractedText As String
    
    ' Sample JSON based on the actual response format
    sampleJson = "{" & _
      """candidates"": [" & _
        "{" & _
          """content"": {" & _
            """parts"": [" & _
              "{" & _
                """text"": ""Please provide the contract text you want me to analyze. I need the limitation of liability clause (and any related clauses that define scope or exceptions) to determine if uncapped liability exists.\n\nOnce you provide the text, I will analyze it and respond with either:\n\n*   **UNCAPPED LIABILITY FOUND**\n*   **No uncapped liability found**\n*   **Explanation of uncertainty** (if applicable, explaining why I can't definitively determine if liability is capped)\n""" & _
              "}" & _
            "]," & _
            """role"": ""model""" & _
          "}," & _
          """finishReason"": ""STOP""," & _
          """avgLogprobs"": -0.33306616179797116" & _
        "}" & _
      "]," & _
      """usageMetadata"": {" & _
        """promptTokenCount"": 63," & _
        """candidatesTokenCount"": 98," & _
        """totalTokenCount"": 161" & _
      "}," & _
      """modelVersion"": ""gemini-2.0-flash""" & _
    "}"
    
    ' Test extraction
    extractedText = ExtractTextFromGeminiResponse(sampleJson)
    
    ' Show result
    If extractedText <> "" Then
        MsgBox "Extraction successful! Extracted text:" & vbCrLf & vbCrLf & extractedText, vbInformation
    Else
        MsgBox "Extraction failed!", vbCritical
    End If
End Sub

' Test API connection with a sample interface
Public Sub TestApiConnection()
    ' Create a simple form
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = UserForm
    
    ' Set form properties
    With frm
        .Properties("Caption") = "Test Gemini API Connection"
        .Properties("Width") = 500
        .Properties("Height") = 400
    End With
    
    ' Add a label for instructions
    Dim lblInstructions As Object
    Set lblInstructions = frm.Designer.Controls.Add("Forms.Label.1")
    With lblInstructions
        .Left = 20
        .Top = 20
        .Width = 460
        .Height = 40
        .Caption = "Enter a sample contract clause below to test the API connection. This will help verify that the API is working properly."
    End With
    
    ' Add a text box for input
    Dim txtInput As Object
    Set txtInput = frm.Designer.Controls.Add("Forms.TextBox.1")
    With txtInput
        .Left = 20
        .Top = 70
        .Width = 460
        .Height = 100
        .MultiLine = True
        .Name = "txtInput"
        .Value = "In no event shall either party be liable for any indirect, incidental, special, punitive, or consequential damages. Direct damages shall be limited to the total amount paid under this agreement."
    End With
    
    ' Add a label for results
    Dim lblResultsCaption As Object
    Set lblResultsCaption = frm.Designer.Controls.Add("Forms.Label.1")
    With lblResultsCaption
        .Left = 20
        .Top = 180
        .Width = 460
        .Height = 20
        .Caption = "API Response:"
        .Name = "lblResultsCaption"
    End With
    
    ' Add a text box for results
    Dim txtResults As Object
    Set txtResults = frm.Designer.Controls.Add("Forms.TextBox.1")
    With txtResults
        .Left = 20
        .Top = 200
        .Width = 460
        .Height = 100
        .MultiLine = True
        .Name = "txtResults"
    End With
    
    ' Add a "Test API" button
    Dim btnTest As Object
    Set btnTest = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnTest
        .Left = 100
        .Top = 320
        .Width = 100
        .Height = 30
        .Caption = "Test API"
        .Name = "btnTest"
    End With
    
    ' Add a "Close" button
    Dim btnClose As Object
    Set btnClose = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnClose
        .Left = 300
        .Top = 320
        .Width = 100
        .Height = 30
        .Caption = "Close"
        .Name = "btnClose"
    End With
    
    ' Add code to the buttons
    Dim buttonCode As String
    buttonCode = "Private Sub btnTest_Click()" & vbCrLf & _
                 "    Dim result As String" & vbCrLf & _
                 "    Dim contractText As String" & vbCrLf & _
                 "    " & vbCrLf & _
                 "    ' Get text from the input box" & vbCrLf & _
                 "    contractText = Me.txtInput.Value" & vbCrLf & _
                 "    " & vbCrLf & _
                 "    ' Only process if there's text" & vbCrLf & _
                 "    If Len(Trim(contractText)) > 0 Then" & vbCrLf & _
                 "        ' Show processing message" & vbCrLf & _
                 "        Me.txtResults.Value = ""Processing... Please wait.""" & vbCrLf & _
                 "        " & vbCrLf & _
                 "        ' Call API to analyze the text" & vbCrLf & _
                 "        Dim rawResponse As String" & vbCrLf & _
                 "        result = AnalyzeWithGemini(contractText, rawResponse)" & vbCrLf & _
                 "        " & vbCrLf & _
                 "        ' Show result" & vbCrLf & _
                 "        Me.txtResults.Value = result" & vbCrLf & _
                 "    Else" & vbCrLf & _
                 "        MsgBox ""Please enter some text to analyze."", vbExclamation" & vbCrLf & _
                 "    End If" & vbCrLf & _
                 "End Sub" & vbCrLf & _
                 "" & vbCrLf & _
                 "Private Sub btnClose_Click()" & vbCrLf & _
                 "    Unload Me" & vbCrLf & _
                 "End Sub"
    
    frm.CodeModule.AddFromString buttonCode
    
    ' Name the form
    frm.Name = "frmTestApi"
    
    ' Show the form
    Dim showFormCode As String
    showFormCode = "Sub ShowTestApiForm()" & vbCrLf & _
                   "    frmTestApi.Show" & vbCrLf & _
                   "End Sub"
    
    ' Add the sub to a new module
    ThisWorkbook.VBProject.VBComponents.Add(1).CodeModule.AddFromString showFormCode
    
    ' Launch the form now
    Application.Run "ShowTestApiForm"
End Sub

' Test procedure to validate tag assignment logic
Public Sub TestTagAssignment()
    Dim result1 As String
    Dim result2 As String
    Dim result3 As String
    Dim tag1 As String
    Dim tag2 As String
    Dim tag3 As String
    
    ' Sample results
    result1 = "UNCAPPED LIABILITY FOUND: This contract has no overall cap on liability."
    result2 = "No uncapped liability found: Liability is capped at $1 million."
    result3 = "UNCERTAIN: Cannot determine due to ambiguous language."
    
    ' Test tag assignment logic
    If InStr(1, result1, "UNCAPPED LIABILITY FOUND", vbTextCompare) > 0 Then
        tag1 = "Yes"
    ElseIf InStr(1, result1, "No uncapped liability found", vbTextCompare) > 0 Then
        tag1 = "No"
    Else
        tag1 = "Uncertain"
    End If
    
    If InStr(1, result2, "UNCAPPED LIABILITY FOUND", vbTextCompare) > 0 Then
        tag2 = "Yes"
    ElseIf InStr(1, result2, "No uncapped liability found", vbTextCompare) > 0 Then
        tag2 = "No"
    Else
        tag2 = "Uncertain"
    End If
    
    If InStr(1, result3, "UNCAPPED LIABILITY FOUND", vbTextCompare) > 0 Then
        tag3 = "Yes"
    ElseIf InStr(1, result3, "No uncapped liability found", vbTextCompare) > 0 Then
        tag3 = "No"
    Else
        tag3 = "Uncertain"
    End If
    
    ' Display results
    Debug.Print "Test 1: " & result1 & " -> Tag: " & tag1
    Debug.Print "Test 2: " & result2 & " -> Tag: " & tag2
    Debug.Print "Test 3: " & result3 & " -> Tag: " & tag3
    
    ' Show results in a message box
    MsgBox "Test 1: " & result1 & vbCrLf & "Tag: " & tag1 & vbCrLf & vbCrLf & _
           "Test 2: " & result2 & vbCrLf & "Tag: " & tag2 & vbCrLf & vbCrLf & _
           "Test 3: " & result3 & vbCrLf & "Tag: " & tag3, _
           vbInformation, "Tag Assignment Test Results"
End Sub 