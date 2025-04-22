Option Explicit

' Creates a UI for the tool
Public Sub ShowContractAnalyzerUI()
    ' Create a simple form
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = UserForm
    
    ' Set form properties
    With frm
        .Properties("Caption") = "Contract Clause Analyzer"
        .Properties("Width") = 400
        .Properties("Height") = 250
    End With
    
    ' Add a label for the title
    Dim lblTitle As Object
    Set lblTitle = frm.Designer.Controls.Add("Forms.Label.1")
    With lblTitle
        .Left = 20
        .Top = 20
        .Width = 360
        .Height = 20
        .Caption = "Contract Clause Analyzer - Liability Cap Detection"
        .Name = "lblTitle"
    End With
    
    ' Add a label for the description
    Dim lblDescription As Object
    Set lblDescription = frm.Designer.Controls.Add("Forms.Label.1")
    With lblDescription
        .Left = 20
        .Top = 50
        .Width = 360
        .Height = 60
        .Caption = "This tool analyzes contract clauses in column A for uncapped liability. Select a processing method below to begin analysis:"
        .Name = "lblDescription"
    End With
    
    ' Add a button for standard processing
    Dim btnStandard As Object
    Set btnStandard = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnStandard
        .Left = 20
        .Top = 120
        .Width = 170
        .Height = 30
        .Caption = "Standard Processing"
        .Name = "btnStandard"
    End With
    
    ' Add a button for batch processing
    Dim btnBatch As Object
    Set btnBatch = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnBatch
        .Left = 210
        .Top = 120
        .Width = 170
        .Height = 30
        .Caption = "Batch Processing"
        .Name = "btnBatch"
    End With
    
    ' Add a button for API testing
    Dim btnTest As Object
    Set btnTest = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnTest
        .Left = 20
        .Top = 160
        .Width = 170
        .Height = 30
        .Caption = "Test API Connection"
        .Name = "btnTest"
    End With
    
    ' Add a close button
    Dim btnClose As Object
    Set btnClose = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnClose
        .Left = 210
        .Top = 160
        .Width = 170
        .Height = 30
        .Caption = "Close"
        .Name = "btnClose"
    End With
    
    ' Add a label for quota info
    Dim lblQuota As Object
    Set lblQuota = frm.Designer.Controls.Add("Forms.Label.1")
    With lblQuota
        .Left = 20
        .Top = 200
        .Width = 360
        .Height = 40
        .Caption = "Note: Gemini API has usage quotas. Use batch processing for large datasets to avoid quota limits."
        .Name = "lblQuota"
    End With
    
    ' Add code to the buttons
    Dim buttonCode As String
    buttonCode = "Private Sub btnStandard_Click()" & vbCrLf & _
                 "    Unload Me" & vbCrLf & _
                 "    Call AnalyzeContractClauses" & vbCrLf & _
                 "End Sub" & vbCrLf & _
                 "" & vbCrLf & _
                 "Private Sub btnBatch_Click()" & vbCrLf & _
                 "    Unload Me" & vbCrLf & _
                 "    Call AnalyzeContractClausesInBatches" & vbCrLf & _
                 "End Sub" & vbCrLf & _
                 "" & vbCrLf & _
                 "Private Sub btnTest_Click()" & vbCrLf & _
                 "    Unload Me" & vbCrLf & _
                 "    Call TestApiConnection" & vbCrLf & _
                 "End Sub" & vbCrLf & _
                 "" & vbCrLf & _
                 "Private Sub btnClose_Click()" & vbCrLf & _
                 "    Unload Me" & vbCrLf & _
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
    Dim result4 As String
    Dim tag1 As String
    Dim tag2 As String
    Dim tag3 As String
    Dim tag4 As String
    
    ' Sample results with the new format
    result1 = "**UNCAPPED LIABILITY FOUND (Primary):** The agreement lacks a clear monetary cap in Section 12.2, which only limits liability for certain specific categories but leaves other potential liability uncapped."
    result2 = "**UNCAPPED LIABILITY FOUND:** While Section 8.3 provides a general cap of $5M, the exceptions in Section 8.4 are unusually broad, effectively swallowing the cap for most likely liability scenarios."
    result3 = "**CAPPED LIABILITY:** Section 10.1 clearly states that 'Total liability under this Agreement shall not exceed fees paid in the preceding 12 months.' Standard exceptions in 10.2 do not undermine this overarching cap."
    result4 = "**UNCERTAIN:** The contract contains contradictory liability provisions in Sections 9.2 and 15.4, making it impossible to determine if an overall cap applies."
    
    ' Test extraction for result1
    If InStr(1, result1, "UNCAPPED LIABILITY FOUND (Primary):") > 0 Then
        tag1 = "UNCAPPED LIABILITY FOUND (Primary)"
    ElseIf InStr(1, result1, "UNCAPPED LIABILITY FOUND:") > 0 Then
        tag1 = "UNCAPPED LIABILITY FOUND"
    ElseIf InStr(1, result1, "CAPPED LIABILITY:") > 0 Then
        tag1 = "CAPPED LIABILITY"
    ElseIf InStr(1, result1, "UNCERTAIN:") > 0 Then
        tag1 = "UNCERTAIN"
    Else
        ' Look for the response formats with asterisks
        If InStr(1, result1, "**UNCAPPED LIABILITY FOUND (Primary)**") > 0 Then
            tag1 = "UNCAPPED LIABILITY FOUND (Primary)"
        ElseIf InStr(1, result1, "**UNCAPPED LIABILITY FOUND**") > 0 Then
            tag1 = "UNCAPPED LIABILITY FOUND"
        ElseIf InStr(1, result1, "**CAPPED LIABILITY**") > 0 Then
            tag1 = "CAPPED LIABILITY"
        ElseIf InStr(1, result1, "**UNCERTAIN**") > 0 Then
            tag1 = "UNCERTAIN"
        Else
            tag1 = "UNRECOGNIZED FORMAT"
        End If
    End If
    
    ' Test extraction for result2
    If InStr(1, result2, "UNCAPPED LIABILITY FOUND (Primary):") > 0 Then
        tag2 = "UNCAPPED LIABILITY FOUND (Primary)"
    ElseIf InStr(1, result2, "UNCAPPED LIABILITY FOUND:") > 0 Then
        tag2 = "UNCAPPED LIABILITY FOUND"
    ElseIf InStr(1, result2, "CAPPED LIABILITY:") > 0 Then
        tag2 = "CAPPED LIABILITY"
    ElseIf InStr(1, result2, "UNCERTAIN:") > 0 Then
        tag2 = "UNCERTAIN"
    Else
        ' Look for the response formats with asterisks
        If InStr(1, result2, "**UNCAPPED LIABILITY FOUND (Primary)**") > 0 Then
            tag2 = "UNCAPPED LIABILITY FOUND (Primary)"
        ElseIf InStr(1, result2, "**UNCAPPED LIABILITY FOUND**") > 0 Then
            tag2 = "UNCAPPED LIABILITY FOUND"
        ElseIf InStr(1, result2, "**CAPPED LIABILITY**") > 0 Then
            tag2 = "CAPPED LIABILITY"
        ElseIf InStr(1, result2, "**UNCERTAIN**") > 0 Then
            tag2 = "UNCERTAIN"
        Else
            tag2 = "UNRECOGNIZED FORMAT"
        End If
    End If
    
    ' Test extraction for result3
    If InStr(1, result3, "UNCAPPED LIABILITY FOUND (Primary):") > 0 Then
        tag3 = "UNCAPPED LIABILITY FOUND (Primary)"
    ElseIf InStr(1, result3, "UNCAPPED LIABILITY FOUND:") > 0 Then
        tag3 = "UNCAPPED LIABILITY FOUND"
    ElseIf InStr(1, result3, "CAPPED LIABILITY:") > 0 Then
        tag3 = "CAPPED LIABILITY"
    ElseIf InStr(1, result3, "UNCERTAIN:") > 0 Then
        tag3 = "UNCERTAIN"
    Else
        ' Look for the response formats with asterisks
        If InStr(1, result3, "**UNCAPPED LIABILITY FOUND (Primary)**") > 0 Then
            tag3 = "UNCAPPED LIABILITY FOUND (Primary)"
        ElseIf InStr(1, result3, "**UNCAPPED LIABILITY FOUND**") > 0 Then
            tag3 = "UNCAPPED LIABILITY FOUND"
        ElseIf InStr(1, result3, "**CAPPED LIABILITY**") > 0 Then
            tag3 = "CAPPED LIABILITY"
        ElseIf InStr(1, result3, "**UNCERTAIN**") > 0 Then
            tag3 = "UNCERTAIN"
        Else
            tag3 = "UNRECOGNIZED FORMAT"
        End If
    End If
    
    ' Test extraction for result4
    If InStr(1, result4, "UNCAPPED LIABILITY FOUND (Primary):") > 0 Then
        tag4 = "UNCAPPED LIABILITY FOUND (Primary)"
    ElseIf InStr(1, result4, "UNCAPPED LIABILITY FOUND:") > 0 Then
        tag4 = "UNCAPPED LIABILITY FOUND"
    ElseIf InStr(1, result4, "CAPPED LIABILITY:") > 0 Then
        tag4 = "CAPPED LIABILITY"
    ElseIf InStr(1, result4, "UNCERTAIN:") > 0 Then
        tag4 = "UNCERTAIN"
    Else
        ' Look for the response formats with asterisks
        If InStr(1, result4, "**UNCAPPED LIABILITY FOUND (Primary)**") > 0 Then
            tag4 = "UNCAPPED LIABILITY FOUND (Primary)"
        ElseIf InStr(1, result4, "**UNCAPPED LIABILITY FOUND**") > 0 Then
            tag4 = "UNCAPPED LIABILITY FOUND"
        ElseIf InStr(1, result4, "**CAPPED LIABILITY**") > 0 Then
            tag4 = "CAPPED LIABILITY"
        ElseIf InStr(1, result4, "**UNCERTAIN**") > 0 Then
            tag4 = "UNCERTAIN"
        Else
            tag4 = "UNRECOGNIZED FORMAT"
        End If
    End If
    
    ' Display results
    Debug.Print "Test 1: " & Left(result1, 50) & "... -> Determination: " & tag1
    Debug.Print "Test 2: " & Left(result2, 50) & "... -> Determination: " & tag2
    Debug.Print "Test 3: " & Left(result3, 50) & "... -> Determination: " & tag3
    Debug.Print "Test 4: " & Left(result4, 50) & "... -> Determination: " & tag4
    
    ' Show results in a message box
    MsgBox "Test 1: " & Left(result1, 50) & "..." & vbCrLf & "Determination: " & tag1 & vbCrLf & vbCrLf & _
           "Test 2: " & Left(result2, 50) & "..." & vbCrLf & "Determination: " & tag2 & vbCrLf & vbCrLf & _
           "Test 3: " & Left(result3, 50) & "..." & vbCrLf & "Determination: " & tag3 & vbCrLf & vbCrLf & _
           "Test 4: " & Left(result4, 50) & "..." & vbCrLf & "Determination: " & tag4, _
           vbInformation, "Determination Extraction Test Results"
End Sub 