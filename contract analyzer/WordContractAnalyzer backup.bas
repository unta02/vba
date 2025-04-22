Option Explicit

' API key for Google Gemini API
Private Const API_KEY As String = "AIzaSyCWC8fiA6mUX590yF4tSchHxLA7iQBe3BY"
Private Const API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

' Word constants in case they're not recognized
#If VBA7 Then
    ' Word constants for highlighting and find operations
    Private Const wdColorTurquoise As Long = 3
    Private Const wdColorRed As Long = 6
    Private Const wdColorGreen As Long = 4
    Private Const wdAuto As Long = 0
    Private Const wdFindStop As Long = 0
    Private Const wdColorYellow As Long = 5
    Private Const wdColorGray50 As Long = 17
    Private Const wdColorPink As Long = 14
    Private Const wdColorBlue As Long = 5
    Private Const wdColorTeal As Long = 11
    Private Const wdColorViolet As Long = 13
#Else
    ' Word constants for highlighting and find operations
    Private Const wdColorTurquoise As Long = 3
    Private Const wdColorRed As Long = 6
    Private Const wdColorGreen As Long = 4
    Private Const wdAuto As Long = 0
    Private Const wdFindStop As Long = 0
    Private Const wdColorYellow As Long = 5
    Private Const wdColorGray50 As Long = 17
    Private Const wdColorPink As Long = 14
    Private Const wdColorBlue As Long = 5
    Private Const wdColorTeal As Long = 11
    Private Const wdColorViolet As Long = 13
#End If

' References needed:
' Microsoft Scripting Runtime
' Microsoft XML, v6.0 (or available version)
' Microsoft Word Object Library

' Comment color constants
Private Const PAYMENT_TERMS_COLOR As Long = wdColorTurquoise
Private Const RATE_CARDS_COLOR As Long = wdColorYellow
Private Const TRAVEL_EXPENSE_COLOR As Long = wdColorGray50
Private Const DIVERSE_SUPPLIER_COLOR As Long = wdColorPink
Private Const TERMINATION_COLOR As Long = wdColorGreen
Private Const LIABILITY_COLOR As Long = wdColorRed
Private Const DATA_PRIVACY_COLOR As Long = wdColorBlue
Private Const INSURANCE_COLOR As Long = wdColorTeal
Private Const BACKGROUND_CHECK_COLOR As Long = wdColorViolet
' Progress form handling functions
Private frmProgress As Object
' Main procedure to analyze Word document
Public Sub AnalyzeWordDocument()
    ' Declare variables
    Dim doc As Word.Document
    Dim prompt As String
    Dim result As String
    Dim textToAnalyze As String
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Get the active document
    Set doc = Word.Application.ActiveDocument
    
    ' Check if a document is open
    If doc Is Nothing Then
        MsgBox "Please open a document to analyze.", vbExclamation
        Exit Sub
    End If
    
    ' Display status in the status bar instead of a progress form
    Dim originalStatusBar As Boolean
    originalStatusBar = Word.Application.DisplayStatusBar
    Word.Application.DisplayStatusBar = True
    
    ' Update status and get document content
    Word.Application.StatusBar = "Extracting document content..."
    textToAnalyze = doc.Content.text
    
    ' Create the analysis prompt
    Word.Application.StatusBar = "Preparing analysis prompt..."
    prompt = GetAnalysisPrompt(textToAnalyze)
    
    ' Call Gemini API to analyze the text
    Word.Application.StatusBar = "Sending document to Gemini for analysis..."
    result = AnalyzeWithGemini(prompt)
    
    ' Process the results and highlight relevant clauses
    Word.Application.StatusBar = "Processing analysis results..."
    ProcessAnalysisResults doc, result
    
    ' Reset status bar
    Word.Application.StatusBar = False
    Word.Application.DisplayStatusBar = originalStatusBar
    
    ' Build completion message with all categories
    Dim completionMsg As String
    completionMsg = "Document analysis complete!" & vbCrLf & vbCrLf & _
                   "Contract clauses have been highlighted as follows:" & vbCrLf & _
                   "• Payment Terms (turquoise)" & vbCrLf & _
                   "• Rate Cards (yellow)" & vbCrLf & _
                   "• Travel & Expense Policy (gray)" & vbCrLf & _
                   "• Diverse Supplier Provisions (pink)" & vbCrLf & _
                   "• Termination Clauses (green)" & vbCrLf & _
                   "• Limitation of Liability (red)" & vbCrLf & _
                   "• Data Privacy (blue)" & vbCrLf & _
                   "• Insurance Provisions (teal)" & vbCrLf & _
                   "• Background Check/Drug Screening (violet)" & vbCrLf & vbCrLf & _
                   "Hover over highlighted text to view analysis comments."
    
    ' Confirm completion
    MsgBox completionMsg, vbInformation, "Contract Analysis Complete"
    
    Exit Sub
    
ErrorHandler:
    ' Reset status bar
    On Error Resume Next
    Word.Application.StatusBar = False
    Word.Application.DisplayStatusBar = originalStatusBar
    On Error GoTo 0
    
    ' Show error message
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Function to analyze text using Gemini API
Public Function AnalyzeWithGemini(textToAnalyze As String) As String
    Dim http As Object
    Dim jsonResponse As String
    Dim requestBody As String
    Dim fullUrl As String
    Dim responseText As String
    
    ' Set up HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    fullUrl = API_URL & "?key=" & API_KEY
    
    ' Prepare the request body (JSON)
    requestBody = "{" & _
                 """contents"": [" & _
                 "{""parts"":[{""text"": """ & JsonEscape(textToAnalyze) & """}]}" & _
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
        AnalyzeWithGemini = "Error in HTTP request: " & Err.Description
        Exit Function
    End If
    
    ' Process the response
    responseText = http.responseText
    
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

' Create the detailed analysis prompt for Gemini
Public Function GetAnalysisPrompt(documentText As String) As String
    Dim prompt As String
    Dim part1 As String
    Dim part2 As String
    Dim part3 As String
    Dim part4 As String
    Dim part5 As String
    Dim part6 As String
    Dim part7 As String
    
    ' Break the prompt into smaller parts to avoid "Too many line continuations" error
    part1 = "You are a contract analysis expert. I will provide you with a contract document that you need to analyze." & vbCrLf & vbCrLf & _
            "Analyze the provided contract document and extract the following information, identifying the exact text for relevant clauses, keywords, and details:" & vbCrLf & vbCrLf
    
    part2 = "1. Payment Terms" & vbCrLf & _
            "Description: Identify payment terms, including standard periods (e.g., 30 days), conditions for exceeding terms, and invoice requirements. Highlight any deviations from the standard policy." & vbCrLf & _
            "Keywords: Pay, Payable, Invoice, Net." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Payment due within [X] days of receipt of a correct invoice." & vbCrLf & _
            "b.) Any deviations require [specific approval or conditions]." & vbCrLf & vbCrLf
    
    part3 = "2. Rate Cards" & vbCrLf & _
            "Description: Extract details about rate cards, hourly rate structures, and personnel-based fees. Reference any rate tables." & vbCrLf & _
            "Keywords: Rate, Rate Cards, Hourly." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Fee structure for hourly rates based on personnel levels, types of work, and geographical regions." & vbCrLf & _
            "b.) Provide examples where applicable, including rate tables or caps." & vbCrLf & vbCrLf & _
            "3. Client Travel and Expense Policy" & vbCrLf & _
            "Description: Identify clauses related to travel and expense reimbursement. Note if the client's travel policy applies and highlight any restrictions or exceptions." & vbCrLf & _
            "Keywords: Expense, Travel, Expense Policy, Travel Guide." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Reimbursement terms for travel-related expenses, including approvals and documentation requirements." & vbCrLf & _
            "b.) Policies about economy fares, preferred accommodations, and non-reimbursable items." & vbCrLf & vbCrLf
    
    part4 = "4. Diverse Supplier Provisions" & vbCrLf & _
            "Description: Highlight provisions encouraging the use of diverse suppliers and related reporting requirements." & vbCrLf & _
            "Keywords: Diversity, Diverse Supplier, Inclusion." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Include language supporting diverse supplier inclusion and any reporting requirements (e.g., quarterly reports)." & vbCrLf & vbCrLf & _
            "5. Termination Clauses" & vbCrLf & _
            "Description: Identify terms for termination, including 'termination for convenience' or 'material breach.' Highlight any fees or specific conditions." & vbCrLf & _
            "Keywords: Termination for convenience, Termination without cause, Material breach." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Right to terminate with [X] days' notice or under specific conditions." & vbCrLf & _
            "b.) Associated termination fees, if any, and their calculation." & vbCrLf & vbCrLf
    
    part5 = "6. Limitation of Liability" & vbCrLf & _
            "Description: Extract limitations of liability clauses, including standard caps or carveouts." & vbCrLf & _
            "Keywords: Indirect damage exclusion, Mutual limits, Super cap." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Liability capped at [X] or a multiple of fees paid within a specific timeframe." & vbCrLf & _
            "b.) Uncapped liability for specific scenarios, such as gross negligence." & vbCrLf & vbCrLf & _
            "7. Data Privacy" & vbCrLf & _
            "Description: Summarize data privacy obligations, including compliance with regulations like GDPR or specific client data protection requirements." & vbCrLf & _
            "Keywords: Data Privacy, Data Processing, GDPR." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Outline responsibilities for protecting personal data and compliance with regulations." & vbCrLf & _
            "b.) Highlight any cross-border data restrictions or notification obligations in case of breaches." & vbCrLf & vbCrLf
    
    part6 = "8. Insurance Provisions" & vbCrLf & _
            "Description: Highlight insurance coverage requirements, including limits and types of coverage (e.g., liability, workers' compensation)." & vbCrLf & _
            "Keywords: Insurance, Coverage." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Required insurance types and coverage limits." & vbCrLf & _
            "b.) Period of coverage and renewal obligations." & vbCrLf & vbCrLf & _
            "9. Background Check/Drug Screening" & vbCrLf & _
            "Description: Extract clauses requiring background checks or drug testing. Note any restrictions or client-specific requirements." & vbCrLf & _
            "Keywords: Background check, Drug, Alcohol testing." & vbCrLf & _
            "Sample Format:" & vbCrLf & _
            "a.) Requirements for background checks or drug screening, including specific roles or access levels." & vbCrLf & vbCrLf
    
    part7 = "FORMAT YOUR RESPONSE EXACTLY LIKE THIS:" & vbCrLf & _
            "PAYMENT TERMS:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "RATE CARDS:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "CLIENT TRAVEL AND EXPENSE POLICY:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "DIVERSE SUPPLIER PROVISIONS:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "TERMINATION CLAUSES:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "LIMITATION OF LIABILITY:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "DATA PRIVACY:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "INSURANCE PROVISIONS:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "BACKGROUND CHECK/DRUG SCREENING:" & vbCrLf & _
            "[EXACT TEXT FROM DOCUMENT]" & vbCrLf & _
            "Analysis: [brief analysis]" & vbCrLf & vbCrLf & _
            "Here is the contract document to analyze:" & vbCrLf & vbCrLf & _
            documentText
    
    ' Combine all parts
    prompt = part1 & part2 & part3 & part4 & part5 & part6 & part7
             
    GetAnalysisPrompt = prompt
End Function

' Process the analysis results and highlight relevant clauses
Private Sub ProcessAnalysisResults(doc As Word.Document, analysisResult As String)
    ' Variables for sections
    Dim paymentSection As String
    Dim rateCardsSection As String
    Dim travelExpenseSection As String
    Dim diverseSupplierSection As String
    Dim terminationSection As String
    Dim liabilitySection As String
    Dim dataPrivacySection As String
    Dim insuranceSection As String
    Dim backgroundCheckSection As String
    
    ' Extract sections using markers
    paymentSection = ExtractSection(analysisResult, "PAYMENT TERMS:", "RATE CARDS:")
    rateCardsSection = ExtractSection(analysisResult, "RATE CARDS:", "CLIENT TRAVEL AND EXPENSE POLICY:")
    travelExpenseSection = ExtractSection(analysisResult, "CLIENT TRAVEL AND EXPENSE POLICY:", "DIVERSE SUPPLIER PROVISIONS:")
    diverseSupplierSection = ExtractSection(analysisResult, "DIVERSE SUPPLIER PROVISIONS:", "TERMINATION CLAUSES:")
    terminationSection = ExtractSection(analysisResult, "TERMINATION CLAUSES:", "LIMITATION OF LIABILITY:")
    liabilitySection = ExtractSection(analysisResult, "LIMITATION OF LIABILITY:", "DATA PRIVACY:")
    dataPrivacySection = ExtractSection(analysisResult, "DATA PRIVACY:", "INSURANCE PROVISIONS:")
    insuranceSection = ExtractSection(analysisResult, "INSURANCE PROVISIONS:", "BACKGROUND CHECK/DRUG SCREENING:")
    backgroundCheckSection = ExtractSection(analysisResult, "BACKGROUND CHECK/DRUG SCREENING:", Chr(0))
    
    ' Apply highlights and comments to document
    HighlightAndComment doc, paymentSection, "Payment Terms", PAYMENT_TERMS_COLOR
    HighlightAndComment doc, rateCardsSection, "Rate Cards", RATE_CARDS_COLOR
    HighlightAndComment doc, travelExpenseSection, "Travel & Expense Policy", TRAVEL_EXPENSE_COLOR
    HighlightAndComment doc, diverseSupplierSection, "Diverse Supplier Provisions", DIVERSE_SUPPLIER_COLOR
    HighlightAndComment doc, terminationSection, "Termination Clause", TERMINATION_COLOR
    HighlightAndComment doc, liabilitySection, "Limitation of Liability", LIABILITY_COLOR
    HighlightAndComment doc, dataPrivacySection, "Data Privacy", DATA_PRIVACY_COLOR
    HighlightAndComment doc, insuranceSection, "Insurance Provisions", INSURANCE_COLOR
    HighlightAndComment doc, backgroundCheckSection, "Background Check/Drug Screening", BACKGROUND_CHECK_COLOR
End Sub

' Extract a section from the analysis result
Private Function ExtractSection(text As String, startMarker As String, endMarker As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim section As String
    
    ' Find the section start
    startPos = InStr(text, startMarker)
    If startPos = 0 Then
        ExtractSection = ""
        Exit Function
    End If
    
    ' Adjust startPos to the beginning of the content after the marker
    startPos = startPos + Len(startMarker)
    
    ' Find the section end
    If endMarker = Chr(0) Then
        ' If endMarker is null character, take everything to the end
        endPos = Len(text) + 1
    Else
        endPos = InStr(startPos, text, endMarker)
        If endPos = 0 Then endPos = Len(text) + 1
    End If
    
    ' Extract the section
    section = Mid(text, startPos, endPos - startPos)
    
    ' Trim whitespace
    ExtractSection = Trim(section)
End Function

' Highlight text and add comments to document
Private Sub HighlightAndComment(doc As Word.Document, section As String, commentTitle As String, highlightColor As Long)
    Dim exactText As String
    Dim analysis As String
    Dim rng As Word.Range
    Dim textLines As Variant
    Dim i As Long
    
    ' Check if section is empty
    If Len(Trim(section)) = 0 Then Exit Sub
    
    ' Extract the exact text and analysis
    exactText = ExtractExactText(section)
    analysis = ExtractAnalysis(section)
    
    ' If exact text is found, highlight and comment
    If Len(Trim(exactText)) > 0 Then
        ' Split into lines in case there are multiple clauses
        textLines = Split(exactText, vbCrLf)
        
        ' Process each line
        For i = LBound(textLines) To UBound(textLines)
            If Len(Trim(textLines(i))) > 0 Then
                ' Update status bar
                Word.Application.StatusBar = "Highlighting " & commentTitle & "..."
                
                ' Find this text in the document
                Set rng = FindTextInDocument(doc, Trim(textLines(i)))
                
                ' If found, highlight and add comment
                If Not rng Is Nothing Then
                    ' Highlight text using a more reliable method
                    rng.Font.ColorIndex = wdAuto  ' Reset color
                    ' Apply highlight
                    With rng
                        .FormattedText.HighlightColorIndex = highlightColor
                    End With
                    
                    ' Add comment
                    Dim commentText As String
                    commentText = commentTitle & ": " & analysis
                    doc.Comments.Add rng, commentText
                End If
            End If
        Next i
    End If
End Sub

' Extract the exact text from a section
Private Function ExtractExactText(section As String) As String
    Dim analysisPos As Long
    Dim text As String
    
    ' Extract text before "Analysis:" marker
    analysisPos = InStr(1, section, "Analysis:", vbTextCompare)
    
    If analysisPos > 0 Then
        text = Left(section, analysisPos - 1)
    Else
        text = section
    End If
    
    ' Return cleaned text
    ExtractExactText = Trim(text)
End Function

' Extract the analysis from a section
Private Function ExtractAnalysis(section As String) As String
    Dim analysisPos As Long
    Dim analysis As String
    
    ' Find "Analysis:" marker and extract text after it
    analysisPos = InStr(1, section, "Analysis:", vbTextCompare)
    
    If analysisPos > 0 Then
        analysis = Mid(section, analysisPos + 9) ' 9 = len("Analysis:")
    Else
        analysis = ""
    End If
    
    ' Return cleaned analysis
    ExtractAnalysis = Trim(analysis)
End Function

' Find text in a document
Private Function FindTextInDocument(doc As Word.Document, textToFind As String) As Word.Range
    Dim rng As Word.Range
    Dim found As Boolean
    
    ' Create range for the document
    Set rng = doc.Content
    
    ' Clean up search text
    textToFind = Trim(textToFind)
    If Len(textToFind) > 255 Then
        ' Word's find has a limitation on search string length
        textToFind = Left(textToFind, 255)
    End If
    
    ' Execute search without relying on Application.Options
    On Error Resume Next
    
    With rng.Find
        .ClearFormatting
        .Text = textToFind
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        found = .Execute
    End With
    
    On Error GoTo 0
    
    ' Return results
    If found Then
        Set FindTextInDocument = rng
    Else
        Set FindTextInDocument = Nothing
    End If
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
    ' This simplified extractor works directly with string operations
    ' rather than attempting to parse the full JSON structure
    
    On Error Resume Next
    
    Dim result As String
    result = ""
    
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

Private Sub ShowProgressForm(message As String)
    ' Create a simple progress form
    Set frmProgress = Application.VBE.ActiveVBProject.VBComponents.Add(3) ' 3 = UserForm
    
    ' Set form properties
    With frmProgress
        .Properties("Caption") = "Analyzing Document"
        .Properties("Width") = 300
        .Properties("Height") = 100
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' Add a label
    Dim lblStatus As Object
    Set lblStatus = frmProgress.Designer.Controls.Add("Forms.Label.1")
    With lblStatus
        .Left = 20
        .Top = 20
        .Width = 260
        .Height = 40
        .Caption = message
        .Name = "lblStatus"
    End With
    
    ' Add code for the form
'    Dim formCode As String
'    formCode = "Public Sub UpdateStatus(message As String)" & vbCrLf & _
'               "    Me.lblStatus.Caption = message" & vbCrLf & _
'               "    DoEvents" & vbCrLf & _
'               "End Sub"
'
'    frmProgress.CodeModule.AddFromString formCode
    
    ' Show the form modeless
    frmProgress.Name = "frmProgress"
    
    ' Use DoEvents to allow the UI to update
    DoEvents
    
    ' Create a procedure to show the form
    Dim showFormCode As String
    showFormCode = "Sub ShowProgressForm()" & vbCrLf & _
                   "    frmProgress.Show vbModeless" & vbCrLf & _
                   "End Sub"
                   
    Application.VBE.ActiveVBProject.VBComponents.Add(1).CodeModule.AddFromString showFormCode
    
    ' Call the procedure to show the form
    Application.Run "ShowProgressForm"
End Sub

Private Sub UpdateProgressStatus(message As String)
    On Error Resume Next
    
    If Not frmProgress Is Nothing Then
        ' Update the status message
        Application.Run "frmProgress.UpdateStatus", message
        DoEvents
    End If
    
    On Error GoTo 0
End Sub

Private Sub CloseProgressForm()
    On Error Resume Next
    
    If Not frmProgress Is Nothing Then
        ' Unload the form
        VBA.UserForms.Unload frmProgress
        
        ' Remove the component
        Application.VBE.ActiveVBProject.VBComponents.Remove frmProgress
        
        Set frmProgress = Nothing
    End If
    
    On Error GoTo 0
End Sub

' Create a ribbon button (code sample that would need to be implemented in a Ribbon XML file)
' This is just for reference and won't actually run
Public Sub OnRibbonButtonClick(control As IRibbonControl)
    AnalyzeWordDocument
End Sub

' Add a menu item to Word
Public Sub AddMenuItemToWord()
    On Error Resume Next
    
    ' Set a reference to the CommandBars collection
    Dim cb As CommandBar
    Dim cbc As CommandBarControl
    
    ' Try to find existing menu item first and remove it to avoid duplicates
    On Error Resume Next
    Word.Application.CommandBars("Menu Bar").Controls("Contract Analyzer").Delete
    On Error GoTo 0
    
    ' Add to the Tools menu
    Set cb = Word.Application.CommandBars("Menu Bar")
    Set cbc = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    
    With cbc
        .Caption = "Contract Analyzer"
        .Style = msoButtonCaption
        .OnAction = "AnalyzeWordDocument"
        .BeginGroup = True
    End With
    
    MsgBox "Contract Analyzer added to the Menu Bar." & vbCrLf & _
           "You can now analyze your contracts by clicking on it.", vbInformation
           
    On Error GoTo 0
End Sub

