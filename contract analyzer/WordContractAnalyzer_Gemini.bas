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

' Define the HighlightInfo type at the top of the module
Private Type HighlightInfo
    text As String
    PageNumber As Long
    analysis As String
End Type

' Add these declarations near the top of the module with other constants
Private Const SUMMARY_TITLE As String = "Contract Analysis Summary"
Private Const SUMMARY_FILENAME As String = "ContractSummary.html"

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
    ' Split the prompt building into smaller functions
    Dim prompt As String
    
    ' Get introduction and category descriptions in smaller chunks
    Dim introduction As String
    introduction = GetPromptIntroduction()
    
    Dim categories1 As String
    categories1 = GetPromptCategories1to3()
    
    Dim categories2 As String
    categories2 = GetPromptCategories4to6()
    
    Dim categories3 As String
    categories3 = GetPromptCategories7to9()
    
    ' Get response format in smaller chunks
    Dim format1 As String
    format1 = GetPromptFormat1()
    
    Dim format2 As String
    format2 = GetPromptFormat2()
    
    Dim format3 As String
    format3 = GetPromptFormat3()
    
    Dim format4 As String
    format4 = GetPromptFormat4(documentText)
    
    ' Combine all parts with minimal line continuations
    prompt = introduction
    prompt = prompt & categories1
    prompt = prompt & categories2
    prompt = prompt & categories3
    prompt = prompt & format1
    prompt = prompt & format2
    prompt = prompt & format3
    prompt = prompt & format4
             
    GetAnalysisPrompt = prompt
End Function

' Helper functions to build different parts of the prompt without too many line continuations

' Part 1: Introduction
Private Function GetPromptIntroduction() As String
    Dim text As String
    
    text = "You are a contract analysis expert. I will provide you with a contract document that you need to analyze." & vbCrLf & vbCrLf
    text = text & "Analyze the provided contract document and extract the following information, identifying the exact text for relevant clauses, keywords, and details:" & vbCrLf & vbCrLf
    
    GetPromptIntroduction = text
End Function

' Part 2: Categories 1-3
Private Function GetPromptCategories1to3() As String
    Dim text As String
    
    ' Category 1: Payment Terms
    text = "1. Payment Terms" & vbCrLf
    text = text & "Description: Identify payment terms, including standard periods (e.g., 30 days), conditions for exceeding terms, and invoice requirements. Highlight any deviations from the standard policy." & vbCrLf
    text = text & "Keywords: Pay, Payable, Invoice, Net." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Payment due within [X] days of receipt of a correct invoice." & vbCrLf
    text = text & "b.) Any deviations require [specific approval or conditions]." & vbCrLf & vbCrLf
    
    ' Category 2: Rate Cards
    text = text & "2. Rate Cards" & vbCrLf
    text = text & "Description: Extract details about rate cards, hourly rate structures, and personnel-based fees. Reference any rate tables." & vbCrLf
    text = text & "Keywords: Rate, Rate Cards, Hourly." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Fee structure for hourly rates based on personnel levels, types of work, and geographical regions." & vbCrLf
    text = text & "b.) Provide examples where applicable, including rate tables or caps." & vbCrLf & vbCrLf
    
    ' Category 3: Travel and Expense
    text = text & "3. Client Travel and Expense Policy" & vbCrLf
    text = text & "Description: Identify clauses related to travel and expense reimbursement. Note if the client's travel policy applies and highlight any restrictions or exceptions." & vbCrLf
    text = text & "Keywords: Expense, Travel, Expense Policy, Travel Guide." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Reimbursement terms for travel-related expenses, including approvals and documentation requirements." & vbCrLf
    text = text & "b.) Policies about economy fares, preferred accommodations, and non-reimbursable items." & vbCrLf & vbCrLf
    
    GetPromptCategories1to3 = text
End Function

' Part 3: Categories 4-6
Private Function GetPromptCategories4to6() As String
    Dim text As String
    
    ' Category 4: Diverse Supplier
    text = "4. Diverse Supplier Provisions" & vbCrLf
    text = text & "Description: Highlight provisions encouraging the use of diverse suppliers and related reporting requirements." & vbCrLf
    text = text & "Keywords: Diversity, Diverse Supplier, Inclusion." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Include language supporting diverse supplier inclusion and any reporting requirements (e.g., quarterly reports)." & vbCrLf & vbCrLf
    
    ' Category 5: Termination
    text = text & "5. Termination Clauses" & vbCrLf
    text = text & "Description: Identify terms for termination, including 'termination for convenience' or 'material breach.' Highlight any fees or specific conditions." & vbCrLf
    text = text & "Keywords: Termination for convenience, Termination without cause, Material breach." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Right to terminate with [X] days' notice or under specific conditions." & vbCrLf
    text = text & "b.) Associated termination fees, if any, and their calculation." & vbCrLf & vbCrLf
    
    ' Category 6: Liability
    text = text & "6. Limitation of Liability" & vbCrLf
    text = text & "Description: Extract limitations of liability clauses, including standard caps or carveouts." & vbCrLf
    text = text & "Keywords: Indirect damage exclusion, Mutual limits, Super cap." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Liability capped at [X] or a multiple of fees paid within a specific timeframe." & vbCrLf
    text = text & "b.) Uncapped liability for specific scenarios, such as gross negligence." & vbCrLf & vbCrLf
    
    GetPromptCategories4to6 = text
End Function

' Part 4: Categories 7-9
Private Function GetPromptCategories7to9() As String
    Dim text As String
    
    ' Category 7: Data Privacy
    text = "7. Data Privacy" & vbCrLf
    text = text & "Description: Summarize data privacy obligations, including compliance with regulations like GDPR or specific client data protection requirements." & vbCrLf
    text = text & "Keywords: Data Privacy, Data Processing, GDPR." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Outline responsibilities for protecting personal data and compliance with regulations." & vbCrLf
    text = text & "b.) Highlight any cross-border data restrictions or notification obligations in case of breaches." & vbCrLf & vbCrLf
    
    ' Category 8: Insurance
    text = text & "8. Insurance Provisions" & vbCrLf
    text = text & "Description: Highlight insurance coverage requirements, including limits and types of coverage (e.g., liability, workers' compensation)." & vbCrLf
    text = text & "Keywords: Insurance, Coverage." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Required insurance types and coverage limits." & vbCrLf
    text = text & "b.) Period of coverage and renewal obligations." & vbCrLf & vbCrLf
    
    ' Category 9: Background Check
    text = text & "9. Background Check/Drug Screening" & vbCrLf
    text = text & "Description: Extract clauses requiring background checks or drug testing. Note any restrictions or client-specific requirements." & vbCrLf
    text = text & "Keywords: Background check, Drug, Alcohol testing." & vbCrLf
    text = text & "Sample Format:" & vbCrLf
    text = text & "a.) Requirements for background checks or drug screening, including specific roles or access levels." & vbCrLf & vbCrLf
    
    GetPromptCategories7to9 = text
End Function

' Part 5: Response Format Part 1
Private Function GetPromptFormat1() As String
    Dim text As String
    
    text = "FORMAT YOUR RESPONSE EXACTLY LIKE THIS:" & vbCrLf
    text = text & "PAYMENT TERMS:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    text = text & "RATE CARDS:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    GetPromptFormat1 = text
End Function

' Part 6: Response Format Part 2
Private Function GetPromptFormat2() As String
    Dim text As String
    
    text = "CLIENT TRAVEL AND EXPENSE POLICY:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    text = text & "DIVERSE SUPPLIER PROVISIONS:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    GetPromptFormat2 = text
End Function

' Part 7: Response Format Part 3
Private Function GetPromptFormat3() As String
    Dim text As String
    
    text = "TERMINATION CLAUSES:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    text = text & "LIMITATION OF LIABILITY:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    text = text & "DATA PRIVACY:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    GetPromptFormat3 = text
End Function

' Part 8: Response Format Part 4 with Document Text
Private Function GetPromptFormat4(documentText As String) As String
    Dim text As String
    
    text = "INSURANCE PROVISIONS:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    text = text & "BACKGROUND CHECK/DRUG SCREENING:" & vbCrLf
    text = text & "[EXACT TEXT FROM DOCUMENT]" & vbCrLf
    text = text & "Analysis: [brief analysis]" & vbCrLf & vbCrLf
    
    text = text & "Here is the contract document to analyze:" & vbCrLf & vbCrLf
    text = text & documentText
    
    GetPromptFormat4 = text
End Function

' Add this new function to display a summary of the analysis results in a side pane with enhanced details
Private Sub ShowAnalysisSummary(analysisResult As String, doc As Word.Document)
    ' Extract sections
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
    
    ' Get highlight information from document
    Dim paymentHighlights() As HighlightInfo
    Dim rateCardsHighlights() As HighlightInfo
    Dim travelExpenseHighlights() As HighlightInfo
    Dim diverseSupplierHighlights() As HighlightInfo
    Dim terminationHighlights() As HighlightInfo
    Dim liabilityHighlights() As HighlightInfo
    Dim dataPrivacyHighlights() As HighlightInfo
    Dim insuranceHighlights() As HighlightInfo
    Dim backgroundCheckHighlights() As HighlightInfo
    
    paymentHighlights = GetHighlightInfo(doc, paymentSection, PAYMENT_TERMS_COLOR)
    rateCardsHighlights = GetHighlightInfo(doc, rateCardsSection, RATE_CARDS_COLOR)
    travelExpenseHighlights = GetHighlightInfo(doc, travelExpenseSection, TRAVEL_EXPENSE_COLOR)
    diverseSupplierHighlights = GetHighlightInfo(doc, diverseSupplierSection, DIVERSE_SUPPLIER_COLOR)
    terminationHighlights = GetHighlightInfo(doc, terminationSection, TERMINATION_COLOR)
    liabilityHighlights = GetHighlightInfo(doc, liabilitySection, LIABILITY_COLOR)
    dataPrivacyHighlights = GetHighlightInfo(doc, dataPrivacySection, DATA_PRIVACY_COLOR)
    insuranceHighlights = GetHighlightInfo(doc, insuranceSection, INSURANCE_COLOR)
    backgroundCheckHighlights = GetHighlightInfo(doc, backgroundCheckSection, BACKGROUND_CHECK_COLOR)
    
    ' Create HTML for the summary with enhanced accordion and styling
    Dim htmlContent As String
    
    ' Get each part of the HTML content separately to avoid line continuations
    htmlContent = GetHtmlHeader()
    
    ' Add each section to the HTML content with enhanced accordion
    htmlContent = htmlContent & CreateAccordionSection("Payment Terms", paymentSection, "payment", paymentHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Rate Cards", rateCardsSection, "rate", rateCardsHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Travel & Expense Policy", travelExpenseSection, "travel", travelExpenseHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Diverse Supplier Provisions", diverseSupplierSection, "diverse", diverseSupplierHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Termination Clauses", terminationSection, "termination", terminationHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Limitation of Liability", liabilitySection, "liability", liabilityHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Data Privacy", dataPrivacySection, "privacy", dataPrivacyHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Insurance Provisions", insuranceSection, "insurance", insuranceHighlights)
    htmlContent = htmlContent & CreateAccordionSection("Background Check/Drug Screening", backgroundCheckSection, "background", backgroundCheckHighlights)
    
    ' Add JavaScript for accordion functionality
    htmlContent = htmlContent & GetAccordionJavaScript()
    
    ' Close the HTML
    htmlContent = htmlContent & "</body>" & vbCrLf & "</html>"
    
    ' Save HTML to temp file
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\" & SUMMARY_FILENAME
    
    ' Write HTML to file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open tempFile For Output As #fileNum
    Print #fileNum, htmlContent
    Close #fileNum
    
    ' Open the HTML file in Word's Document Explorer
    On Error Resume Next
    
    ' Check if Explorer object is available
    Dim explorerAvailable As Boolean
    explorerAvailable = True
    
    ' Try to use Document Explorer (Task Pane)
    If explorerAvailable Then
        ' Try to use Windows Explorer Web browser
        Dim IE As Object
        Set IE = CreateObject("InternetExplorer.Application")
        
        If Not IE Is Nothing Then
            ' Configure Internet Explorer window
            ConfigureIEWindow IE, tempFile
        Else
            ' Fallback - open in default browser
            Shell "explorer.exe " & tempFile, vbNormalFocus
        End If
    Else
        ' Fallback - open in default browser
        Shell "explorer.exe " & tempFile, vbNormalFocus
    End If
    
    On Error GoTo 0
End Sub

' Helper function to configure Internet Explorer window
Private Sub ConfigureIEWindow(IE As Object, url As String)
    With IE
        .Navigate url
        .Visible = True
        .Width = 400
        .Height = Application.Height
        .Left = Application.Left + Application.Width - 400
        .Top = Application.Top
        .MenuBar = False
        .ToolBar = False
        .StatusBar = False
        .Resizable = True
    End With
End Sub

' Helper function to get HTML header and styles
Private Function GetHtmlHeader() As String
    Dim html As String
    
    ' Start HTML document
    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html>" & vbCrLf
    html = html & "<head>" & vbCrLf
    html = html & "<style>" & vbCrLf
    
    ' Add CSS styles in chunks to avoid line continuations
    html = html & AddStyles1()
    html = html & AddStyles2()
    html = html & AddStyles3()
    
    html = html & "</style>" & vbCrLf
    html = html & "</head>" & vbCrLf
    html = html & "<body>" & vbCrLf
    html = html & "<h1>Contract Analysis Summary</h1>"
    
    GetHtmlHeader = html
End Function

' Part 1 of CSS styles
Private Function AddStyles1() As String
    Dim css As String
    
    css = "body { font-family: 'Segoe UI', Arial, sans-serif; margin: 10px; background-color: #f8f9fa; color: #333; }" & vbCrLf
    css = css & "h1 { font-size: 20px; color: #333; margin-bottom: 20px; text-align: center; border-bottom: 1px solid #ddd; padding-bottom: 10px; }" & vbCrLf
    css = css & ".accordion { margin-bottom: 10px; border-radius: 5px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }" & vbCrLf
    css = css & ".accordion-header { font-size: 14px; margin: 0; padding: 10px 15px; cursor: pointer; display: flex; justify-content: flex-start; align-items: center; }" & vbCrLf
    css = css & ".accordion-header:after { content: '+'; font-size: 18px; font-weight: bold; margin-left: auto; }" & vbCrLf
    css = css & ".active:after { content: '-'; }" & vbCrLf
    css = css & ".accordion-content { max-height: 0; overflow: hidden; transition: max-height 0.3s ease-out; background-color: white; }" & vbCrLf
    css = css & ".summary { padding: 15px; margin: 0; background-color: white; border-left: 5px solid #ddd; }" & vbCrLf
    css = css & ".highlight-list { margin: 0; padding: 0 15px 15px; list-style-type: none; }" & vbCrLf
    css = css & ".highlight-item { padding: 8px; margin-bottom: 5px; background-color: #f3f4f6; border-radius: 3px; border-left: 3px solid #ddd; }" & vbCrLf
    
    AddStyles1 = css
End Function

' Part 2 of CSS styles
Private Function AddStyles2() As String
    Dim css As String
    
    css = ".page-number { font-weight: bold; background-color: #eee; border-radius: 3px; padding: 2px 5px; display: inline-block; margin-right: 8px; }" & vbCrLf
    css = css & ".counter { display: inline-block; background-color: #eee; border-radius: 10px; padding: 2px 8px; font-size: 12px; margin-right: 10px; color: #000000; }" & vbCrLf
    css = css & ".payment-color { background-color: #30d5c8; color: white; }" & vbCrLf
    css = css & ".payment-border { border-color: #30d5c8; }" & vbCrLf
    css = css & ".rate-color { background-color: #ffff00; color: black; }" & vbCrLf
    css = css & ".rate-border { border-color: #ffff00; }" & vbCrLf
    css = css & ".travel-color { background-color: #808080; color: white; }" & vbCrLf
    css = css & ".travel-border { border-color: #808080; }" & vbCrLf
    css = css & ".diverse-color { background-color: #ffc0cb; color: black; }" & vbCrLf
    css = css & ".diverse-border { border-color: #ffc0cb; }" & vbCrLf
    
    AddStyles2 = css
End Function

' Part 3 of CSS styles
Private Function AddStyles3() As String
    Dim css As String
    
    css = ".termination-color { background-color: #00ff00; color: black; }" & vbCrLf
    css = css & ".termination-border { border-color: #00ff00; }" & vbCrLf
    css = css & ".liability-color { background-color: #ff0000; color: white; }" & vbCrLf
    css = css & ".liability-border { border-color: #ff0000; }" & vbCrLf
    css = css & ".privacy-color { background-color: #0000ff; color: white; }" & vbCrLf
    css = css & ".privacy-border { border-color: #0000ff; }" & vbCrLf
    css = css & ".insurance-color { background-color: #008080; color: white; }" & vbCrLf
    css = css & ".insurance-border { border-color: #008080; }" & vbCrLf
    css = css & ".background-color { background-color: #ee82ee; color: white; }" & vbCrLf
    css = css & ".background-border { border-color: #ee82ee; }" & vbCrLf
    
    AddStyles3 = css
End Function

' Helper function to get JavaScript for accordion functionality
Private Function GetAccordionJavaScript() As String
    Dim js As String
    
    js = "<script>" & vbCrLf
    js = js & "var acc = document.getElementsByClassName('accordion-header');" & vbCrLf
    js = js & "for (var i = 0; i < acc.length; i++) {" & vbCrLf
    js = js & "  acc[i].addEventListener('click', function() {" & vbCrLf
    js = js & "    this.classList.toggle('active');" & vbCrLf
    js = js & "    var content = this.nextElementSibling;" & vbCrLf
    js = js & "    if (content.style.maxHeight) {" & vbCrLf
    js = js & "      content.style.maxHeight = null;" & vbCrLf
    js = js & "    } else {" & vbCrLf
    js = js & "      content.style.maxHeight = content.scrollHeight + 'px';" & vbCrLf
    js = js & "    }" & vbCrLf
    js = js & "  });" & vbCrLf
    js = js & "}" & vbCrLf
    js = js & "</script>" & vbCrLf
    
    GetAccordionJavaScript = js
End Function

' Helper function to create accordion sections for HTML content
Private Function CreateAccordionSection(title As String, section As String, cssClass As String, highlights() As HighlightInfo) As String
    Dim analysis As String
    Dim html As String
    Dim count As Integer
    
    analysis = ExtractAnalysis(section)
    
    ' If no data found, show "None found"
    If Len(Trim(analysis)) = 0 Then
        analysis = "None found in document"
        count = 0
    Else
        ' Count highlights if we have them
        On Error Resume Next
        count = UBound(highlights) + 1
        If Err.Number <> 0 Then count = 0
        On Error GoTo 0
    End If
    
    ' Start accordion section - build in parts to avoid line continuations
    html = "<div class=""accordion"">" & vbCrLf
    html = html & "<div class=""accordion-header " & cssClass & "-color"">"
    html = html & "<span class=""counter"">" & count & "</span>" & title & "</div>" & vbCrLf
    html = html & "<div class=""accordion-content"">" & vbCrLf
    html = html & "<p class=""summary"">" & analysis & "</p>" & vbCrLf
    
    ' Add highlight list if there are any
    If count > 0 Then
        html = html & "<ul class=""highlight-list"">" & vbCrLf
        
        Dim i As Integer
        For i = 0 To UBound(highlights)
            ' Truncate text if too long
            Dim displayText As String
            If Len(highlights(i).text) > 100 Then
                displayText = Left(highlights(i).text, 97) & "..."
            Else
                displayText = highlights(i).text
            End If
            
            ' Add highlight item in parts
            html = html & "<li class=""highlight-item " & cssClass & "-border"">"
            html = html & "<span class=""page-number"">p." & highlights(i).PageNumber & "</span>"
            html = html & displayText
            html = html & "</li>" & vbCrLf
        Next i
        
        html = html & "</ul>" & vbCrLf
    End If
    
    ' Close accordion
    html = html & "</div>" & vbCrLf & "</div>" & vbCrLf
    
    CreateAccordionSection = html
End Function

' Modify ProcessAnalysisResults to pass the document to ShowAnalysisSummary
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
    
    ' Show the analysis summary in side pane, now passing the document
    ShowAnalysisSummary analysisResult, doc
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

' Function to extract highlight information from document
Private Function GetHighlightInfo(doc As Word.Document, section As String, highlightColor As Long) As HighlightInfo()
    Dim exactText As String
    Dim analysis As String
    Dim rng As Word.Range
    Dim textLines As Variant
    Dim i As Long
    Dim highlights() As HighlightInfo
    Dim highlightCount As Long
    
    ' Initialize
    highlightCount = 0
    
    ' Check if section is empty
    If Len(Trim(section)) = 0 Then
        ' Return empty array
        GetHighlightInfo = highlights
        Exit Function
    End If
    
    ' Extract the exact text and analysis
    exactText = ExtractExactText(section)
    analysis = ExtractAnalysis(section)
    
    ' If exact text is found, collect highlight info
    If Len(Trim(exactText)) > 0 Then
        ' Split into lines in case there are multiple clauses
        textLines = Split(exactText, vbCrLf)
        
        ' First pass to count highlights
        For i = LBound(textLines) To UBound(textLines)
            If Len(Trim(textLines(i))) > 0 Then
                ' Find this text in the document
                Set rng = FindTextInDocument(doc, Trim(textLines(i)))
                
                ' If found, increment counter
                If Not rng Is Nothing Then
                    highlightCount = highlightCount + 1
                End If
            End If
        Next i
        
        ' If we found highlights, initialize array
        If highlightCount > 0 Then
            ReDim highlights(0 To highlightCount - 1)
            
            ' Reset counter for second pass
            highlightCount = 0
            
            ' Second pass to collect highlight info
            For i = LBound(textLines) To UBound(textLines)
                If Len(Trim(textLines(i))) > 0 Then
                    ' Find this text in the document
                    Set rng = FindTextInDocument(doc, Trim(textLines(i)))
                    
                    ' If found, collect info
                    If Not rng Is Nothing Then
                        ' Get page number
                        Dim pageNum As Long
                        pageNum = rng.Information(wdActiveEndPageNumber)
                        
                        ' Add to array
                        highlights(highlightCount).text = Trim(textLines(i))
                        highlights(highlightCount).PageNumber = pageNum
                        highlights(highlightCount).analysis = analysis
                        
                        ' Increment counter
                        highlightCount = highlightCount + 1
                    End If
                End If
            Next i
        End If
    End If
    
    ' Return the array
    GetHighlightInfo = highlights
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
        .text = textToFind
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





