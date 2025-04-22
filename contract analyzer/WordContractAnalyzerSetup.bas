Option Explicit

' Word constants in case they're not recognized
#If VBA7 Then
    ' Word constants we might need
    Private Const wdColorTurquoise As Long = 3
    Private Const wdColorRed As Long = 6
    Private Const wdColorGreen As Long = 4
    Private Const wdAuto As Long = 0
    Private Const wdFindStop As Long = 0
    Private Const wdAlignParagraphCenter As Long = 1
#Else
    ' Word constants we might need
    Private Const wdColorTurquoise As Long = 3
    Private Const wdColorRed As Long = 6
    Private Const wdColorGreen As Long = 4
    Private Const wdAuto As Long = 0
    Private Const wdFindStop As Long = 0
    Private Const wdAlignParagraphCenter As Long = 1
#End If

' Setup instructions for Word Contract Analyzer
' This module provides functions to help users set up and use the Word Contract Analyzer

Public Sub ShowSetupInstructions()
    ' Show setup instructions to the user
    Dim msg As String
    Dim part1 As String
    Dim part2 As String
    Dim part3 As String
    
    ' Break into multiple parts to avoid line continuation limits
    part1 = "WORD CONTRACT ANALYZER SETUP INSTRUCTIONS" & vbCrLf & vbCrLf & _
            "This tool analyzes Word documents to identify and highlight:" & vbCrLf & _
            "  • Payment Terms (highlighted in blue)" & vbCrLf & _
            "  • Limitation of Liability clauses (highlighted in red)" & vbCrLf & _
            "  • Termination clauses (highlighted in green)" & vbCrLf & vbCrLf
    
    part2 = "SETUP REQUIREMENTS:" & vbCrLf & _
            "1. In the VBA Editor, add these references (Tools > References):" & vbCrLf & _
            "   • Microsoft Scripting Runtime" & vbCrLf & _
            "   • Microsoft XML, v6.0 (or available version)" & vbCrLf & _
            "   • Microsoft Word Object Library" & vbCrLf & vbCrLf & _
            "2. Import both modules:" & vbCrLf & _
            "   • WordContractAnalyzer.bas" & vbCrLf & _
            "   • WordContractAnalyzerSetup.bas (this module)" & vbCrLf & vbCrLf
    
    part3 = "3. Run the 'AddMenuItemToWord' function to add the analyzer to Word's menu" & vbCrLf & vbCrLf & _
            "USAGE:" & vbCrLf & _
            "1. Open a contract document in Word" & vbCrLf & _
            "2. Click the 'Contract Analyzer' menu item" & vbCrLf & _
            "3. Wait for analysis to complete" & vbCrLf & _
            "4. Review highlighted sections and comments" & vbCrLf & vbCrLf & _
            "Do you want to add the Contract Analyzer to Word's menu now?"
    
    msg = part1 & part2 & part3
          
    If MsgBox(msg, vbQuestion + vbYesNo, "Word Contract Analyzer Setup") = vbYes Then
        ' Call the function to add the menu item
        AddMenuItemToWord
    End If
End Sub

Public Sub TestApiConnection()
    ' Test the connection to the Google Gemini API
    On Error GoTo ErrorHandler
    
    ' Show a message
    MsgBox "This will test the connection to the Google Gemini API." & vbCrLf & _
           "A small test document will be sent for analysis.", vbInformation
    
    ' Create a simple test prompt
    Dim testPrompt As String
    Dim testText As String
    
    ' Use a shorter text to avoid line continuation issues
    testText = "The vendor shall make payment within 30 days of receiving an invoice. " & _
               "The client may terminate this agreement with 60 days written notice. " & _
               "Liability shall be limited to the amount paid in the preceding 12 months."
    
    ' Get the analysis prompt
    testPrompt = GetAnalysisPrompt(testText)
    
    ' Display a status message
    Dim originalStatusBar As Boolean
    originalStatusBar = Word.Application.DisplayStatusBar
    Word.Application.DisplayStatusBar = True
    Word.Application.StatusBar = "Testing API connection..."
    
    ' Call the API
    Dim result As String
    result = AnalyzeWithGemini(testPrompt)
    
    ' Restore status bar
    Word.Application.StatusBar = False
    Word.Application.DisplayStatusBar = originalStatusBar
    
    ' Check if the result contains expected sections
    If InStr(1, result, "PAYMENT TERMS:", vbTextCompare) > 0 And _
       InStr(1, result, "LIMITATION OF LIABILITY:", vbTextCompare) > 0 And _
       InStr(1, result, "TERMINATION CLAUSES:", vbTextCompare) > 0 Then
        
        ' Success
        MsgBox "API test successful!" & vbCrLf & vbCrLf & _
               "The Google Gemini API is working correctly." & vbCrLf & _
               "You are ready to analyze documents.", vbInformation
    Else
        ' Error
        MsgBox "API test failed or returned unexpected results." & vbCrLf & vbCrLf & _
               "Please check your internet connection and API key." & vbCrLf & _
               "Response received:" & vbCrLf & vbCrLf & _
               Left(result, 500), vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Restore status bar
    On Error Resume Next
    Word.Application.StatusBar = False
    Word.Application.DisplayStatusBar = originalStatusBar
    On Error GoTo 0
    
    ' Show error message
    MsgBox "Error testing API connection: " & Err.Description, vbCritical
End Sub

Public Sub SampleUsage()
    ' Show a sample usage demonstration
    Dim msg As String
    Dim part1 As String
    Dim part2 As String
    Dim part3 As String
    
    part1 = "SAMPLE USAGE GUIDE" & vbCrLf & vbCrLf & _
            "This tool helps legal professionals quickly identify key contract clauses:" & vbCrLf & vbCrLf & _
            "1. PAYMENT TERMS (Turquoise highlight)" & vbCrLf & _
            "   • Standard payment periods" & vbCrLf & _
            "   • Invoice requirements" & vbCrLf & _
            "   • Late payment consequences" & vbCrLf & vbCrLf & _
            "2. RATE CARDS (Yellow highlight)" & vbCrLf & _
            "   • Hourly rate structures" & vbCrLf & _
            "   • Personnel-based fees" & vbCrLf & _
            "   • Rate tables and caps" & vbCrLf & vbCrLf & _
            "3. TRAVEL & EXPENSE POLICY (Gray highlight)" & vbCrLf & _
            "   • Reimbursement terms" & vbCrLf & _
            "   • Travel restrictions" & vbCrLf & _
            "   • Documentation requirements" & vbCrLf & vbCrLf
    
    part2 = "4. DIVERSE SUPPLIER PROVISIONS (Pink highlight)" & vbCrLf & _
            "   • Supplier diversity requirements" & vbCrLf & _
            "   • Reporting obligations" & vbCrLf & _
            "   • Inclusion efforts" & vbCrLf & vbCrLf & _
            "5. TERMINATION CLAUSES (Green highlight)" & vbCrLf & _
            "   • Notice periods" & vbCrLf & _
            "   • Termination for convenience" & vbCrLf & _
            "   • Termination fees" & vbCrLf & vbCrLf & _
            "6. LIMITATION OF LIABILITY (Red highlight)" & vbCrLf & _
            "   • Liability caps" & vbCrLf & _
            "   • Excluded damages" & vbCrLf & _
            "   • Uncapped scenarios" & vbCrLf & vbCrLf
    
    part3 = "7. DATA PRIVACY (Blue highlight)" & vbCrLf & _
            "   • Data protection obligations" & vbCrLf & _
            "   • Regulatory compliance" & vbCrLf & _
            "   • Breach notification requirements" & vbCrLf & vbCrLf & _
            "8. INSURANCE PROVISIONS (Teal highlight)" & vbCrLf & _
            "   • Coverage requirements" & vbCrLf & _
            "   • Coverage limits" & vbCrLf & _
            "   • Policy requirements" & vbCrLf & vbCrLf & _
            "9. BACKGROUND CHECK/DRUG SCREENING (Violet highlight)" & vbCrLf & _
            "   • Background check requirements" & vbCrLf & _
            "   • Drug testing policies" & vbCrLf & _
            "   • Role-specific requirements" & vbCrLf & vbCrLf & _
            "Would you like to try analyzing a document now?"
    
    msg = part1 & part2 & part3
          
    If MsgBox(msg, vbQuestion + vbYesNo, "Word Contract Analyzer Demo") = vbYes Then
        ' Call the main analysis function
        AnalyzeWordDocument
    End If
End Sub

Public Sub CreateSampleDocument()
    ' Create a sample contract document for testing
    On Error GoTo ErrorHandler
    
    ' Create a new document
    Dim doc As Word.Document
    Set doc = Word.Application.Documents.Add
    
    ' Add sample content
    With doc
        ' Add title
        .Content.InsertAfter "SAMPLE CONTRACT AGREEMENT" & vbCrLf & vbCrLf
        
        ' Format title
        doc.Paragraphs(1).Range.Font.Bold = True
        doc.Paragraphs(1).Range.Font.Size = 14
        doc.Paragraphs(1).Alignment = wdAlignParagraphCenter
        
        ' Add introduction
        .Content.InsertAfter "This Agreement is made as of the Effective Date between Company A and Company B." & vbCrLf & vbCrLf
        
        ' Add payment terms section
        .Content.InsertAfter "1. PAYMENT TERMS" & vbCrLf & vbCrLf
        .Content.InsertAfter "1.1 Customer shall pay all invoices within thirty (30) days of receipt. Any amounts not paid when due will accrue interest at the rate of 1.5% per month or the maximum amount allowed by law, whichever is less." & vbCrLf & vbCrLf
        .Content.InsertAfter "1.2 All payments shall be made in US Dollars via electronic funds transfer to the account designated by Company A." & vbCrLf & vbCrLf
        
        ' Add rate cards section
        .Content.InsertAfter "2. RATE CARDS" & vbCrLf & vbCrLf
        .Content.InsertAfter "2.1 Service Provider's hourly rates shall be in accordance with the Rate Card attached as Exhibit A. Rates are categorized by personnel level (Junior, Mid-level, Senior) and geographic location." & vbCrLf & vbCrLf
        .Content.InsertAfter "2.2 Rate Card adjustments may be made annually, with increases not to exceed 3% per year unless approved in writing by Customer." & vbCrLf & vbCrLf
        
        ' Add travel and expense section
        .Content.InsertAfter "3. TRAVEL AND EXPENSE POLICY" & vbCrLf & vbCrLf
        .Content.InsertAfter "3.1 All travel expenses must be pre-approved by Customer. Service Provider shall comply with Customer's Travel and Expense Policy, which requires economy class airfare, standard hotel accommodations, and excludes reimbursement for alcohol, entertainment, or first-class travel." & vbCrLf & vbCrLf
        .Content.InsertAfter "3.2 Expense reports must be submitted within 30 days with original receipts for reimbursement." & vbCrLf & vbCrLf
        
        ' Add diverse supplier section
        .Content.InsertAfter "4. DIVERSE SUPPLIER PROVISIONS" & vbCrLf & vbCrLf
        .Content.InsertAfter "4.1 Service Provider shall make good faith efforts to include diverse suppliers in its supply chain for this Agreement. Service Provider will report quarterly on diverse supplier utilization, including percentage of contract value allocated to minority-owned, women-owned, veteran-owned, and other historically underrepresented business enterprises." & vbCrLf & vbCrLf
        
        ' Add termination section
        .Content.InsertAfter "5. TERMINATION" & vbCrLf & vbCrLf
        .Content.InsertAfter "5.1 Either party may terminate this Agreement for convenience upon sixty (60) days prior written notice to the other party." & vbCrLf & vbCrLf
        .Content.InsertAfter "5.2 Either party may terminate this Agreement immediately upon written notice if the other party materially breaches this Agreement and fails to cure such breach within thirty (30) days after receiving written notice." & vbCrLf & vbCrLf
        .Content.InsertAfter "5.3 Upon termination for Customer's breach or termination for convenience by Customer, Customer shall pay: (a) all unpaid fees committed in the Order Form, and (b) a termination fee equal to 50% of the remaining contract value." & vbCrLf & vbCrLf
        
        ' Add liability section
        .Content.InsertAfter "6. LIMITATION OF LIABILITY" & vbCrLf & vbCrLf
        .Content.InsertAfter "6.1 To the maximum extent permitted by applicable law, in no event shall either party be liable for any indirect, incidental, special, consequential, or punitive damages, including without limitation, loss of profits, data, use, goodwill, or other intangible losses, resulting from this Agreement." & vbCrLf & vbCrLf
        .Content.InsertAfter "6.2 Company A's aggregate liability under this Agreement shall be limited to the total amount paid by Customer during the twelve (12) months preceding the event giving rise to liability." & vbCrLf & vbCrLf
        .Content.InsertAfter "6.3 The limitations in this section shall not apply to either party's indemnification obligations or Customer's payment obligations." & vbCrLf & vbCrLf
        
        ' Add data privacy section
        .Content.InsertAfter "7. DATA PRIVACY" & vbCrLf & vbCrLf
        .Content.InsertAfter "7.1 Service Provider shall comply with all applicable data protection laws, including GDPR and CCPA. Service Provider shall implement appropriate technical and organizational measures to protect personal data processed under this Agreement." & vbCrLf & vbCrLf
        .Content.InsertAfter "7.2 Service Provider shall promptly notify Customer of any data breach affecting Customer data within 48 hours of discovery. Service Provider shall not transfer Customer personal data outside the European Economic Area without Customer's prior written consent." & vbCrLf & vbCrLf
        
        ' Add insurance section
        .Content.InsertAfter "8. INSURANCE PROVISIONS" & vbCrLf & vbCrLf
        .Content.InsertAfter "8.1 Service Provider shall maintain the following insurance coverage during the term of this Agreement: (a) Commercial General Liability insurance with limits of not less than $2,000,000 per occurrence; (b) Professional Liability insurance with limits of not less than $2,000,000 per claim; (c) Workers' Compensation insurance as required by applicable law." & vbCrLf & vbCrLf
        .Content.InsertAfter "8.2 Service Provider shall provide Customer with certificates of insurance upon request and shall name Customer as an additional insured on all policies except Workers' Compensation." & vbCrLf & vbCrLf
        
        ' Add background check section
        .Content.InsertAfter "9. BACKGROUND CHECKS" & vbCrLf & vbCrLf
        .Content.InsertAfter "9.1 Service Provider shall conduct background checks on all personnel assigned to Customer projects or who will have access to Customer facilities or systems. Background checks shall include criminal history, employment verification, and education verification." & vbCrLf & vbCrLf
        .Content.InsertAfter "9.2 Service Provider shall conduct drug testing for all personnel prior to assignment to Customer projects, and shall comply with Customer's drug-free workplace policy while on Customer premises." & vbCrLf & vbCrLf
        
        ' Format headings
        Dim para As Word.Paragraph
        For Each para In .Paragraphs
            If InStr(para.Range.Text, "1. PAYMENT TERMS") > 0 Or _
               InStr(para.Range.Text, "2. RATE CARDS") > 0 Or _
               InStr(para.Range.Text, "3. TRAVEL AND EXPENSE POLICY") > 0 Or _
               InStr(para.Range.Text, "4. DIVERSE SUPPLIER PROVISIONS") > 0 Or _
               InStr(para.Range.Text, "5. TERMINATION") > 0 Or _
               InStr(para.Range.Text, "6. LIMITATION OF LIABILITY") > 0 Or _
               InStr(para.Range.Text, "7. DATA PRIVACY") > 0 Or _
               InStr(para.Range.Text, "8. INSURANCE PROVISIONS") > 0 Or _
               InStr(para.Range.Text, "9. BACKGROUND CHECKS") > 0 Then
                para.Range.Font.Bold = True
            End If
        Next para
    End With
    
    ' Ask if user wants to analyze the sample
    If MsgBox("Sample contract document created!" & vbCrLf & vbCrLf & _
              "Would you like to analyze this sample document now?", _
              vbQuestion + vbYesNo, "Sample Document Created") = vbYes Then
        
        ' Run the analyzer
        AnalyzeWordDocument
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating sample document: " & Err.Description, vbCritical
End Sub

' Helper function to add the menu item to Word
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