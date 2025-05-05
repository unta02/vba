Attribute VB_Name = "SOWBuilderSinglePage"
Option Explicit

' Main entry point to show the SOW Builder
Public Sub ShowSOWBuilder()
    ' Load and show the form
    frmSOWBuilder3.Show
End Sub

' Generate SOW document based on the form inputs
Public Sub GenerateSOWDocument(ByVal clientInfo As Object, ByVal compensationOption As String, _
                              ByVal annualFee As String, ByVal billingOption As String, _
                              ByVal policies As Object, ByVal optionalClauses As Object, _
                              ByVal additionalNotes As String)
    On Error GoTo ErrorHandler
    
    ' Type validation
    If TypeName(clientInfo) <> "Dictionary" Then
        Debug.Print "Warning: clientInfo is not a Dictionary object, it is a " & TypeName(clientInfo)
    End If
    
    If TypeName(policies) <> "Collection" Then
        Debug.Print "Warning: policies is not a Collection object, it is a " & TypeName(policies)
    End If
    
    If TypeName(optionalClauses) <> "Dictionary" Then
        Debug.Print "Warning: optionalClauses is not a Dictionary object, it is a " & TypeName(optionalClauses)
    End If
    
    ' Make sure annualFee is a string or can be converted to a string
    Dim feeStr As String
    If IsNumeric(annualFee) Then
        feeStr = CStr(annualFee)
    Else
        feeStr = annualFee
    End If
    
    Dim doc As Document
    Set doc = Documents.Add
    
    ' Create the document content
    FillSOWDocument doc, clientInfo, compensationOption, feeStr, billingOption, _
                   policies, optionalClauses, additionalNotes
    
    ' Format document
    FormatSOWDocument doc
    
    ' Display success message
    MsgBox "SOW document has been generated successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in GenerateSOWDocument: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
    Debug.Print "Error in GenerateSOWDocument: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Resume Next
End Sub

' Fill the document with all the embedded text and user inputs
Private Sub FillSOWDocument(doc As Document, ByVal clientInfo As Object, _
                          ByVal compensationOption As String, ByVal annualFee As String, _
                          ByVal billingOption As String, ByVal policies As Object, _
                          ByVal optionalClauses As Object, ByVal additionalNotes As String)
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = doc.Content
    
    ' Clear existing content
    rng.Text = ""
    
    ' Collapse to start
    rng.Collapse Direction:=wdCollapseStart
    
    ' Add document header and client information
    AddDocumentHeader rng, clientInfo
    
    ' Add Terms and Conditions section
    AddTermsAndConditionsSection rng
    
    ' Add Term and Termination section
    AddTermAndTerminationSection rng, clientInfo, optionalClauses
    
    ' Add Compensation section based on selected option
    AddCompensationSection rng, compensationOption, annualFee, billingOption, policies
    
    ' Add Additional Terms section
    AddAdditionalTermsSection rng, optionalClauses, additionalNotes
    
    ' Add signature blocks
    AddSignatureBlocks rng, clientInfo
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in FillSOWDocument: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Add document header and client information
Private Sub AddDocumentHeader(rng As Range, clientInfo As Object)
    On Error GoTo ErrorHandler
    
    ' Add document title
    rng.InsertAfter "STATEMENT OF WORK" & vbCrLf & vbCrLf
    
    ' Ensure clientInfo is a Dictionary type
    If TypeName(clientInfo) <> "Dictionary" Then
        rng.InsertAfter "[Date]" & vbCrLf & vbCrLf
        rng.InsertAfter "[Contact Name]" & vbCrLf
        rng.InsertAfter "[Company Name]" & vbCrLf
        rng.InsertAfter "[Address Line 1]" & vbCrLf
        rng.InsertAfter "[Address Line 2]" & vbCrLf & vbCrLf
        rng.InsertAfter "Subject: Statement of Work for Health & Benefits Services" & vbCrLf & vbCrLf
        rng.InsertAfter "Dear [Name]:" & vbCrLf & vbCrLf
        rng.InsertAfter "This statement of work (""SOW"") will confirm the terms of the engagement of [WTW Entity] (""WTW"", ""we"" or ""us"") by [Client Legal Name] (""Client"" or ""you"")." & vbCrLf & vbCrLf
        Exit Sub
    End If
    
    ' Add date
    If clientInfo.Exists("Date") Then
        rng.InsertAfter clientInfo("Date") & vbCrLf & vbCrLf
    Else
        rng.InsertAfter "[Date]" & vbCrLf & vbCrLf
    End If
    
    ' Add client contact information
    If clientInfo.Exists("ContactName") Then
        rng.InsertAfter clientInfo("ContactName") & vbCrLf
    Else
        rng.InsertAfter "[Contact Name]" & vbCrLf
    End If
    
    If clientInfo.Exists("CompanyName") Then
        rng.InsertAfter clientInfo("CompanyName") & vbCrLf
    Else
        rng.InsertAfter "[Company Name]" & vbCrLf
    End If
    
    If clientInfo.Exists("Address1") And Trim(clientInfo("Address1")) <> "" Then
        rng.InsertAfter clientInfo("Address1") & vbCrLf
    Else
        rng.InsertAfter "[Address Line 1]" & vbCrLf
    End If
    
    If clientInfo.Exists("Address2") And Trim(clientInfo("Address2")) <> "" Then
        rng.InsertAfter clientInfo("Address2") & vbCrLf
    Else
        rng.InsertAfter "[Address Line 2]" & vbCrLf
    End If
    
    rng.InsertAfter vbCrLf
    
    ' Add subject line with bold formatting
    Dim subjectStart As Long
    subjectStart = rng.End
    rng.InsertAfter "Subject: Statement of Work for Health & Benefits Services" & vbCrLf & vbCrLf
    
    ' Set the subject line to bold
    Dim subjectRange As Range
    Set subjectRange = rng.Document.Range(subjectStart, rng.Document.Range(subjectStart).End - 2)
    subjectRange.Bold = True
    
    ' Add salutation
    If clientInfo.Exists("ContactName") Then
        rng.InsertAfter "Dear " & clientInfo("ContactName") & ":" & vbCrLf & vbCrLf
    Else
        rng.InsertAfter "Dear [Name]:" & vbCrLf & vbCrLf
    End If
    
    ' Add opening paragraph
    If clientInfo.Exists("WTWParty") And clientInfo.Exists("ClientName") Then
        ' Insert text with parts to be bolded marked
        Dim openingStart As Long
        openingStart = rng.End
        rng.InsertAfter "This statement of work ("
        
        ' Make "SOW" bold
        Dim sowStart As Long
        sowStart = rng.End
        rng.InsertAfter "SOW"
        Dim sowRange As Range
        Set sowRange = rng.Document.Range(sowStart, rng.End)
        sowRange.Bold = True
        
        rng.InsertAfter ") will confirm the terms of the engagement of " & clientInfo("WTWParty") & " ("
        
        ' Make "WTW", "we", "us" bold
        Dim wtwStart As Long
        wtwStart = rng.End
        rng.InsertAfter "WTW"
        Dim wtwRange As Range
        Set wtwRange = rng.Document.Range(wtwStart, rng.End)
        wtwRange.Bold = True
        
        rng.InsertAfter ", "
        
        Dim weStart As Long
        weStart = rng.End
        rng.InsertAfter "we"
        Dim weRange As Range
        Set weRange = rng.Document.Range(weStart, rng.End)
        weRange.Bold = True
        
        rng.InsertAfter " or "
        
        Dim usStart As Long
        usStart = rng.End
        rng.InsertAfter "us"
        Dim usRange As Range
        Set usRange = rng.Document.Range(usStart, rng.End)
        usRange.Bold = True
        
        rng.InsertAfter ") by " & clientInfo("ClientName") & " ("
        
        ' Make "Client", "you" bold
        Dim clientStart As Long
        clientStart = rng.End
        rng.InsertAfter "Client"
        Dim clientRange As Range
        Set clientRange = rng.Document.Range(clientStart, rng.End)
        clientRange.Bold = True
        
        rng.InsertAfter " or "
        
        Dim youStart As Long
        youStart = rng.End
        rng.InsertAfter "you"
        Dim youRange As Range
        Set youRange = rng.Document.Range(youStart, rng.End)
        youRange.Bold = True
        
        rng.InsertAfter ")." & vbCrLf & vbCrLf
    Else
        ' Insert text with parts to be bolded marked
        Dim defaultOpeningStart As Long
        defaultOpeningStart = rng.End
        rng.InsertAfter "This statement of work ("
        
        ' Make "SOW" bold
        Dim defaultSowStart As Long
        defaultSowStart = rng.End
        rng.InsertAfter "SOW"
        Dim defaultSowRange As Range
        Set defaultSowRange = rng.Document.Range(defaultSowStart, rng.End)
        defaultSowRange.Bold = True
        
        rng.InsertAfter ") will confirm the terms of the engagement of [WTW Entity] ("
        
        ' Make "WTW", "we", "us" bold
        Dim defaultWtwStart As Long
        defaultWtwStart = rng.End
        rng.InsertAfter "WTW"
        Dim defaultWtwRange As Range
        Set defaultWtwRange = rng.Document.Range(defaultWtwStart, rng.End)
        defaultWtwRange.Bold = True
        
        rng.InsertAfter ", "
        
        Dim defaultWeStart As Long
        defaultWeStart = rng.End
        rng.InsertAfter "we"
        Dim defaultWeRange As Range
        Set defaultWeRange = rng.Document.Range(defaultWeStart, rng.End)
        defaultWeRange.Bold = True
        
        rng.InsertAfter " or "
        
        Dim defaultUsStart As Long
        defaultUsStart = rng.End
        rng.InsertAfter "us"
        Dim defaultUsRange As Range
        Set defaultUsRange = rng.Document.Range(defaultUsStart, rng.End)
        defaultUsRange.Bold = True
        
        rng.InsertAfter ") by [Client Legal Name] ("
        
        ' Make "Client", "you" bold
        Dim defaultClientStart As Long
        defaultClientStart = rng.End
        rng.InsertAfter "Client"
        Dim defaultClientRange As Range
        Set defaultClientRange = rng.Document.Range(defaultClientStart, rng.End)
        defaultClientRange.Bold = True
        
        rng.InsertAfter " or "
        
        Dim defaultYouStart As Long
        defaultYouStart = rng.End
        rng.InsertAfter "you"
        Dim defaultYouRange As Range
        Set defaultYouRange = rng.Document.Range(defaultYouStart, rng.End)
        defaultYouRange.Bold = True
        
        rng.InsertAfter ")." & vbCrLf & vbCrLf
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddDocumentHeader: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Debug.Print "clientInfo type: " & TypeName(clientInfo)
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Add Terms and Conditions section
Private Sub AddTermsAndConditionsSection(rng As Range)
    ' Add section header with proper formatting
    Dim sectionStart As Long
    sectionStart = rng.End
    
    rng.InsertAfter "I. Terms and Conditions of SOW:" & vbCrLf & vbCrLf
    
    ' Format the section header as bold
    Dim sectionRange As Range
    Set sectionRange = rng.Document.Range(sectionStart, rng.Document.Range(sectionStart).End - 2)
    sectionRange.Bold = True
    
    ' Insert text with parts to be bolded
    rng.InsertAfter "Client desires to procure and WTW is willing to provide the services listed in Attachment 1 (the "
    
    ' Make "Services" bold
    Dim servicesStart As Long
    servicesStart = rng.End
    rng.InsertAfter "Services"
    Dim servicesRange As Range
    Set servicesRange = rng.Document.Range(servicesStart, rng.End)
    servicesRange.Bold = True
    
    rng.InsertAfter "). These Services will be provided subject to the WTW Health & Benefits Brokerage Terms, Conditions & Disclosures available at: " & _
                 "https://www.wtwco.com/-/media/WTW/Notices/h-b-brokerage-terms-no-MSA.pdf (the "
    
    ' Make "Brokerage Terms" bold
    Dim brokTermsStart As Long
    brokTermsStart = rng.End
    rng.InsertAfter "Brokerage Terms"
    Dim brokTermsRange As Range
    Set brokTermsRange = rng.Document.Range(brokTermsStart, rng.End)
    brokTermsRange.Bold = True
    
    rng.InsertAfter "). Copies of the Brokerage Terms are available upon request." & vbCrLf & vbCrLf
End Sub

' Add Term and Termination section
Private Sub AddTermAndTerminationSection(rng As Range, clientInfo As Object, optionalClauses As Object)
    On Error GoTo ErrorHandler
    
    ' Add section header with proper formatting
    Dim sectionStart As Long
    sectionStart = rng.End
    
    rng.InsertAfter "II. Term and Termination:" & vbCrLf & vbCrLf
    
    ' Format the section header as bold
    Dim sectionRange As Range
    Set sectionRange = rng.Document.Range(sectionStart, rng.Document.Range(sectionStart).End - 2)
    sectionRange.Bold = True
    
    ' Ensure clientInfo is a Dictionary and has the required keys
    If TypeName(clientInfo) = "Dictionary" Then
        If clientInfo.Exists("StartDate") And clientInfo.Exists("EndDate") Then
            rng.InsertAfter "The term of this SOW will begin on " & clientInfo("StartDate") & " and end on " & _
                         clientInfo("EndDate") & ". Either party may terminate this SOW upon 60 days prior written notice to the other party." & vbCrLf & vbCrLf
        Else
            rng.InsertAfter "The term of this SOW will begin on _______________ and end on _______________. " & _
                         "Either party may terminate this SOW upon 60 days prior written notice to the other party." & vbCrLf & vbCrLf
        End If
    Else
        rng.InsertAfter "The term of this SOW will begin on _______________ and end on _______________. " & _
                     "Either party may terminate this SOW upon 60 days prior written notice to the other party." & vbCrLf & vbCrLf
    End If
    
    ' Add auto-renewal if selected
    If TypeName(optionalClauses) = "Dictionary" Then
        If optionalClauses.Exists("AutoRenewal") Then
            If optionalClauses("AutoRenewal") Then
                rng.InsertAfter "Upon the expiration of the term, or any renewal term, this SOW will renew automatically for successive one-year terms " & _
                             "unless either party gives notice of non-renewal at least 60 days before the scheduled expiration date." & vbCrLf & vbCrLf
            End If
        End If
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddTermAndTerminationSection: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Debug.Print "clientInfo type: " & TypeName(clientInfo)
    Debug.Print "optionalClauses type: " & TypeName(optionalClauses)
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Add Compensation section based on selected option
Private Sub AddCompensationSection(rng As Range, compensationOption As String, _
                                  annualFee As String, billingOption As String, policies As Object)
    On Error GoTo ErrorHandler
    
    ' Add section header with proper formatting
    Dim sectionStart As Long
    sectionStart = rng.End
    
    rng.InsertAfter "III. Compensation" & vbCrLf & vbCrLf
    
    ' Format the section header as bold
    Dim sectionRange As Range
    Set sectionRange = rng.Document.Range(sectionStart, rng.Document.Range(sectionStart).End - 2)
    sectionRange.Bold = True
    
    Select Case compensationOption
        Case "A"
            InsertFeeOnly rng, annualFee, billingOption
        Case "B"
            InsertFeePlusCommission rng, annualFee, billingOption, policies
        Case "C"
            InsertFeeOffset rng, annualFee, billingOption, policies
        Case "D"
            InsertCommissionOnly rng, policies
        Case Else
            ' Default to fee only if option not recognized
            InsertFeeOnly rng, annualFee, billingOption
    End Select
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddCompensationSection: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Debug.Print "compensationOption: " & compensationOption
    Debug.Print "annualFee: " & annualFee
    Debug.Print "billingOption: " & billingOption
    Debug.Print "policies type: " & TypeName(policies)
    If TypeName(policies) = "Collection" Then
        Debug.Print "policies count: " & policies.Count
    End If
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Insert fee-only compensation option
Private Sub InsertFeeOnly(rng As Range, annualFee As String, billingOption As String)
    rng.InsertAfter "You agree that our compensation for the Services will be an annual fee of $" & _
                 annualFee & ", payable by you to us as follows." & vbCrLf & vbCrLf
    
    ' Insert billing details
    InsertBillingDetails rng, billingOption
    
    ' Insert expenses paragraph
    InsertExpensesParagraph rng
    
    ' Insert fee statement
    rng.InsertAfter "The fee is in addition to the premiums you must pay for your policies. " & _
                 "Information regarding other compensation we may receive is described in the Brokerage Terms." & vbCrLf & vbCrLf
    
    ' Insert earned fee table
    InsertEarnedFeeTable rng
End Sub

' Insert fee plus commission option
Private Sub InsertFeePlusCommission(rng As Range, annualFee As String, billingOption As String, policies As Object)
    On Error GoTo ErrorHandler
    
    rng.InsertAfter "You agree that our compensation for the Services will be an annual fee of $" & _
                 annualFee & ", payable by you to us as follows." & vbCrLf & vbCrLf
    
    ' Insert billing details
    InsertBillingDetails rng, billingOption
    
    ' Insert expenses paragraph
    InsertExpensesParagraph rng
    
    ' Insert fee statement
    rng.InsertAfter "The fee is in addition to the premiums you must pay for your policies. " & _
                 "Information regarding other compensation we may receive is described in the Brokerage Terms." & vbCrLf & vbCrLf
    
    ' Insert commission paragraph
    rng.InsertAfter "You also agree that, in addition to the fee, we will be entitled to compensation in the form of " & _
                 "commissions paid to us by insurers for the sale of the following insurance policies:" & vbCrLf & vbCrLf
    
    ' Insert policy list
    If TypeName(policies) = "Collection" Then
        If policies.Count > 0 Then
            InsertPolicyList rng, policies
        Else
            rng.InsertAfter "[Policy list to be determined]" & vbCrLf & vbCrLf
        End If
    Else
        rng.InsertAfter "[Policy list to be determined]" & vbCrLf & vbCrLf
    End If
    
    ' Insert additional explanation
    rng.InsertAfter "To the extent that we receive both fees and commissions for Services related to the same insurance policies, " & _
                 "the commissions compensate us for the placement and servicing of those policies, while the fee compensates us " & _
                 "for the Services which are in addition to the placement and routine servicing of the policies." & vbCrLf & vbCrLf
    
    ' Insert commission adjustment clause
    rng.InsertAfter "The parties agree that should the commissions increase or decrease by an amount exceeding ten percent (10%), " & _
                 "due to a change in covered lives, added or deleted policies or for other reasons, the parties will discuss " & _
                 "changes to our compensation and/or Services. Any such changes will be agreed to in writing by the parties." & vbCrLf & vbCrLf
    
    ' Insert earned fee table
    InsertEarnedFeeTable rng
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertFeePlusCommission: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Insert fee offset by commission option
Private Sub InsertFeeOffset(rng As Range, annualFee As String, billingOption As String, policies As Object)
    On Error GoTo ErrorHandler
    
    rng.InsertAfter "You agree that our compensation for the Services will be an annual fee of $" & _
                 annualFee & ", payable by you to us as follows." & vbCrLf & vbCrLf
    
    ' Insert billing details
    InsertBillingDetails rng, billingOption
    
    ' Insert expenses paragraph
    InsertExpensesParagraph rng
    
    ' Insert fee statement
    rng.InsertAfter "The fee is in addition to the premiums you must pay for your policies. " & _
                 "Information regarding other compensation we may receive is described in the Brokerage Terms." & vbCrLf & vbCrLf
    
    ' Insert offset paragraph
    rng.InsertAfter "To the extent that we also receive during the term of this SOW commissions paid by insurers for the sale of " & _
                 "the insurance policies that you purchase, we will use those base commissions to offset our fee, but only to " & _
                 "the extent allowable by law. You acknowledge that we cannot return commissions to you under any circumstance. " & _
                 "We will account to you periodically during the term of the SOW and at the termination of the SOW for all " & _
                 "commissions received." & vbCrLf & vbCrLf
    
    ' Insert policy list
    If TypeName(policies) = "Collection" Then
        If policies.Count > 0 Then
            InsertPolicyList rng, policies
        Else
            rng.InsertAfter "[Policy list to be determined]" & vbCrLf & vbCrLf
        End If
    Else
        rng.InsertAfter "[Policy list to be determined]" & vbCrLf & vbCrLf
    End If
    
    ' Insert earned fee table
    InsertEarnedFeeTable rng
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertFeeOffset: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Insert commission-only option
Private Sub InsertCommissionOnly(rng As Range, policies As Object)
    On Error GoTo ErrorHandler
    
    rng.InsertAfter "You agree that we will be compensated by commissions paid to us by insurers for the sale of the insurance policies " & _
                 "that you purchase. All commissions will be fully disclosed to you prior to our placing coverage. The commissions " & _
                 "will be earned for the entire policy period at the time we place insurance policies for you." & vbCrLf & vbCrLf
    
    ' Insert policy list if there are policies
    If TypeName(policies) = "Collection" Then
        If policies.Count > 0 Then
            InsertPolicyList rng, policies
        Else
            rng.InsertAfter "[Policy list to be determined]" & vbCrLf & vbCrLf
        End If
    Else
        rng.InsertAfter "[Policy list to be determined]" & vbCrLf & vbCrLf
    End If
    
    ' Insert commission adjustment clause
    rng.InsertAfter "The parties agree that should commissions increase or decrease by an amount exceeding ten percent (10%), " & _
                 "due to a change in covered lives, added or deleted policies or for other reasons, the parties will discuss " & _
                 "changes to our compensation and/or Services. Any such changes will be agreed to in writing by the parties." & vbCrLf & vbCrLf
    
    ' Insert additional information
    rng.InsertAfter "Information regarding other compensation we may receive is described in the Brokerage Terms." & vbCrLf & vbCrLf
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertCommissionOnly: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Insert billing details based on selection
Private Sub InsertBillingDetails(rng As Range, billingOption As String)
    If billingOption = "Milestone" Then
        rng.InsertAfter "in installments with 30% of the fee due upon execution of this statement of work, 50% of the fee due at the " & _
                     "completion of the plan selection, and 20% of the fee due upon completion of our Services." & vbCrLf & vbCrLf
    ElseIf billingOption = "Installments" Then
        rng.InsertAfter "at the beginning of each quarter (or month, or in equal installments over the period of the project " & _
                     "if shorter than a year) beginning with the execution of this SOW." & vbCrLf & vbCrLf
    End If
End Sub

' Insert expenses paragraph
Private Sub InsertExpensesParagraph(rng As Range)
    rng.InsertAfter "In addition to the fee, our charges will include the following:" & vbCrLf & vbCrLf
    
    ' Use proper indentation for bullet points with paragraph formatting
    Dim bulletPara1Start As Long
    bulletPara1Start = rng.End
    
    rng.InsertAfter "• reimbursement, at cost, of direct expenses reasonably incurred by us in connection with the performance " & _
                 "of our Services, such as travel and other vendor expenses, and itemized extraordinary expenses such as " & _
                 "large-volume color printing, large-volume courier shipments and the like, plus an administrative fee of 5% " & _
                 "of any vendor charges other than travel, unless arrangements are made in advance for charges to be invoiced " & _
                 "to and paid by you directly; and" & vbCrLf & vbCrLf
    
    ' Apply left indent to the first bullet
    Dim bulletPara1 As Range
    Set bulletPara1 = rng.Document.Range(bulletPara1Start, rng.End - 2)  ' -2 to account for vbCrLf
    With bulletPara1.ParagraphFormat
        .LeftIndent = 36  ' 0.5 inch in points
        .FirstLineIndent = -18  ' Hanging indent for bullet
    End With
    
    Dim bulletPara2Start As Long
    bulletPara2Start = rng.End
    
    rng.InsertAfter "• the amount of any tax or similar assessment based upon our charges." & vbCrLf & vbCrLf
    
    ' Apply left indent to the second bullet
    Dim bulletPara2 As Range
    Set bulletPara2 = rng.Document.Range(bulletPara2Start, rng.End - 2)  ' -2 to account for vbCrLf
    With bulletPara2.ParagraphFormat
        .LeftIndent = 36  ' 0.5 inch in points
        .FirstLineIndent = -18  ' Hanging indent for bullet
    End With
    
    rng.InsertAfter "We will bill you for the fee payments as they become due. At the end of each month during which we perform " & _
                 "Services for you, we will also bill you for all other charges accrued for the month, such as travel and " & _
                 "vendor expenses." & vbCrLf & vbCrLf
End Sub

' Insert policy list from collection
Private Sub InsertPolicyList(rng As Range, policies As Object)
    On Error GoTo ErrorHandler
    
    ' Validate policies parameter
    If TypeName(policies) <> "Collection" Then
        Exit Sub
    End If
    
    ' Add safeguard against null or empty collection
    If policies Is Nothing Then
        Exit Sub
    End If
    
    If policies.Count = 0 Then
        Exit Sub
    End If
    
    Dim policy As Variant
    
    For Each policy In policies
        rng.InsertAfter "• " & policy & vbCrLf
    Next policy
    
    rng.InsertAfter vbCrLf
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InsertPolicyList: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Debug.Print "Policies type: " & TypeName(policies)
    If TypeName(policies) = "Collection" And Not policies Is Nothing Then
        Debug.Print "Policies count: " & policies.Count
    End If
    Resume Next ' Continue execution despite errors in this function
End Sub

' Insert earned fee table for options A, B, and C
Private Sub InsertEarnedFeeTable(rng As Range)
    rng.InsertAfter "You acknowledge that, even though we may regularly invoice you on a different schedule during the term of this " & _
                 "SOW, a substantial portion of our work is provided prior to and at the effective date of your benefit plan. " & _
                 "Therefore, if this SOW is terminated before the end of the term, in order to compensate us fully for the " & _
                 "services actually provided to you, the parties agree that the fee is earned and that you will pay us as " & _
                 "provided in the following table:" & vbCrLf & vbCrLf
    
    ' Create a table instead of text format
    Dim doc As Document
    Set doc = rng.Document
    
    ' Insert the table at the end of the range
    Dim tbl As Table
    Set tbl = doc.Tables.Add(Range:=rng, NumRows:=3, NumColumns:=3)
    
    ' Format the table
    With tbl
        ' Add borders to the table
        .Borders.InsideLineStyle = wdLineStyleSingle
        .Borders.OutsideLineStyle = wdLineStyleSingle
        
        ' Set column widths (in points)
        .Columns(1).Width = 150 ' Component column
        .Columns(2).Width = 60  ' Percentage column
        .Columns(3).Width = 300 ' Description column
        
        ' Row 1: Strategic Planning
        .Cell(1, 1).Range.Text = "Strategic Planning"
        .Cell(1, 2).Range.Text = "15%"
        .Cell(1, 3).Range.Text = "Earned in equal monthly installments prior to the benefit plan effective date (fully earned at benefit plan effective date)"
        
        ' Row 2: Program Renewal / Placement Process
        .Cell(2, 1).Range.Text = "Program Renewal / Placement Process"
        .Cell(2, 2).Range.Text = "35%"
        .Cell(2, 3).Range.Text = "Earned in equal monthly installments prior to the benefit plan effective date (fully earned at benefit plan effective date)"
        
        ' Row 3: Ongoing Service and Resources
        .Cell(3, 1).Range.Text = "Ongoing Service and Resources"
        .Cell(3, 2).Range.Text = "50%"
        .Cell(3, 3).Range.Text = "Earned in 12 equal monthly installments (starting at benefit plan effective date)"
        
        ' Center the percentage column
        .Columns(2).Select
        doc.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
    
    ' Add paragraph after the table
    rng.InsertAfter vbCrLf
End Sub

' Add Additional Terms section
Private Sub AddAdditionalTermsSection(rng As Range, optionalClauses As Object, additionalNotes As String)
    On Error GoTo ErrorHandler
    
    ' Add section header with proper formatting
    Dim sectionStart As Long
    sectionStart = rng.End
    
    rng.InsertAfter "IV. Additional Terms" & vbCrLf & vbCrLf
    
    ' Format the section header as bold
    Dim sectionRange As Range
    Set sectionRange = rng.Document.Range(sectionStart, rng.Document.Range(sectionStart).End - 2)
    sectionRange.Bold = True
    
    ' Add GDPR clause if applicable
    If TypeName(optionalClauses) = "Dictionary" Then
        If optionalClauses.Exists("GDPR") Then
            If optionalClauses("GDPR") Then
                rng.InsertAfter "The parties acknowledge that WTW may access personal data from Europe that could trigger GDPR requirements. " & _
                             "For Health & Benefits Brokerage services, WTW will act as a data controller in accordance with applicable " & _
                             "data protection laws." & vbCrLf & vbCrLf
            End If
        End If
    End If
    
    ' Add standard out-of-scope services paragraph
    rng.InsertAfter "If you would like us to vary the Services under this SOW, or to perform additional services that are not included, " & _
                 "please advise us. Also, if we believe certain services you have asked us to carry out are not within the defined scope, " & _
                 "we will promptly notify you. All out of scope services will be covered under a separate Statement of Work that will " & _
                 "specify the additional services that we will perform and the additional compensation that we will receive." & vbCrLf & vbCrLf
    
    ' Add any custom notes if provided
    If Trim(additionalNotes) <> "" Then
        rng.InsertAfter "Additional Notes: " & additionalNotes & vbCrLf & vbCrLf
    End If
    
    ' Add closing
    rng.InsertAfter "Please have an authorized representative of Client countersign below (and do the same in respect of the enclosed copies, " & _
                 "returning a set of countersigned documents to me for our records)." & vbCrLf & vbCrLf
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddAdditionalTermsSection: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Debug.Print "optionalClauses type: " & TypeName(optionalClauses)
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Add signature blocks
Private Sub AddSignatureBlocks(rng As Range, clientInfo As Object)
    On Error GoTo ErrorHandler
    
    rng.InsertAfter "IN WITNESS WHEREOF, the parties have executed this SOW effective as of the _____ day of _______________, 20__." & vbCrLf & vbCrLf
    
    ' Check if clientInfo is a Dictionary and has the WTWParty key
    If TypeName(clientInfo) = "Dictionary" And clientInfo.Exists("WTWParty") Then
        rng.InsertAfter "Signed by and on behalf of " & clientInfo("WTWParty") & vbCrLf & vbCrLf
    Else
        rng.InsertAfter "Signed by and on behalf of [WTW Entity]" & vbCrLf & vbCrLf
    End If
    
    rng.InsertAfter "By: ____________________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print name: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print title: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Date: __________________________" & vbCrLf & vbCrLf
    
    ' Check if clientInfo is a Dictionary and has the ClientName key
    If TypeName(clientInfo) = "Dictionary" And clientInfo.Exists("ClientName") Then
        rng.InsertAfter "Accepted and agreed on behalf of " & clientInfo("ClientName") & vbCrLf & vbCrLf
    Else
        rng.InsertAfter "Accepted and agreed on behalf of [Client Legal Name]" & vbCrLf & vbCrLf
    End If
    
    rng.InsertAfter "By: ____________________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print name: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print title: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Date: __________________________" & vbCrLf & vbCrLf
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in AddSignatureBlocks: " & Err.Description
    Debug.Print "Error number: " & Err.Number
    Debug.Print "Error source: " & Err.Source
    Debug.Print "clientInfo type: " & TypeName(clientInfo)
    Resume Next ' Continue execution rather than re-raising the error
End Sub

' Apply formatting to the document
Private Sub FormatSOWDocument(doc As Document)
    ' Apply basic formatting
    With doc
        ' Format title
        If .Paragraphs.Count > 0 Then
            .Paragraphs(1).Range.Font.Bold = True
            .Paragraphs(1).Range.Font.Size = 14
            .Paragraphs(1).Alignment = wdAlignParagraphCenter
        End If
        
        ' Format section headers
        Dim para As Paragraph
        For Each para In .Paragraphs
            If InStr(para.Range.Text, "I. Terms and Conditions") > 0 Or _
               InStr(para.Range.Text, "II. Term and Termination") > 0 Or _
               InStr(para.Range.Text, "III. Compensation") > 0 Or _
               InStr(para.Range.Text, "IV. Additional Terms") > 0 Then
                para.Range.Font.Bold = True
                para.Range.Font.Size = 11
                para.Range.Font.Underline = wdUnderlineSingle
                para.Format.SpaceAfter = 8
                para.Format.SpaceBefore = 12
            End If
        Next para
        
        ' Set default paragraph formatting
        .Paragraphs.Format.SpaceAfter = 8
        .Paragraphs.Format.SpaceBefore = 0
        .Paragraphs.Format.LineSpacing = 12 ' Single spacing (12 points)
        
        ' Default font
        .Content.Font.Name = "Arial"
        .Content.Font.Size = 10
    End With
End Sub

' Helper function to create a Dictionary object
Public Function CreateDictionary() As Object
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
End Function

' Helper function to create a Collection object
Public Function CreateCollection() As Collection
    Set CreateCollection = New Collection
End Function 
