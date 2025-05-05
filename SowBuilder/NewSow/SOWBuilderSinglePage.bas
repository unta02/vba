Attribute VB_Name = "SOWBuilderSinglePage"
Option Explicit

' Main entry point to show the SOW Builder
Public Sub ShowSOWBuilder()
    ' Load and show the form
    frmSOWBuilder3.Show
End Sub

' Generate SOW document based on the form inputs
Public Sub GenerateSOWDocument(ByVal clientInfo As Dictionary, ByVal compensationOption As String, _
                              ByVal annualFee As String, ByVal billingOption As String, _
                              ByVal policies As Collection, ByVal optionalClauses As Dictionary, _
                              ByVal additionalNotes As String)
    Dim doc As Document
    Set doc = Documents.Add
    
    ' Create the document content
    FillSOWDocument doc, clientInfo, compensationOption, annualFee, billingOption, _
                   policies, optionalClauses, additionalNotes
    
    ' Format document
    FormatSOWDocument doc
    
    ' Display success message
    MsgBox "SOW document has been generated successfully!", vbInformation
End Sub

' Fill the document with all the embedded text and user inputs
Private Sub FillSOWDocument(doc As Document, ByVal clientInfo As Dictionary, _
                          ByVal compensationOption As String, ByVal annualFee As String, _
                          ByVal billingOption As String, ByVal policies As Collection, _
                          ByVal optionalClauses As Dictionary, ByVal additionalNotes As String)
    Dim rng As Range
    Set rng = doc.Content
    
    ' Clear existing content
    rng.Text = ""
    
    ' Collapse to start
    rng.Collapse Direction:=wdCollapseStart
    
    ' Add document title and header
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
    
    ' Add Attachment for Scope of Services
    AddAttachment rng
End Sub

' Add document header and client information
Private Sub AddDocumentHeader(rng As Range, clientInfo As Dictionary)
    ' Add document title
    rng.InsertAfter "STATEMENT OF WORK" & vbCrLf & vbCrLf
    
    ' Add date
    rng.InsertAfter clientInfo("Date") & vbCrLf & vbCrLf
    
    ' Add client contact information
    rng.InsertAfter clientInfo("ContactName") & vbCrLf
    rng.InsertAfter clientInfo("CompanyName") & vbCrLf
    
    If Trim(clientInfo("Address1")) <> "" Then
        rng.InsertAfter clientInfo("Address1") & vbCrLf
    End If
    
    If Trim(clientInfo("Address2")) <> "" Then
        rng.InsertAfter clientInfo("Address2") & vbCrLf
    End If
    
    rng.InsertAfter vbCrLf
    
    ' Add subject line
    rng.InsertAfter "Subject: Statement of Work for Health & Benefits Services" & vbCrLf & vbCrLf
    
    ' Add salutation
    rng.InsertAfter "Dear " & clientInfo("ContactName") & ":" & vbCrLf & vbCrLf
    
    ' Add opening paragraph
    rng.InsertAfter "This statement of work (""SOW"") will confirm the terms of the engagement of " & _
                 clientInfo("WTWParty") & " (""WTW"", ""we"" or ""us"") by " & clientInfo("ClientName") & _
                 " (""Client"" or ""you"")." & vbCrLf & vbCrLf
End Sub

' Add Terms and Conditions section
Private Sub AddTermsAndConditionsSection(rng As Range)
    rng.InsertAfter "I. Terms and Conditions of SOW:" & vbCrLf & vbCrLf
    rng.InsertAfter "Client desires to procure and WTW is willing to provide the services listed in Attachment 1 (the ""Services""). " & _
                 "These Services will be provided subject to the WTW Health & Benefits Brokerage Terms, Conditions & Disclosures available at: " & _
                 "https://www.wtwco.com/-/media/WTW/Notices/h-b-brokerage-terms-no-MSA.pdf (the ""Brokerage Terms""). " & _
                 "Copies of the Brokerage Terms are available upon request." & vbCrLf & vbCrLf
End Sub

' Add Term and Termination section
Private Sub AddTermAndTerminationSection(rng As Range, clientInfo As Dictionary, optionalClauses As Dictionary)
    rng.InsertAfter "II. Term and Termination:" & vbCrLf & vbCrLf
    rng.InsertAfter "The term of this SOW will begin on " & clientInfo("StartDate") & " and end on " & _
                 clientInfo("EndDate") & ". Either party may terminate this SOW upon 60 days prior written notice to the other party." & vbCrLf & vbCrLf
    
    ' Add auto-renewal if selected
    If optionalClauses("AutoRenewal") Then
        rng.InsertAfter "Upon the expiration of the term, or any renewal term, this SOW will renew automatically for successive one-year terms " & _
                     "unless either party gives notice of non-renewal at least 60 days before the scheduled expiration date." & vbCrLf & vbCrLf
    End If
End Sub

' Add Compensation section based on selected option
Private Sub AddCompensationSection(rng As Range, compensationOption As String, annualFee As String, _
                                  billingOption As String, policies As Collection)
    rng.InsertAfter "III. Compensation" & vbCrLf & vbCrLf
    
    Select Case compensationOption
        Case "A"
            InsertFeeOnly rng, annualFee, billingOption
        Case "B"
            InsertFeePlusCommission rng, annualFee, billingOption, policies
        Case "C"
            InsertFeeOffset rng, annualFee, billingOption, policies
        Case "D"
            InsertCommissionOnly rng, policies
    End Select
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
Private Sub InsertFeePlusCommission(rng As Range, annualFee As String, billingOption As String, policies As Collection)
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
    InsertPolicyList rng, policies
    
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
End Sub

' Insert fee offset by commission option
Private Sub InsertFeeOffset(rng As Range, annualFee As String, billingOption As String, policies As Collection)
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
    InsertPolicyList rng, policies
    
    ' Insert earned fee table
    InsertEarnedFeeTable rng
End Sub

' Insert commission-only option
Private Sub InsertCommissionOnly(rng As Range, policies As Collection)
    rng.InsertAfter "You agree that we will be compensated by commissions paid to us by insurers for the sale of the insurance policies " & _
                 "that you purchase. All commissions will be fully disclosed to you prior to our placing coverage. The commissions " & _
                 "will be earned for the entire policy period at the time we place insurance policies for you." & vbCrLf & vbCrLf
    
    ' Insert policy list if there are policies
    If policies.Count > 0 Then
        InsertPolicyList rng, policies
    Else
        rng.InsertAfter "[Policy list to be determined]" & vbCrLf & vbCrLf
    End If
    
    ' Insert commission adjustment clause
    rng.InsertAfter "The parties agree that should commissions increase or decrease by an amount exceeding ten percent (10%), " & _
                 "due to a change in covered lives, added or deleted policies or for other reasons, the parties will discuss " & _
                 "changes to our compensation and/or Services. Any such changes will be agreed to in writing by the parties." & vbCrLf & vbCrLf
    
    ' Insert additional information
    rng.InsertAfter "Information regarding other compensation we may receive is described in the Brokerage Terms." & vbCrLf & vbCrLf
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
    
    rng.InsertAfter "• reimbursement, at cost, of direct expenses reasonably incurred by us in connection with the performance " & _
                 "of our Services, such as travel and other vendor expenses, and itemized extraordinary expenses such as " & _
                 "large-volume color printing, large-volume courier shipments and the like, plus an administrative fee of 5% " & _
                 "of any vendor charges other than travel, unless arrangements are made in advance for charges to be invoiced " & _
                 "to and paid by you directly; and" & vbCrLf & vbCrLf
    
    rng.InsertAfter "• the amount of any tax or similar assessment based upon our charges." & vbCrLf & vbCrLf
    
    rng.InsertAfter "We will bill you for the fee payments as they become due. At the end of each month during which we perform " & _
                 "Services for you, we will also bill you for all other charges accrued for the month, such as travel and " & _
                 "vendor expenses." & vbCrLf & vbCrLf
End Sub

' Insert policy list from collection
Private Sub InsertPolicyList(rng As Range, policies As Collection)
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
End Sub

' Insert earned fee table for options A, B, and C
Private Sub InsertEarnedFeeTable(rng As Range)
    rng.InsertAfter "You acknowledge that, even though we may regularly invoice you on a different schedule during the term of this " & _
                 "SOW, a substantial portion of our work is provided prior to and at the effective date of your benefit plan. " & _
                 "Therefore, if this SOW is terminated before the end of the term, in order to compensate us fully for the " & _
                 "services actually provided to you, the parties agree that the fee is earned and that you will pay us as " & _
                 "provided in the following table:" & vbCrLf & vbCrLf
    
    ' Insert table (simplified for text output)
    rng.InsertAfter "Strategic Planning                15%     Earned in equal monthly installments prior to the benefit plan" & vbCrLf
    rng.InsertAfter "                                          effective date (fully earned at benefit plan effective date)" & vbCrLf & vbCrLf
    
    rng.InsertAfter "Program Renewal /                35%     Earned in equal monthly installments prior to the benefit plan" & vbCrLf
    rng.InsertAfter "Placement Process                        effective date (fully earned at benefit plan effective date)" & vbCrLf & vbCrLf
    
    rng.InsertAfter "Ongoing Service and              50%     Earned in 12 equal monthly installments (starting at benefit" & vbCrLf
    rng.InsertAfter "Resources                                plan effective date)" & vbCrLf & vbCrLf
End Sub

' Add Additional Terms section
Private Sub AddAdditionalTermsSection(rng As Range, optionalClauses As Dictionary, additionalNotes As String)
    rng.InsertAfter "IV. Additional Terms" & vbCrLf & vbCrLf
    
    ' Add GDPR clause if applicable
    If optionalClauses("GDPR") Then
        rng.InsertAfter "The parties acknowledge that WTW may access personal data from Europe that could trigger GDPR requirements. " & _
                     "For Health & Benefits Brokerage services, WTW will act as a data controller in accordance with applicable " & _
                     "data protection laws." & vbCrLf & vbCrLf
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
End Sub

' Add signature blocks
Private Sub AddSignatureBlocks(rng As Range, clientInfo As Dictionary)
    rng.InsertAfter "IN WITNESS WHEREOF, the parties have executed this SOW effective as of the _____ day of _______________, 20__." & vbCrLf & vbCrLf
    
    rng.InsertAfter "Signed by and on behalf of " & clientInfo("WTWParty") & vbCrLf & vbCrLf
    rng.InsertAfter "By: ____________________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print name: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print title: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Date: __________________________" & vbCrLf & vbCrLf
    
    rng.InsertAfter "Accepted and agreed on behalf of " & clientInfo("ClientName") & vbCrLf & vbCrLf
    rng.InsertAfter "By: ____________________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print name: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Print title: _____________________" & vbCrLf & vbCrLf
    rng.InsertAfter "Date: __________________________" & vbCrLf & vbCrLf
End Sub

' Add attachment for scope of services
Private Sub AddAttachment(rng As Range)
    rng.InsertAfter "Attachments: Attachment 1 -- Scope of Services" & vbCrLf & vbCrLf
    
    rng.InsertAfter "Attachment 1" & vbCrLf & vbCrLf
    rng.InsertAfter "Services" & vbCrLf & vbCrLf
    rng.InsertAfter "[Attach Scope of Services]" & vbCrLf
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
               InStr(para.Range.Text, "IV. Additional Terms") > 0 Or _
               InStr(para.Range.Text, "Attachment 1") > 0 Then
                para.Range.Font.Bold = True
                para.Range.Font.Underline = wdUnderlineSingle
            End If
        Next para
        
        ' Set default paragraph formatting
        .Paragraphs.Format.SpaceAfter = 6
        .Paragraphs.Format.SpaceBefore = 0
        
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
