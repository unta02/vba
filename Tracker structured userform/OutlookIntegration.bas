' OutlookIntegration.bas - Email and Outlook Interaction
'Attribute VB_Name = "OutlookIntegration"
Option Explicit

' Function to extract email content from the selected Outlook email
Public Sub ExtractEmailContent()
    On Error Resume Next
    
    ' Add confirmation dialog before extraction
    Dim confirmResponse As Integer
    confirmResponse = MsgBox("Are you sure you want to extract content from the selected email?", vbQuestion + vbYesNo, "Confirm Extraction")
    
    ' Exit if user selects No
    If confirmResponse = vbNo Then
        Exit Sub
    End If
    
    ' Clear all form fields first
    ClearFormFields
    
    ' Reset row number when extracting new content
    frmInput.lblRowNum.Caption = ""
    
    ' Reset the save button caption to "Save Record" when extracting a new email
    frmInput.cmdSave.Caption = "Save Record"
    
    ' Reference to Outlook
    Dim outlookApp As Object
    Dim explorer As Object
    Dim selection As Object
    
    ' Create Outlook Application instance if not already running
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    
    ' Get active explorer
    Set explorer = outlookApp.ActiveExplorer
    If explorer Is Nothing Then
        frmInput.txtEmailOutput.Text = "Please open Outlook and select an email."
        Exit Sub
    End If
    
    ' Check if an email is selected
    Set selection = explorer.selection
    If selection.Count = 0 Then
        frmInput.txtEmailOutput.Text = "Please select an email first."
        Exit Sub
    End If
    
    ' Get the selected email
    Dim objMail As Object
    Set objMail = selection.Item(1)
    
    If objMail.Class <> 43 Then ' 43 = olMail
        frmInput.txtEmailOutput.Text = "The selected item is not an email."
        Exit Sub
    End If
    
    ' Get both plain text and HTML content of the email
    Dim plainContent As String
    Dim htmlContent As String
    plainContent = objMail.Body
    htmlContent = objMail.htmlBody
    
    ' Find the starting position of "Received - Legal Matter"
    Dim startPos As Long
    startPos = InStr(htmlContent, "Received - Legal Matter")
    
    If startPos = 0 Then
        frmInput.txtEmailOutput.Text = "Could not find 'Received - Legal Matter' in the email content."
        frmInput.lblMsg.Text = "Could not find 'Received - Legal Matter' in the email content."
        frmInput.lblMsg.Visible = True
        Exit Sub
    End If
    
    ' Format the content for display using plain text
    frmInput.txtEmailOutput.Text = FormatPlainContent(plainContent)
    
    ' Extract specific information and populate form fields (still using HTML for reliable extraction)
    PopulateFormFields htmlContent
    
    ' Enable the Download Files button since we've successfully extracted an email
    frmInput.cmdDownloadFiles.Enabled = True
    DisplayMessage "Email content extraction successful"
    
    ' Clean up
    Set objMail = Nothing
    Set selection = Nothing
    Set explorer = Nothing
End Sub

' Function to download attachments from the selected email
Public Sub DownloadEmailAttachments()
    ' Check if an email has been extracted
    If Trim(frmInput.txtEmailOutput.Text) = "" Then
        DisplayMessage "Please extract an email first.", True
        Exit Sub
    End If
    
    ' Ask for confirmation before downloading attachments
    Dim response As Integer
    response = MsgBox("Are you sure you want to download document attachments from the selected email?", _
                      vbQuestion + vbYesNo, "Confirm Download")
    
    If response = vbNo Then
        Exit Sub
    End If
    
    ' Reference to Outlook
    Dim outlookApp As Object
    Dim explorer As Object
    Dim selection As Object
    
    ' Create Outlook Application instance if not already running
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    
    ' Get active explorer
    Set explorer = outlookApp.ActiveExplorer
    If explorer Is Nothing Then
        DisplayMessage "Please open Outlook and select an email.", True
        Exit Sub
    End If
    
    ' Check if an email is selected
    Set selection = explorer.selection
    If selection.Count = 0 Then
        DisplayMessage "Please select an email first.", True
        Exit Sub
    End If
    
    ' Get the selected email
    Dim objMail As Object
    Set objMail = selection.Item(1)
    
    If objMail.Class <> 43 Then ' 43 = olMail
        DisplayMessage "The selected item is not an email.", True
        Exit Sub
    End If
    
    ' Verify this is the same email that was extracted
    Dim currentHtmlContent As String
    Dim currentLMNumber As String
    Dim originalLMNumber As String
    
    currentHtmlContent = objMail.htmlBody
    currentLMNumber = ExtractLMNumber(currentHtmlContent)
    originalLMNumber = frmInput.txtLM.Text
    
    ' Check if LM numbers match
    If currentLMNumber <> originalLMNumber Then
        DisplayMessage "The currently selected email does not match the one initially extracted." & vbCrLf & _
               "Expected LM number: " & originalLMNumber & vbCrLf & _
               "Current LM number: " & currentLMNumber & vbCrLf & vbCrLf & _
               "Please select the original email with Legal Matter " & originalLMNumber & ".", _
               True
        Exit Sub
    End If
    
    ' Count document attachments (PDFs and Word documents only)
    Dim docCount As Integer
    Dim attachment As Object
    docCount = 0
    
    For Each attachment In objMail.Attachments
        If IsDocumentFile(attachment.fileName) Then
            docCount = docCount + 1
        End If
    Next attachment
    
    ' Check if the email has any document attachments
    If docCount = 0 Then
        DisplayMessage "This email does not have any document attachments (PDF or Word documents).", True
        Exit Sub
    End If
    
    ' Ask the user for a folder to save the attachments
    Dim folderDialog As Object
    Set folderDialog = CreateObject("Shell.Application").BrowseForFolder(0, "Select folder to save attachments:", 0, 0)
    
    ' Check if user selected a folder
    If folderDialog Is Nothing Then
        DisplayMessage "Download canceled. No folder was selected.", True
        Exit Sub
    End If
    
    ' Get the selected folder path
    Dim folderPath As String
    folderPath = folderDialog.Self.Path
    
    ' Download document attachments only
    Dim saveCount As Integer
    Dim skipCount As Integer
    Dim errorCount As Integer
    Dim filePath As String
    
    saveCount = 0
    skipCount = 0
    errorCount = 0
    
    On Error Resume Next
    For Each attachment In objMail.Attachments
        ' Check if this is a document file type we want to download
        If IsDocumentFile(attachment.fileName) Then
            ' Create a safe file name (replace invalid characters)
            Dim safeFileName As String
            safeFileName = Replace(attachment.fileName, "/", "_")
            safeFileName = Replace(safeFileName, "\", "_")
            safeFileName = Replace(safeFileName, ":", "_")
            safeFileName = Replace(safeFileName, "*", "_")
            safeFileName = Replace(safeFileName, "?", "_")
            safeFileName = Replace(safeFileName, """", "_")
            safeFileName = Replace(safeFileName, "<", "_")
            safeFileName = Replace(safeFileName, ">", "_")
            safeFileName = Replace(safeFileName, "|", "_")
            
            ' Create the full file path
            filePath = folderPath & "\" & safeFileName
            
            ' Check if file already exists and modify name if needed
            Dim counter As Integer
            counter = 1
            Dim baseName As String
            Dim fileExt As String
            
            ' Split filename into base name and extension
            If InStr(safeFileName, ".") > 0 Then
                baseName = Left(safeFileName, InStrRev(safeFileName, ".") - 1)
                fileExt = mid(safeFileName, InStrRev(safeFileName, "."))
            Else
                baseName = safeFileName
                fileExt = ""
            End If
            
            ' If file exists, add a counter to the filename
            While Dir(filePath) <> ""
                filePath = folderPath & "\" & baseName & "(" & counter & ")" & fileExt
                counter = counter + 1
            Wend
            
            ' Save the attachment
            attachment.SaveAsFile filePath
            
            ' Check for errors
            If Err.Number = 0 Then
                saveCount = saveCount + 1
            Else
                errorCount = errorCount + 1
                Err.Clear
            End If
        Else
            ' Skip this attachment since it's not a document type
            skipCount = skipCount + 1
        End If
    Next attachment
    
    On Error GoTo 0
    
    ' Show summary message
    Dim summaryMsg As String
    summaryMsg = "Download complete." & vbCrLf & vbCrLf & _
                 "Successfully saved: " & saveCount & " document(s)" & vbCrLf
    
    If skipCount > 0 Then
        summaryMsg = summaryMsg & "Skipped: " & skipCount & " attachment(s) (not document files)" & vbCrLf
    End If
    
    If errorCount > 0 Then
        summaryMsg = summaryMsg & "Failed: " & errorCount & " attachment(s) (errors occurred)" & vbCrLf
    End If
    
    DisplayMessage summaryMsg, True
End Sub

' Function to send email using the extracted information
Public Sub SendEmail()
    ' Check if there are unsaved changes and prompt user to save first
    If HasChanges Then
        Dim saveFirst As Integer
        saveFirst = MsgBox("You have unsaved changes. Would you like to save them before sending the email?", _
                          vbQuestion + vbYesNoCancel, "Unsaved Changes")
        
        ' Handle user's choice
        Select Case saveFirst
            Case vbYes
                ' Call SaveFormData and check if it was successful
                SaveFormData
                ' If there are still changes, it means saving failed - abort sending email
                If HasChanges Then
                    Exit Sub
                End If
            Case vbCancel
                ' User cancelled, abort sending email
                Exit Sub
            Case vbNo
                ' Continue without saving
        End Select
    End If

    ' First validate RCL email format if it's Contract Review
    If frmInput.cboRequest.Text = "Contract Review" Then
        If Trim(frmInput.txtRCLemail.Text) <> "" And Not IsValidEmail(Trim(frmInput.txtRCLemail.Text)) Then
            DisplayMessage "The RCL Email address format is invalid. Please enter a valid email address.", True
            frmInput.txtRCLemail.SetFocus
            Exit Sub
        End If
    End If
    
    ' Validate required fields next
    If Not ValidateRequiredFieldsForEmail Then
        Exit Sub
    End If
    
    ' Reference to Outlook
    Dim outlookApp As Object
    Dim explorer As Object
    Dim selection As Object
    Dim originalEmail As Object
    Dim replyEmail As Object
    Dim htmlBody As String
    Dim emailSubject As String
    Dim toRecipient As String
    Dim ccRecipient As String
    Dim requestedForName As String
    Dim contractManagerFullName As String
    Dim contractManagerShortName As String
    Dim assignedRCL As String
    Dim contractRequestType As String
    Dim emailTemplate As String
    
    ' Get the contract request type
    contractRequestType = frmInput.cboRequest.Text
    
    ' Get the email template type
    emailTemplate = frmInput.cboEmailType.Text
    
    ' Create Outlook Application instance if not already running
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    
    ' Get active explorer and selection
    Set explorer = outlookApp.ActiveExplorer
    If explorer Is Nothing Then
        DisplayMessage "Please open Outlook and select an email.", True
        Exit Sub
    End If
    
    Set selection = explorer.selection
    If selection.Count = 0 Then
        DisplayMessage "Please select an email first.", True
        Exit Sub
    End If
    
    ' Get the selected email
    Set originalEmail = selection.Item(1)
    If originalEmail.Class <> 43 Then ' 43 = olMail
        DisplayMessage "The selected item is not an email.", True
        Exit Sub
    End If
    
    ' Get HTML content of the currently selected email and check if it matches the originally extracted email
    Dim currentHtmlContent As String
    Dim currentLMNumber As String
    Dim originalLMNumber As String
    
    currentHtmlContent = originalEmail.htmlBody
    currentLMNumber = ExtractLMNumber(currentHtmlContent)
    originalLMNumber = frmInput.txtLM.Text
    
    ' Verify that the currently selected email matches the one used for extraction
    If currentLMNumber <> originalLMNumber Then
        DisplayMessage "The currently selected email does not match the one initially extracted." & vbCrLf & _
               "Expected LM number: " & originalLMNumber & vbCrLf & _
               "Current LM number: " & currentLMNumber & vbCrLf & vbCrLf & _
               "Please select the original email with Legal Matter " & originalLMNumber & " or re-extract data from the current email.", _
               True
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' Extract required information
    ' 1. Get the Requested For name from the HTML content
    requestedForName = ExtractRequestedForName(frmInput.txtEmailOutput.Text)
    
    ' 2. Get the Contract Manager information from Admin sheet
    contractManagerFullName = GetContractManagerFullName(frmInput.cboCM.Text)
    contractManagerShortName = frmInput.cboCM.Text
    
    ' 3. Get the Assigned RCL (only if not Contract Upload)
    assignedRCL = ""
    If contractRequestType <> "Contract Upload" Then
        assignedRCL = frmInput.cboRCL.Text
    End If
    
    ' 4. Get email subject
    emailSubject = frmInput.txtEmailSubject.Text
    
    ' 5. Get recipient email address from the HTML content
    toRecipient = ExtractContactEmail(frmInput.txtEmailOutput.Text)
    
    ' 6. Get CC recipient (Contract Manager's email)
    ccRecipient = GetContractManagerEmail(frmInput.cboCM.Text)
    
    ' Set team email address
    Dim teamEmailAddress As String
    teamEmailAddress = "Contracting.COE@wtwco.com"
    
    ' Determine email content based on contract request type
    If contractRequestType = "Contract Upload" Then
        ' Simple message for Contract Upload
        htmlBody = "<!DOCTYPE html><html><body style=""font-family:Arial;font-size:11pt;"">" & _
                  "<p>Hello " & contractManagerShortName & ",</p>" & _
                  "<p>Please upload the contract to COMET.</p>" & _
                  "<p>Thanks,<br>Contracting Management</p>" & _
                  "</body></html>"
    ElseIf contractRequestType = "Out of Scope" Then
        ' Out of Scope email with specified format
        htmlBody = "<!DOCTYPE html><html><body style=""font-family:Arial;font-size:11pt;"">" & _
                  "<p>Hello " & requestedForName & ",</p>" & _
                  "<p>Thank you for submitting your Contract Support Request. Please be advised however that this type of request is out of scope of the Client Contract Management Team. </p>" & _
                  "<p>Thank you,<br>Sales Operation | Contract and Proposal Centre of Excellence<br>Client Contract Management</p>" & _
                  "</body></html>"
        

    ElseIf contractRequestType = "Duplicate Request" Then
        ' Simple message for Duplicate Request
        htmlBody = "<!DOCTYPE html><html><body style=""font-family:Arial;font-size:11pt;"">" & _
                  "<p>Hello " & requestedForName & ",</p>" & _
                  "<p>This has been identified as a Duplicate Request. Please see the original request.</p>" & _
                   "<p>Thank you,<br>Sales Operation | Contract and Proposal Centre of Excellence<br>Client Contract Management</p>" & _
                  "</body></html>"
    ElseIf contractRequestType = "Contract Review" Then
        ' Use appropriate template based on selection
        If emailTemplate = "RFP" Then
            htmlBody = LoadRFPEmailTemplate()
            
            ' Replace placeholders with actual data for RFP template
            htmlBody = Replace(htmlBody, "<<Requested For: Name>>", requestedForName)
            htmlBody = Replace(htmlBody, "<<Contract Manager Full Name>>", contractManagerFullName)
            htmlBody = Replace(htmlBody, "<<Contract Manager Short Name>>", contractManagerShortName)
            htmlBody = Replace(htmlBody, "<<Assigned RCL cboRCL>>", assignedRCL)
            htmlBody = Replace(htmlBody, "<<Client or Supplier Name>>", frmInput.txtClient.Text)
        ElseIf emailTemplate = "Standard Urgent" Then
            ' Use Urgent template
            htmlBody = LoadUrgentEmailTemplate()
            
            ' Replace placeholders with actual data
            htmlBody = Replace(htmlBody, "<<Requested For: Name>>", requestedForName)
            htmlBody = Replace(htmlBody, "<<Contract Manager Full Name>>", contractManagerFullName)
            htmlBody = Replace(htmlBody, "<<Contract Manager Short Name>>", contractManagerShortName)
            htmlBody = Replace(htmlBody, "<<Assigned RCL cboRCL>>", assignedRCL)
        Else
            ' Standard template
            htmlBody = LoadEmailTemplate()
            
            ' Replace placeholders with actual data for standard template
            htmlBody = Replace(htmlBody, "<<Requested For: Name>>", requestedForName)
            htmlBody = Replace(htmlBody, "<<Contract Manager Full Name>>", contractManagerFullName)
            htmlBody = Replace(htmlBody, "<<Contract Manager Short Name>>", contractManagerShortName)
            htmlBody = Replace(htmlBody, "<<Assigned RCL cboRCL>>", assignedRCL)
        End If
    Else
        ' Use standard template for all other cases
        htmlBody = LoadEmailTemplate()
        
        ' Replace placeholders with actual data
        htmlBody = Replace(htmlBody, "<<Requested For: Name>>", requestedForName)
        htmlBody = Replace(htmlBody, "<<Contract Manager Full Name>>", contractManagerFullName)
        htmlBody = Replace(htmlBody, "<<Contract Manager Short Name>>", contractManagerShortName)
        htmlBody = Replace(htmlBody, "<<Assigned RCL cboRCL>>", assignedRCL)
    End If
    
    ' Ask user if they want to create a new email or forward the original
    Dim response As Integer
    Dim promptMsg As String
    
    If contractRequestType = "Contract Upload" Then
        promptMsg = "You're about to send a Contract Upload email. Do you want to reply to the original email with this template?"
    ElseIf contractRequestType = "Out of Scope" Then
        promptMsg = "You're about to send an Out of Scope email. Do you want to reply to the original email with this template?"
    ElseIf contractRequestType = "Duplicate Request" Then
        promptMsg = "You're about to send a Duplicate Request email. Do you want to reply to the original email with this template?"
    ElseIf contractRequestType = "Contract Review" Then
        promptMsg = "You're about to send a Contract Review email using the " & emailTemplate & " template. Do you want to reply to the original email with this template?"
    Else
        promptMsg = "Do you want to reply to the original email with the template?"
    End If
    
    response = MsgBox(promptMsg, vbQuestion + vbYesNo, "Email Options")
    
    ' If No is selected, exit the process
    If response = vbNo Then
        Exit Sub
    End If
    
    ' Create appropriate email type based on request type
    If contractRequestType = "Contract Upload" Then
        Set replyEmail = originalEmail.Forward
    Else
        Set replyEmail = originalEmail.Reply
    End If
    
    ' Set mail properties for replied email
    With replyEmail
        .Subject = emailSubject
        .htmlBody = htmlBody & "<hr>" & .htmlBody
        
        ' Set the sender to the team mailbox
        .SentOnBehalfOfName = teamEmailAddress
        
        ' Add recipients based on contract request type
        If contractRequestType = "Contract Upload" Then
            ' For Contract Upload, only send to Contract Manager
            If ccRecipient <> "" Then
                .To = ccRecipient
            End If
        ElseIf contractRequestType = "Out of Scope" Then
            ' For Out of Scope, set specific recipients based on the template
            If toRecipient <> "" Then
                .To = toRecipient
            End If
            
            If ccRecipient <> "" Then
                .CC = ccRecipient
            End If
            .CC = teamEmailAddress
        ElseIf contractRequestType = "Duplicate Request" Then
            If toRecipient <> "" Then
                .To = toRecipient
            End If
            .CC = teamEmailAddress
            ' Do nothing here - don't set any recipients
        ElseIf contractRequestType = "Contract Review" Then
            ' For Contract Review, include RCL email in CC
            If toRecipient <> "" Then
                .To = toRecipient
            End If
            
            ' Combine Contract Manager email and RCL email in CC field
            Dim combinedCC As String
            combinedCC = ""
            
            If ccRecipient <> "" Then
                combinedCC = ccRecipient
            End If
            
            If Trim(frmInput.txtRCLemail.Text) <> "" Then
                If combinedCC <> "" Then
                    combinedCC = combinedCC & ";" & frmInput.txtRCLemail.Text
                Else
                    combinedCC = frmInput.txtRCLemail.Text
                End If
            End If
            
            If combinedCC <> "" Then
                .CC = combinedCC
            End If
        Else
            ' For other types, use both to and cc recipients
            If toRecipient <> "" Then
                .To = toRecipient
            End If
            
            If ccRecipient <> "" Then
                .CC = ccRecipient
            End If
        End If
        
        ' Display the email without sending
        .Display
    End With
    
    DisplayMessage "Reply email has been composed with your template.", True
    
    ' Reset the change status after successfully sending the email
    ResetChangeStatus
    
    ' Clean up
    Set replyEmail = Nothing
    Set originalEmail = Nothing
    Set selection = Nothing
    Set explorer = Nothing
    Set outlookApp = Nothing
End Sub 