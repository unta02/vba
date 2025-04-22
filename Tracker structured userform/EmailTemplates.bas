' EmailTemplates.bas - Email Templates Module

Option Explicit

' Function to load the standard email template
Public Function LoadEmailTemplate() As String
    ' This version splits the HTML into multiple smaller sections to avoid line continuation limits
    Dim html1 As String, html2 As String, html3 As String
    Dim html2_1 As String
    
    ' Part 1: Basic opening and intro text
    html1 = "<!DOCTYPE html><html><body style=""font-family:Arial;font-size:11pt;"">"
    html1 = html1 & "<p>Hello <<Requested For: Name>>,</p>"
    html1 = html1 & "<p>Thank you for submitting your Contract Support Request. <b><<Contract Manager Full Name>></b> will be your Contract Specialist for this request.</p>"
    html1 = html1 & "<p>Your contract review is currently with your project manager who is completing a high-level review, prior to us sending it out to the necessary reviewers.</p>"
    html1 = html1 & "<p><b><<Contract Manager Short Name>></b> will be keeping you updated on the status as we move through the review process. If you have any questions, please contact the Contract Specialist for this request.</p>"
    
    ' Code of conduct section
    html2_1 = "<p>However, we saw that you requested a review of a <b>supplier code of conduct</b>.</p>"
    html2_1 = html2_1 & "<p>We typically do not review/sign client's COC due to compliance and other reasons. To do so would be administratively unworkable across our entire client base as each client's policies are different (and potentially inconsistent) and we need to manage our infrastructure, and train our associates, to one set of policies firm-wide. We can provide a copy of the WTW code of conduct and can provide additional detail upon request.</p>"
    html2_1 = html2_1 & "<p>However, of course, if we know that will not work and will harm our ability to win new business or keep old, we will sometimes consider it on an exception basis. If the code is kept high level and consistent with our own (as many are) then we may take a quick look and if the business agrees we can comply and based on our review there is not something objectionable (sometimes there could be something specific or beyond our code on social or environmental issues, for example), then we may take a risk based approach and agree.</p>"
    html2_1 = html2_1 & "<p>Please let us know</p>"
    
    ' Part 2: Middle section with details
    html2 = "<p><b>Please provide</b> any applications that are in scope for this project, if not already provided in the intake form. If an ICS review is needed, this is required due to how ICS assigns out. Their SLA will not start until the applications are provided to their team.</p>"
    html2 = html2 & "<p><b>Your assigned legal contact</b> at this point in time is <<Assigned RCL cboRCL>>.</p>"
    html2 = html2 & "<p>As a reminder, please make sure to include your project manager in <b>any and all internal</b> communications with pertinent stakeholders regarding this contract request. This will ensure that the project manager is able to provide the proper level of support and keep your reviews moving forward.</p>"
    html2 = html2 & "<p>Thank you,<br>Sales Operation | Contract and Proposal Centre of Excellence<br>Contracting Management</p>"
    
    ' Part 3: Table section with process steps
    html3 = "<table style=""width:100%;background-color:#D9E2F3;border:1px solid #ccc;"">"
    html3 = html3 & "<tr><td style=""padding:10px;"">"
    html3 = html3 & "<p><b><u>Process Steps</u></b></p>"
    html3 = html3 & "<p style=""font-style:italic;"">*Estimated time from receiving this email to receiving the 1<sup>st</sup> redline draft to be sent to the client will vary depending on complexity of the contract and the number of SMEs involved. If only legal review is required, then the process time will be approximately 5 to 10 business days or less depending on the nature of the request. If other SME reviews are required, the process time can take up to 15 business days or more in some extreme cases.</p>"
    html3 = html3 & "</td></tr><tr><td style=""padding:10px;"">"
    html3 = html3 & "<ul>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to do a high-level review</li>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to send and manage identified SMEs reviews (ICS, Privacy, HR, Insurance...etc.)</li>"
    html3 = html3 & "<li style=""font-style:italic;"">SME reviews completed – Sent to Legal for full Legal Review</li>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to consolidate and clean up document/s</li>"
    html3 = html3 & "<li style=""font-style:italic;"">If needed - Requestor and necessary members of the team to meet internally to discuss redlines</li>"
    html3 = html3 & "<li style=""font-style:italic;"">Requestor to send to client for review</li>"
    html3 = html3 & "</ul>"
    html3 = html3 & "<p style=""font-style:italic;"">**If client has added redlines and/or comments, including exceptions to WTW standard (modified language or additional provisions), additional involvement with Legal and/or SMEs will be needed after the first round of WTW redlines</p>"
    html3 = html3 & "<p style=""font-style:italic;"">Due to capacity issues within the current Privacy team, documents with privacy language are initially reviewed by the assigned Legal Counsel, if they're unable to approve, they will send directly to the Privacy team for review. This is likely to increase processing time</p>"
    html3 = html3 & "</td></tr></table></body></html>"
    
    ' Combine all parts
    LoadEmailTemplate = html1 & html2_1 & html2 & html3
End Function

' Function to load the Standard Urgent email template - includes the urgency line
Public Function LoadUrgentEmailTemplate() As String
    ' This version splits the HTML into multiple smaller sections to avoid line continuation limits
    Dim html1 As String, html2 As String, html3 As String
    Dim html2_1 As String
    
    ' Part 1: Basic opening and intro text
    html1 = "<!DOCTYPE html><html><body style=""font-family:Arial;font-size:11pt;"">"
    html1 = html1 & "<p>Hello <<Requested For: Name>>,</p>"
    html1 = html1 & "<p>Thank you for submitting your Contract Support Request. <b><<Contract Manager Full Name>></b> will be your Contract Specialist for this request.</p>"
    html1 = html1 & "<p>Your contract review is currently with your project manager who is completing a high-level review, prior to us sending it out to the necessary reviewers.</p>"
    html1 = html1 & "<p><b><<Contract Manager Short Name>></b> will be keeping you updated on the status as we move through the review process. If you have any questions, please contact the Contract Specialist for this request.</p>"
    
    ' Code of conduct section
    html2_1 = "<p>However, we saw that you requested a review of a <b>supplier code of conduct</b>.</p>"
    html2_1 = html2_1 & "<p>We typically do not review/sign client's COC due to compliance and other reasons. To do so would be administratively unworkable across our entire client base as each client's policies are different (and potentially inconsistent) and we need to manage our infrastructure, and train our associates, to one set of policies firm-wide. We can provide a copy of the WTW code of conduct and can provide additional detail upon request.</p>"
    html2_1 = html2_1 & "<p>However, of course, if we know that will not work and will harm our ability to win new business or keep old, we will sometimes consider it on an exception basis. If the code is kept high level and consistent with our own (as many are) then we may take a quick look and if the business agrees we can comply and based on our review there is not something objectionable (sometimes there could be something specific or beyond our code on social or environmental issues, for example), then we may take a risk based approach and agree.</p>"
    html2_1 = html2_1 & "<p>Please let us know</p>"
    
    ' Part 2: Middle section with details - including the urgency line
    html2 = "<p>We acknowledge the urgency of the matter. However, please be advised that, given the current workload, it may not be feasible to accommodate the requested deadline. Kindly ensure that client expectations are managed correctly. However, if there is a business reason for the urgency, please provide and we will do our best to meet the requested deadline.</p>"
    html2 = html2 & "<p><b>Please provide</b> any applications that are in scope for this project, if not already provided in the intake form. If an ICS review is needed, this is required due to how ICS assigns out. Their SLA will not start until the applications are provided to their team.</p>"
    html2 = html2 & "<p><b>Your assigned legal contact</b> at this point in time is <<Assigned RCL cboRCL>>.</p>"
    html2 = html2 & "<p>As a reminder, please make sure to include your project manager in <b>any and all internal</b> communications with pertinent stakeholders regarding this contract request. This will ensure that the project manager is able to provide the proper level of support and keep your reviews moving forward.</p>"
    html2 = html2 & "<p>Thank you,<br>Sales Operation | Contract and Proposal Centre of Excellence<br>Contracting Management</p>"
    
    ' Part 3: Table section with process steps
    html3 = "<table style=""width:100%;background-color:#D9E2F3;border:1px solid #ccc;"">"
    html3 = html3 & "<tr><td style=""padding:10px;"">"
    html3 = html3 & "<p><b><u>Process Steps</u></b></p>"
    html3 = html3 & "<p style=""font-style:italic;"">*Estimated time from receiving this email to receiving the 1<sup>st</sup> redline draft to be sent to the client will vary depending on complexity of the contract and the number of SMEs involved. If only legal review is required, then the process time will be approximately 5 to 10 business days or less depending on the nature of the request. If other SME reviews are required, the process time can take up to 15 business days or more in some extreme cases.</p>"
    html3 = html3 & "</td></tr><tr><td style=""padding:10px;"">"
    html3 = html3 & "<ul>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to do a high-level review</li>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to send and manage identified SMEs reviews (ICS, Privacy, HR, Insurance...etc.)</li>"
    html3 = html3 & "<li style=""font-style:italic;"">SME reviews completed – Sent to Legal for full Legal Review</li>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to consolidate and clean up document/s</li>"
    html3 = html3 & "<li style=""font-style:italic;"">If needed - Requestor and necessary members of the team to meet internally to discuss redlines</li>"
    html3 = html3 & "<li style=""font-style:italic;"">Requestor to send to client for review</li>"
    html3 = html3 & "</ul>"
    html3 = html3 & "<p style=""font-style:italic;"">**If client has added redlines and/or comments, including exceptions to WTW standard (modified language or additional provisions), additional involvement with Legal and/or SMEs will be needed after the first round of WTW redlines</p>"
    html3 = html3 & "</td></tr></table></body></html>"
    
    ' Combine all parts
    LoadUrgentEmailTemplate = html1 & html2_1 & html2 & html3
End Function

' Function to load the RFP email template
Public Function LoadRFPEmailTemplate() As String
    ' This version splits the HTML into multiple smaller sections to avoid line continuation limits
    Dim html1 As String, html2 As String, html3 As String
    
    ' Part 1: Basic opening and intro text
    html1 = "<!DOCTYPE html><html><body style=""font-family:Arial;font-size:11pt;"">"
    html1 = html1 & "<p>Hello <<Requested For: Name>>,</p>"
    html1 = html1 & "<p>Thank you for submitting the Contract Support Request. <b><<Contract Manager Full Name>></b> will be your Contract Specialist for this request.</p>"
    html1 = html1 & "<p>Your contract review is currently with your contract manager who is completing a high-level review, prior to us sending it out to the necessary reviewers.</p>"
    
    ' Part 2: Middle section with RFP-specific details
    html2 = "<p>It is standard procedure for ICS and Legal to not redline contracts that come in with RFPs.</p>"
    html2 = html2 & "<p>However, if we are currently doing business with <b><<Client or Supplier Name>></b>, and have an active signed MSA in place, some clients will say that there is no need to do a review. <<Assigned RCL cboRCL>> is the legal resource for this request, and can be the decider on how to proceed, to redline or not.</p>"
    html2 = html2 & "<p>But of course, if need be, we will do a review, as we do not want to lose business.</p>"
    html2 = html2 & "<p>However, we saw that you requested a review of a <b>supplier code of conduct</b>.</p>"
    html2 = html2 & "<p>We typically do not review/sign client's COC due to compliance and other reasons. To do so would be administratively unworkable across our entire client base as each client's policies are different (and potentially inconsistent) and we need to manage our infrastructure, and train our associates, to one set of policies firm-wide. We can provide a copy of the WTW code of conduct and can provide additional detail upon request.</p>"
    html2 = html2 & "<p>However, of course, if we know that will not work and will harm our ability to win new business or keep old, we will sometimes consider it on an exception basis. If the code is kept high level and consistent with our own (as many are) then we may take a quick look and if the business agrees we can comply and based on our review there is not something objectionable (sometimes there could be something specific or beyond our code on social or environmental issues, for example), then we may take a risk based approach and agree.</p>"
    html2 = html2 & "<p>Please let us know</p>"
    html2 = html2 & "<p><b><u>If not already provided</u></b><b>, please provide the client RFP requirements i.e. instruction documentation. The review is on hold until we receive this.</b></p>"
    html2 = html2 & "<p><b><<Contract Manager Short Name>></b> will be keeping you updated on the status as we move though the review process. If you have any questions, please contact the Contract Specialist for this request.</p>"
    html2 = html2 & "<p>Thank you,<br>Sales Operation | Contract and Proposal Centre of Excellence<br>Contracting Management</p>"
    
    ' Part 3: Table section with process steps
    html3 = "<table style=""width:100%;background-color:#D9E2F3;border:1px solid #ccc;"">"
    html3 = html3 & "<tr><td style=""padding:10px;"">"
    html3 = html3 & "<p><b><u>Process Steps</u></b></p>"
    html3 = html3 & "<p style=""font-style:italic;"">*Estimated time from receiving this email to receiving the 1<sup>st</sup> redline draft to be sent to the client will vary depending on complexity of the contract and the number of SMEs involved. If only legal review is required, then the process time will be approximately 5 to 10 business days or less depending on the nature of the request. If other SME reviews are required, the process time can take up to 15 business days.</p>"
    html3 = html3 & "</td></tr><tr><td style=""padding:10px;"">"
    html3 = html3 & "<ul>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to do a high-level review</li>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to send and manage identified SMEs reviews (ICS, Privacy, HR, Insurance...etc.)</li>"
    html3 = html3 & "<li style=""font-style:italic;"">SME reviews completed – Sent to Legal for full Legal Review</li>"
    html3 = html3 & "<li style=""font-style:italic;"">PM to consolidate and clean up document/s</li>"
    html3 = html3 & "<li style=""font-style:italic;"">If needed - Requestor and necessary members of the team to meet internally to discuss redlines</li>"
    html3 = html3 & "<li style=""font-style:italic;"">Requestor to send to client for review</li>"
    html3 = html3 & "</ul>"
    html3 = html3 & "<p style=""font-style:italic;"">**If client has added redlines and/or comments, including exceptions to WTW standard (modified language or additional provisions), additional involvement with Legal and/or SMEs will be needed after the first round of WTW redlines</p>"
    html3 = html3 & "</td></tr></table></body></html>"
    
    ' Combine all parts
    LoadRFPEmailTemplate = html1 & html2 & html3
End Function

' Function to format email content from HTML content
Public Function FormatEmailContent(htmlContent As String) As String
    Dim formattedContent As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempStr As String
    
    formattedContent = ""
    
    ' Extract and format Legal Matter Number
    startPos = InStr(htmlContent, "Received - Legal Matter")
    If startPos > 0 Then
        startPos = startPos + Len("Received - Legal Matter")
        endPos = InStr(startPos, htmlContent, "</strong>")
        If endPos > 0 Then
            tempStr = Trim(mid(htmlContent, startPos, endPos - startPos))
            formattedContent = formattedContent & "Legal Matter Number: " & tempStr & vbCrLf & vbCrLf
        End If
    End If
    
    ' Format Client Name
    startPos = InStr(htmlContent, "<td><b>Client or Supplier Name (full legal entity name if known)</b></td>" & vbCrLf & "<td>")
    If startPos > 0 Then
        startPos = startPos + Len("<td><b>Client or Supplier Name (full legal entity name if known)</b></td>" & vbCrLf & "<td>")
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            tempStr = mid(htmlContent, startPos, endPos - startPos)
            formattedContent = formattedContent & "Client Name: " & tempStr & vbCrLf
        End If
    End If
    
    ' Format Request Type
    startPos = InStr(htmlContent, "<td><b>Request Type</b></td>" & vbCrLf & "<td>")
    If startPos > 0 Then
        startPos = startPos + Len("<td><b>Request Type</b></td>" & vbCrLf & "<td>")
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            tempStr = mid(htmlContent, startPos, endPos - startPos)
            formattedContent = formattedContent & "Request Type: " & tempStr & vbCrLf
        End If
    End If
    
    ' Format Document Type
    startPos = InStr(htmlContent, "<td><b>Document Type being requested</b></td>" & vbCrLf & "<td>")
    If startPos > 0 Then
        startPos = startPos + Len("<td><b>Document Type being requested</b></td>" & vbCrLf & "<td>")
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            tempStr = mid(htmlContent, startPos, endPos - startPos)
            formattedContent = formattedContent & "Document Type: " & tempStr & vbCrLf
        End If
    End If
    
    ' Format Line of Business
    startPos = InStr(htmlContent, "<td><b>Line of Business</b></td>" & vbCrLf & "<td>")
    If startPos > 0 Then
        startPos = startPos + Len("<td><b>Line of Business</b></td>" & vbCrLf & "<td>")
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            tempStr = mid(htmlContent, startPos, endPos - startPos)
            formattedContent = formattedContent & "Line of Business: " & tempStr & vbCrLf
        End If
    End If
    
    ' Format Region
    startPos = InStr(htmlContent, "<td><b>Client / Counterparty Location</b></td>" & vbCrLf & "<td>")
    If startPos > 0 Then
        startPos = startPos + Len("<td><b>Client / Counterparty Location</b></td>" & vbCrLf & "<td>")
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            tempStr = mid(htmlContent, startPos, endPos - startPos)
            formattedContent = formattedContent & "Region: " & tempStr & vbCrLf
        End If
    End If
    
    ' Format Contract Value
    startPos = InStr(htmlContent, "<td><b>Contract Value</b></td>" & vbCrLf & "<td>")
    If startPos > 0 Then
        startPos = startPos + Len("<td><b>Contract Value</b></td>" & vbCrLf & "<td>")
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            tempStr = mid(htmlContent, startPos, endPos - startPos)
            formattedContent = formattedContent & "Contract Value: " & tempStr & vbCrLf
        End If
    End If
    
    ' Format Due Date
    startPos = InStr(htmlContent, "<td><b>Due date</b></td>" & vbCrLf & "<td>")
    If startPos > 0 Then
        startPos = startPos + Len("<td><b>Due date</b></td>" & vbCrLf & "<td>")
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            tempStr = mid(htmlContent, startPos, endPos - startPos)
            formattedContent = formattedContent & "Due Date: " & tempStr & vbCrLf
        End If
    End If
    
    FormatEmailContent = formattedContent
End Function
