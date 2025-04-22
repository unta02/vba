' FormOperations.bas - Form Operations Module

Option Explicit

' Global variable to track if changes have been made to the form - moved from UserForm
Public HasChanges As Boolean

' Function to populate the form fields with data from HTML content
Public Sub PopulateFormFields(htmlContent As String)
    ' Set Coordinator to current user with lookup
    Dim currentUser As String
    Dim i As Integer
    currentUser = Environ("USERNAME")
    frmInput.txtCoordinator.Text = LookupCoordinator(currentUser)
    
    ' Set Date to current UTC date using GetUTCTime
    frmInput.txtDate.Text = Format(Module1.GetUTCTime(), "yyyy-mm-dd")
    
    ' Extract LM number
    Dim lmNumber As String
    lmNumber = ExtractLMNumber(htmlContent)
    frmInput.txtLM.Text = lmNumber
    
    ' Extract Client Name
    Dim clientName As String
    clientName = ExtractClientName(htmlContent)
    frmInput.txtClient.Text = clientName
    
    ' Extract Request Type
    Dim requestType As String
    requestType = ExtractRequestType(htmlContent)
    
    ' Set the Request Type in the combo box, validate it's in the list
    Dim validRequest As Boolean
    validRequest = False
    For i = 0 To frmInput.cboRequest.ListCount - 1
        If frmInput.cboRequest.List(i) = requestType Then
            frmInput.cboRequest.Text = requestType
            validRequest = True
            Exit For
        End If
    Next i
    
    ' If not a valid request type, set ListIndex to -1 (no selection)
    If Not validRequest Then
        frmInput.cboRequest.ListIndex = -1
    End If
    
    ' Check if the request type is "Contract Upload" and set Contract Type to "Upload"
    If frmInput.cboRequest.Text = "Contract Upload" Then
        frmInput.txtContractType.Text = "Upload"
    Else
        ' For other request types, extract Contract Type from email
        Dim contractType As String
        contractType = ExtractContractType(htmlContent)
        frmInput.txtContractType.Text = contractType
    End If
    
    ' Extract LOB
    Dim lob As String
    lob = ExtractLOB(htmlContent)
    frmInput.txtLOB.Text = lob
    
    ' Extract Region
    Dim region As String
    region = ExtractRegion(htmlContent)
    
    ' Set the Region in the combo box, validate it's in the list
    Dim validRegion As Boolean
    validRegion = False
    For i = 0 To frmInput.cboRegion.ListCount - 1
        If frmInput.cboRegion.List(i) = region Then
            frmInput.cboRegion.Text = region
            validRegion = True
            Exit For
        End If
    Next i
    
    ' If not a valid region, set to "Not Applicable"
    If Not validRegion Then
        frmInput.cboRegion.Text = "Not Applicable"
    End If
    
    ' Set default Contract Manager to empty
    Dim cm As String
    cm = ""  ' Default value
    
    ' Set the Contract Manager in the combo box, validate it's in the list
    Dim validCM As Boolean
    validCM = False
    For i = 0 To frmInput.cboCM.ListCount - 1
        If frmInput.cboCM.List(i) = cm Then
            frmInput.cboCM.Text = cm
            validCM = True
            Exit For
        End If
    Next i
    
    ' If not a valid contract manager, set ListIndex to -1 (no selection)
    If Not validCM Then
        frmInput.cboCM.ListIndex = -1
    End If
    
    ' Use the UpdateEmailSubject function to format the email subject
    Call UpdateEmailSubject
End Sub

' Function to update email subject based on form fields
Public Sub UpdateEmailSubject()
    ' Apply the formula logic to update email subject
    If frmInput.cboRequest.Text = "Contract Review" Then
        ' For Contract Review: "Assignment: Request# [LM], [CDR]-[Client]-[ContractType]"
        frmInput.txtEmailSubject.Text = "Assignment: Request# " & frmInput.txtLM.Text & ", " & frmInput.txtCDR.Text & "-" & frmInput.txtClient.Text & "-" & frmInput.txtContractType.Text
        
        ' Clean up the subject by removing excess hyphens if CDR is empty
        If frmInput.txtCDR.Text = "" Then
            frmInput.txtEmailSubject.Text = Replace(frmInput.txtEmailSubject.Text, ", -", ", ")
        End If
    Else
        ' For other types (like Contract Upload): "[Last 7 chars of LM]-[Client]"
        ' Extract the last 7 characters of LM number
        Dim lmRightPart As String
        If Len(frmInput.txtLM.Text) >= 7 Then
            lmRightPart = Right(frmInput.txtLM.Text, 7)
        Else
            lmRightPart = frmInput.txtLM.Text
        End If
        frmInput.txtEmailSubject.Text = lmRightPart & "-" & frmInput.txtClient.Text
    End If
End Sub

' Subroutine to clear all form fields
Public Sub ClearFormFields()
    ' Clear all text boxes and reset selection fields
    frmInput.txtCoordinator.Text = ""
    frmInput.txtDate.Text = Format(Date, "mm/dd/yyyy")  ' Default to today's date
    frmInput.txtEmailSubject.Text = ""
    frmInput.txtLM.Text = ""
    frmInput.txtClient.Text = ""
    frmInput.txtContractType.Text = ""
    frmInput.txtLOB.Text = ""
    frmInput.cboRCL.Text = ""
    frmInput.txtRCLemail.Text = ""
    frmInput.txtCDR.Text = ""
    frmInput.txtRemarks.Text = ""
    frmInput.txtEmailOutput.Text = ""
    
    ' Reset combo boxes - use ListIndex property instead of Text
    frmInput.cboRequest.ListIndex = -1
    frmInput.cboRegion.ListIndex = -1
    frmInput.cboCM.ListIndex = -1
    
    ' Reset checkboxes
    frmInput.chkComet.Value = False
    frmInput.chkOverrideCM.Value = False
    
    ' Lock CM combobox again
    frmInput.cboCM.Locked = True
    
    ' Reset the change status for the cleared form
    ResetChangeStatus
End Sub

' Function to validate required fields before saving
Public Function ValidateRequiredFields() As Boolean
    Dim errorMessage As String
    errorMessage = ""
    
    ' Check each required field and build an error message if any are missing
    If Trim(frmInput.txtCoordinator.Text) = "" Then
        errorMessage = errorMessage & "- Coordinator" & vbCrLf
    End If
    
    If Trim(frmInput.txtDate.Text) = "" Then
        errorMessage = errorMessage & "- Date" & vbCrLf
    End If
    
    If Trim(frmInput.txtEmailSubject.Text) = "" Then
        errorMessage = errorMessage & "- Email Subject" & vbCrLf
    End If
    
    If Trim(frmInput.txtLM.Text) = "" Then
        errorMessage = errorMessage & "- LM" & vbCrLf
    End If
    
    If Trim(frmInput.cboRequest.Text) = "" Then
        errorMessage = errorMessage & "- Contract Request" & vbCrLf
    End If
    
    If Trim(frmInput.txtClient.Text) = "" Then
        errorMessage = errorMessage & "- Client Name" & vbCrLf
    End If
    
    If Trim(frmInput.cboRegion.Text) = "" Then
        errorMessage = errorMessage & "- Region" & vbCrLf
    End If
    
    ' Check if additional fields are required for Contract Review
    If frmInput.cboRequest.Text = "Contract Review" Then
        If Trim(frmInput.cboRCL.Text) = "" Then
            errorMessage = errorMessage & "- Assigned RCL" & vbCrLf
        End If
    End If
    
    If frmInput.cboRequest.Text = "Contract Review" Then
        If Trim(frmInput.txtRCLemail.Text) = "" Then
            errorMessage = errorMessage & "- RCL email address" & vbCrLf
        ElseIf Not IsValidEmail(Trim(frmInput.txtRCLemail.Text)) Then
            errorMessage = errorMessage & "- RCL email address (invalid format)" & vbCrLf
        End If
    End If
    
    ' If there are any missing fields, show error message and return False
    If errorMessage <> "" Then
        DisplayMessage "The following required fields are missing:" & vbCrLf & vbCrLf & _
               errorMessage & vbCrLf & _
               "Please fill in all required fields before saving.", True
        ValidateRequiredFields = False
    Else
        ValidateRequiredFields = True
    End If
End Function

' Function to validate required fields for email
Public Function ValidateRequiredFieldsForEmail() As Boolean
    Dim errorMessage As String
    Dim contractRequestType As String
    
    errorMessage = ""
    contractRequestType = frmInput.cboRequest.Text
    
    ' Check each required field
    If Trim(frmInput.txtCoordinator.Text) = "" Then
        errorMessage = errorMessage & "- Coordinator" & vbCrLf
    End If
    
    If Trim(frmInput.txtDate.Text) = "" Then
        errorMessage = errorMessage & "- Date" & vbCrLf
    End If
    
    If Trim(frmInput.txtEmailSubject.Text) = "" Then
        errorMessage = errorMessage & "- Email Subject" & vbCrLf
    End If
    
    If Trim(frmInput.txtLM.Text) = "" Then
        errorMessage = errorMessage & "- LM" & vbCrLf
    End If
    
    If Trim(frmInput.cboRequest.Text) = "" Then
        errorMessage = errorMessage & "- Contract Request" & vbCrLf
    End If
    
    If Trim(frmInput.txtClient.Text) = "" Then
        errorMessage = errorMessage & "- Client Name" & vbCrLf
    End If
    
    If Trim(frmInput.cboRegion.Text) = "" Then
        errorMessage = errorMessage & "- Region" & vbCrLf
    End If
    
    ' Additional required fields for email - RCL and CDR only required for certain request types
    If contractRequestType <> "Contract Upload" And contractRequestType <> "Out of Scope" And contractRequestType <> "Duplicate Request" Then
        If Trim(frmInput.cboRCL.Text) = "" Then
            errorMessage = errorMessage & "- Assigned RCL" & vbCrLf
        End If
        
        If Trim(frmInput.txtRCLemail.Text) = "" Then
            errorMessage = errorMessage & "- RCL Email" & vbCrLf
        ElseIf Not IsValidEmail(Trim(frmInput.txtRCLemail.Text)) Then
            errorMessage = errorMessage & "- RCL Email (invalid format)" & vbCrLf
        End If
        
        If Trim(frmInput.txtCDR.Text) = "" Then
            errorMessage = errorMessage & "- CDR or Request number" & vbCrLf
        End If
    End If
    
    If Trim(frmInput.cboCM.Text) = "" Then
        errorMessage = errorMessage & "- Contract Manager" & vbCrLf
    End If
    
    ' If there are any missing fields, show error message and return False
    If errorMessage <> "" Then
        DisplayMessage "The following required fields must be filled before sending an email:" & vbCrLf & vbCrLf & _
               errorMessage & vbCrLf & _
               "Please fill in all required fields.", True
        ValidateRequiredFieldsForEmail = False
    Else
        ValidateRequiredFieldsForEmail = True
    End If
End Function

' Function to save form data to the worksheet
Public Sub SaveFormData()
    ' Validate required fields before proceeding
    If Not ValidateRequiredFields Then
        Exit Sub
    End If
    
    If frmInput.cboRequest.Text = "Contract Review" Then
        If Trim(frmInput.txtRCLemail.Text) <> "" And Not IsValidEmail(Trim(frmInput.txtRCLemail.Text)) Then
            DisplayMessage "The RCL Email address format is invalid. Please enter a valid email address.", True
            frmInput.txtRCLemail.SetFocus
            Exit Sub
        End If
    End If
    
    ' Reference the workbook and worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim updateExisting As Boolean
    Dim rowToUse As Long
    Dim savedRowNumber As Long
    
    ' Disable events to prevent Worksheet_Change from triggering
'    Application.EnableEvents = False
    
    ' Set reference to the Box Assignment Tracker workbook and sheet
    ' Assumes the workbook is already open
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("Box Assignment Tracker")
    
    ' Check if we're updating an existing record or creating a new one
    updateExisting = False
    If Trim(frmInput.lblRowNum.Caption) <> "" Then
        ' Try to get the row number from the label caption
        On Error Resume Next
        ' Extract the number from the caption (which may contain text like "Record saved at row: 42")
        Dim captionText As String
        Dim rowPosition As Long
        
        captionText = frmInput.lblRowNum.Caption
        rowPosition = InStr(1, captionText, "row:")
        
        If rowPosition > 0 Then
            ' Extract just the number after "row:"
            savedRowNumber = CLng(Trim(mid(captionText, rowPosition + 4)))
            rowToUse = savedRowNumber
            updateExisting = True
        End If
'        On Error GoTo ErrorHandler
    End If
    
    ' If not updating an existing record, find the next available row
    If Not updateExisting Then
        ' Check multiple columns (A through E) to find the last row with data
        Dim lastRowA As Long, lastRowB As Long, lastRowD As Long, lastRowE As Long
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        
        ' Use the maximum row number from all checked columns
        rowToUse = Application.WorksheetFunction.Max(lastRowA, lastRowB, lastRowD, lastRowE) + 1
    End If
    
    ' Display a confirmation message with the LM number
    Dim confirmMsg As String
    If updateExisting Then
        confirmMsg = "You're updating the record for ticket number " & frmInput.txtLM.Text & ". Continue?"
    Else
        confirmMsg = "You're saving a new record for ticket number " & frmInput.txtLM.Text & ". Continue?"
    End If
    
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "Confirm Save") = vbNo Then
'        Application.EnableEvents = True
        Exit Sub
    End If
    
    ' Save form data to the worksheet
    ws.Cells(rowToUse, "A").Value = frmInput.txtCoordinator.Text
    ws.Cells(rowToUse, "B").Value = frmInput.txtDate.Text
    ws.Cells(rowToUse, "C").Value = frmInput.txtEmailSubject.Text
    ws.Cells(rowToUse, "D").Value = frmInput.txtLM.Text
    ws.Cells(rowToUse, "E").Value = frmInput.cboRequest.Text  ' Changed from txtRequest to cboRequest
    
    ' Check if Comet is checked and set value accordingly
    If frmInput.chkComet.Value = True Then
        ws.Cells(rowToUse, "G").Value = "Yes"
    Else
        ws.Cells(rowToUse, "G").Value = "No"
    End If
    
    ws.Cells(rowToUse, "H").Value = frmInput.txtClient.Text
    ws.Cells(rowToUse, "I").Value = frmInput.txtContractType.Text
    ws.Cells(rowToUse, "J").Value = frmInput.txtLOB.Text
    ws.Cells(rowToUse, "K").Value = frmInput.cboRegion.Text  ' Changed from txtRegion to cboRegion
   
    ws.Cells(rowToUse, "P").Value = frmInput.cboRCL.Text
    ws.Cells(rowToUse, "R").Value = frmInput.txtCDR.Text
    ws.Cells(rowToUse, "S").Value = frmInput.txtRemarks.Text
    
    ' Handle Contract Manager field based on Override CM checkbox and value
    If frmInput.chkOverrideCM.Value = False Then
        ' If override is not checked, use the default behavior
'        If frmInput.cboCM.Text = "" Then
            ' Trigger the worksheet_change event by updating region
            ws.Cells(rowToUse, "K").Value = frmInput.cboRegion.Text
            ' Update the form with the value from the worksheet
            frmInput.cboCM.Text = ws.Cells(rowToUse, "Q").Value
        Else
            ' Still trigger the worksheet_change event for other processing
            ws.Cells(rowToUse, "K").Value = frmInput.cboRegion.Text
            ' Save the current value to the worksheet
            ws.Cells(rowToUse, "Q").Value = frmInput.cboCM.Text
'        End If
'    Else
'        ' If override is checked, use the user-selected value regardless
'        ws.Cells(rowToUse, "Q").Value = frmInput.cboCM.Text
    End If
    
    ' Update the row number display
    frmInput.lblRowNum.Caption = "Record saved at row: " & rowToUse
    
    ' Change the save button caption to "Update Record" after saving
    frmInput.cmdSave.Caption = "Update Record"
    
    ' Re-enable events after all data is entered
    'Application.EnableEvents = True
    
    ' Display success message
    If updateExisting Then
        DisplayMessage "Record updated successfully at row " & rowToUse, True
    Else
        DisplayMessage "Record saved successfully at row " & rowToUse, True
    End If
    
    ' Reset the change status after successful save
    ResetChangeStatus
    
    Exit Sub
    
'ErrorHandler:
    ' Make sure to re-enable events even if an error occurs
'    Application.EnableEvents = True
    DisplayMessage "Error saving data: " & Err.Description, True
End Sub

' Function to load RCL data from the References sheet
Public Sub LoadRCLData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set Reference to the References worksheet
    Set ws = ThisWorkbook.Sheets("References")
    
    ' Find the last row with data in column R (changed from column C)
    lastRow = ws.Cells(ws.Rows.Count, "R").End(xlUp).Row
    
    ' Clear the RCL combo box
    frmInput.cboRCL.Clear
    
    ' Loop through the data starting from row 2 (assuming row 1 is header)
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, "R").Value) Then
            frmInput.cboRCL.AddItem ws.Cells(i, "R").Value
        End If
    Next i
    
    ' Set to no selection initially
    frmInput.cboRCL.ListIndex = -1
End Sub

' Function to mark the form as having unsaved changes
Public Sub MarkAsChanged()
    ' Set the global flag to indicate changes
    HasChanges = True
    
    ' Update the label to show unsaved changes
    frmInput.lblChangeStatus.Caption = "Unsaved Changes"
    frmInput.lblChangeStatus.ForeColor = RGB(192, 0, 0) ' Dark red color
End Sub

' Function to reset the change status after saving
Public Sub ResetChangeStatus()
    ' Reset the global flag
    HasChanges = False
    
    ' Clear the label
    frmInput.lblChangeStatus.Caption = ""
End Sub

