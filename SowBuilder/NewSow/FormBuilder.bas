Attribute VB_Name = "FormBuilder"
Option Explicit

' This module automatically creates the SOW Builder form with all controls
' Run the CreateSOWBuilderForm sub to generate the form

Public Sub CreateSOWBuilderForm()
    Dim frmNew As Object
    Dim ctrl As Object
    
    ' Delete form if it already exists
    On Error Resume Next
    For Each ctrl In ThisWorkbook.VBProject.VBComponents
        If ctrl.Name = "frmSOWBuilderSingle" Then
            ThisWorkbook.VBProject.VBComponents.Remove ctrl
            Exit For
        End If
    Next ctrl
    On Error GoTo 0
    
    ' Add the form
    Set frmNew = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = UserForm
    With frmNew
        .Name = "frmSOWBuilderSingle"
        .Properties("Caption") = "SOW Builder"
        .Properties("Width") = 600
        .Properties("Height") = 680 ' Increased to accommodate all controls
        .Properties("StartUpPosition") = 1 ' CenterOwner
    End With
    
    ' Create all the controls
    
    ' 1. Client Information Section
    AddFrame frmNew, "fraClientInfo", "General & Contract Information", 15, 15, 570, 180
    
    AddLabel frmNew, "lblDate", "Date:", 25, 35
    AddTextBox frmNew, "txtDate", 120, 30, 120
    
    AddLabel frmNew, "lblContactName", "Contact Name:", 25, 60
    AddTextBox frmNew, "txtContactName", 120, 55, 160
    
    AddLabel frmNew, "lblCompanyName", "Company Name:", 25, 85
    AddTextBox frmNew, "txtCompanyName", 120, 80, 160
    
    AddLabel frmNew, "lblAddress1", "Address Line 1:", 25, 110
    AddTextBox frmNew, "txtAddress1", 120, 105, 160
    
    AddLabel frmNew, "lblAddress2", "Address Line 2:", 25, 135
    AddTextBox frmNew, "txtAddress2", 120, 130, 160
    
    AddLabel frmNew, "lblWTWParty", "WTW Contracting:", 290, 35
    AddComboBox frmNew, "cmbWTWParty", 400, 30, 160, 2 ' Style: 2-Dropdown List
    
    AddLabel frmNew, "lblClientName", "Client Legal Name:", 290, 60
    AddTextBox frmNew, "txtClientName", 400, 55, 160
    
    AddLabel frmNew, "lblStartDate", "Start Date:", 290, 85
    AddTextBox frmNew, "txtStartDate", 400, 80, 120
    
    AddLabel frmNew, "lblEndDate", "End Date:", 290, 110
    AddTextBox frmNew, "txtEndDate", 400, 105, 120
    
    ' 2. Compensation Options Section
    AddFrame frmNew, "fraCompOptions", "Compensation Options", 15, 205, 570, 100
    
    AddOptionButton frmNew, "optA", "Option A: Fee Only", 25, 225
    AddOptionButton frmNew, "optB", "Option B: Fee + Commission", 25, 250
    AddOptionButton frmNew, "optC", "Option C: Fee Offset by Commission", 25, 275
    AddOptionButton frmNew, "optD", "Option D: Commission Only", 280, 225
    
    AddLabel frmNew, "lblAnnualFee", "Annual Fee:", 280, 255
    AddTextBox frmNew, "txtAnnualFee", 350, 250, 120
    
    ' 3. Fee Details Section
    AddFrame frmNew, "fraFeeDetails", "Fee Details", 15, 315, 275, 80
    
    AddFrame frmNew, "fraBillingOptions", "Billing Options", 35, 335, 235, 50
    
    AddOptionButton frmNew, "optMilestone", "Milestone Billing", 45, 355
    AddOptionButton frmNew, "optInstallments", "Regular Installments", 145, 355
    
    ' 4. Commission Details Section
    AddFrame frmNew, "fraCommissionDetails", "Commission Details", 300, 315, 285, 160
    
    AddLabel frmNew, "lblPolicyName", "Policy Name:", 310, 335
    AddTextBox frmNew, "txtPolicyName", 380, 332, 190
    
    AddLabel frmNew, "lblCommission", "Commission (%):", 310, 360
    AddTextBox frmNew, "txtCommission", 395, 357, 50
    
    AddCommandButton frmNew, "btnAddPolicy", "Add Policy", 310, 385, 80, 22
    AddCommandButton frmNew, "btnRemovePolicy", "Remove Policy", 400, 385, 80, 22
    
    AddListBox frmNew, "lstPolicies", 310, 410, 260, 60
    
    ' 5. Optional Clauses Section
    AddFrame frmNew, "fraOptionalClauses", "Optional Clauses", 15, 405, 275, 70
    
    AddCheckBox frmNew, "chkAutoRenewal", "Auto-Renewal", 25, 425
    AddCheckBox frmNew, "chkGDPR", "GDPR Applies", 25, 445
    
    ' 6. Additional Notes Section
    AddFrame frmNew, "fraAdditionalNotes", "Additional Notes", 15, 485, 570, 90
    
    AddMultilineTextBox frmNew, "txtAdditionalNotes", 25, 505, 550, 60
    
    ' 7. Action Buttons
    AddCommandButton frmNew, "btnGenerate", "Generate Doc", 415, 580, 80, 25
    AddCommandButton frmNew, "btnCancel", "Cancel", 505, 580, 80, 25
    
    MsgBox "Form created successfully! You can now access 'frmSOWBuilderSingle' in your VBA editor.", vbInformation
End Sub

' Helper functions to add different control types

Private Sub AddFrame(frm As Object, name As String, caption As String, _
                    left As Integer, top As Integer, width As Integer, height As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.Frame.1", name, True)
    With ctrl
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Sub

Private Sub AddLabel(frm As Object, name As String, caption As String, _
                    left As Integer, top As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1", name, True)
    With ctrl
        .Caption = caption
        .Left = left
        .Top = top
        .AutoSize = True
    End With
End Sub

Private Sub AddTextBox(frm As Object, name As String, _
                      left As Integer, top As Integer, width As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1", name, True)
    With ctrl
        .Left = left
        .Top = top
        .Width = width
        .Height = 20
    End With
End Sub

Private Sub AddComboBox(frm As Object, name As String, _
                       left As Integer, top As Integer, width As Integer, style As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.ComboBox.1", name, True)
    With ctrl
        .Left = left
        .Top = top
        .Width = width
        .Style = style
    End With
End Sub

Private Sub AddOptionButton(frm As Object, name As String, caption As String, _
                           left As Integer, top As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.OptionButton.1", name, True)
    With ctrl
        .Caption = caption
        .Left = left
        .Top = top
        .Width = 200
        .Height = 18
    End With
End Sub

Private Sub AddCheckBox(frm As Object, name As String, caption As String, _
                       left As Integer, top As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.CheckBox.1", name, True)
    With ctrl
        .Caption = caption
        .Left = left
        .Top = top
        .Width = 200
        .Height = 18
    End With
End Sub

Private Sub AddCommandButton(frm As Object, name As String, caption As String, _
                            left As Integer, top As Integer, width As Integer, height As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1", name, True)
    With ctrl
        .Caption = caption
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Sub

Private Sub AddListBox(frm As Object, name As String, _
                      left As Integer, top As Integer, width As Integer, height As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.ListBox.1", name, True)
    With ctrl
        .Left = left
        .Top = top
        .Width = width
        .Height = height
    End With
End Sub

Private Sub AddMultilineTextBox(frm As Object, name As String, _
                               left As Integer, top As Integer, width As Integer, height As Integer)
    Dim ctrl As Object
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1", name, True)
    With ctrl
        .Left = left
        .Top = top
        .Width = width
        .Height = height
        .MultiLine = True
        .ScrollBars = 2 ' 2 = Vertical
    End With
End Sub 
