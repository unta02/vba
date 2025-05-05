VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSOWBuilderSingle 
   Caption         =   "SOW Builder"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9615
   OleObjectBlob   =   "frmSOWBuilderSingle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSOWBuilderSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Collection to store policies and commission rates
Private policyCollection As Collection

' Initialize the form
Private Sub UserForm_Initialize()
    ' Initialize the collection for policies
    Set policyCollection = New Collection
    
    ' Initialize dropdown for WTW Contracting Party
    With cmbWTWParty
        .AddItem "Willis Towers Watson US LLC"
        .AddItem "Willis Towers Watson Canada"
        .AddItem "Willis Towers Watson UK Limited"
        .AddItem "Willis Towers Watson Australia"
        .AddItem "Willis Towers Watson Other"
    End With
    
    ' Set default dates
    txtDate.Value = Format(Date, "mm/dd/yyyy")
    txtStartDate.Value = Format(Date, "mm/dd/yyyy")
    txtEndDate.Value = Format(DateAdd("yyyy", 1, Date), "mm/dd/yyyy")
    
    ' Set option A as default
    optA.Value = True
    
    ' Default billing option
    optMilestone.Value = True
    
    ' Update UI based on default options
    UpdateCompensationVisibility
End Sub

' Update visibility of controls based on compensation option selected
Private Sub UpdateCompensationVisibility()
    ' Hide/show fee related controls
    fraFeeDetails.Visible = (optA.Value Or optB.Value Or optC.Value)
    fraBillingOptions.Visible = (optA.Value Or optB.Value Or optC.Value)
    
    ' Hide/show commission related controls
    fraCommissionDetails.Visible = (optB.Value Or optC.Value Or optD.Value)
    
    ' Enable/disable commission controls based on option
    If optB.Value Or optC.Value Or optD.Value Then
        txtPolicyName.Enabled = True
        txtCommission.Enabled = True
        btnAddPolicy.Enabled = True
        btnRemovePolicy.Enabled = True
        lstPolicies.Enabled = True
    Else
        txtPolicyName.Enabled = False
        txtCommission.Enabled = False
        btnAddPolicy.Enabled = False
        btnRemovePolicy.Enabled = False
        lstPolicies.Enabled = False
    End If
End Sub

' Compensation option change handlers
Private Sub optA_Click()
    UpdateCompensationVisibility
End Sub

Private Sub optB_Click()
    UpdateCompensationVisibility
End Sub

Private Sub optC_Click()
    UpdateCompensationVisibility
End Sub

Private Sub optD_Click()
    UpdateCompensationVisibility
End Sub

' Policy list management
Private Sub btnAddPolicy_Click()
    ' Validate input
    If Trim(txtPolicyName.Value) = "" Then
        MsgBox "Please enter a policy name.", vbExclamation
        txtPolicyName.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtCommission.Value) Then
        MsgBox "Please enter a valid commission percentage.", vbExclamation
        txtCommission.SetFocus
        Exit Sub
    End If
    
    ' Format policy entry
    Dim policyEntry As String
    policyEntry = txtPolicyName.Value & " (" & txtCommission.Value & "% commission)"
    
    ' Add to list and collection
    lstPolicies.AddItem policyEntry
    policyCollection.Add policyEntry
    
    ' Clear inputs
    txtPolicyName.Value = ""
    txtCommission.Value = ""
    txtPolicyName.SetFocus
End Sub

Private Sub btnRemovePolicy_Click()
    If lstPolicies.ListIndex = -1 Then
        MsgBox "Please select a policy to remove.", vbExclamation
    Else
        ' Remove from list and rebuild collection
        lstPolicies.RemoveItem lstPolicies.ListIndex
        
        ' Rebuild collection from listbox
        Set policyCollection = New Collection
        Dim i As Integer
        For i = 0 To lstPolicies.ListCount - 1
            policyCollection.Add lstPolicies.List(i)
        Next i
    End If
End Sub

' Generate document button
Private Sub btnGenerate_Click()
    ' Validate required fields
    If Not ValidateForm Then Exit Sub
    
    ' Create dictionary for client info
    Dim clientInfo As Object
    Set clientInfo = SOWBuilderSinglePage.CreateDictionary
    
    ' Fill client info
    clientInfo.Add "Date", txtDate.Value
    clientInfo.Add "ContactName", txtContactName.Value
    clientInfo.Add "CompanyName", txtCompanyName.Value
    clientInfo.Add "Address1", txtAddress1.Value
    clientInfo.Add "Address2", txtAddress2.Value
    clientInfo.Add "WTWParty", cmbWTWParty.Value
    clientInfo.Add "ClientName", txtClientName.Value
    clientInfo.Add "StartDate", txtStartDate.Value
    clientInfo.Add "EndDate", txtEndDate.Value
    
    ' Determine compensation option
    Dim compensationOption As String
    If optA.Value Then
        compensationOption = "A"
    ElseIf optB.Value Then
        compensationOption = "B"
    ElseIf optC.Value Then
        compensationOption = "C"
    ElseIf optD.Value Then
        compensationOption = "D"
    End If
    
    ' Determine billing option
    Dim billingOption As String
    If optMilestone.Value Then
        billingOption = "Milestone"
    ElseIf optInstallments.Value Then
        billingOption = "Installments"
    Else
        billingOption = ""
    End If
    
    ' Create dictionary for optional clauses
    Dim optionalClauses As Object
    Set optionalClauses = SOWBuilderSinglePage.CreateDictionary
    optionalClauses.Add "AutoRenewal", chkAutoRenewal.Value
    optionalClauses.Add "GDPR", chkGDPR.Value
    
    ' Generate the document
    SOWBuilderSinglePage.GenerateSOWDocument clientInfo, compensationOption, _
        txtAnnualFee.Value, billingOption, policyCollection, optionalClauses, txtAdditionalNotes.Value
    
    ' Unload form
    Unload Me
End Sub

' Validate all required fields
Private Function ValidateForm() As Boolean
    ' Check client information
    If Trim(txtClientName.Value) = "" Then
        MsgBox "Please enter the Client Legal Name.", vbExclamation
        txtClientName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Trim(txtContactName.Value) = "" Then
        MsgBox "Please enter the Contact Name.", vbExclamation
        txtContactName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If cmbWTWParty.ListIndex = -1 Then
        MsgBox "Please select a WTW Contracting Party.", vbExclamation
        cmbWTWParty.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Check dates
    If Not IsDate(txtStartDate.Value) Then
        MsgBox "Please enter a valid Start Date.", vbExclamation
        txtStartDate.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Not IsDate(txtEndDate.Value) Then
        MsgBox "Please enter a valid End Date.", vbExclamation
        txtEndDate.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Check compensation options
    If Not (optA.Value Or optB.Value Or optC.Value Or optD.Value) Then
        MsgBox "Please select a compensation option.", vbExclamation
        ValidateForm = False
        Exit Function
    End If
    
    ' Check fee amount if applicable
    If (optA.Value Or optB.Value Or optC.Value) Then
        If Not IsNumeric(txtAnnualFee.Value) Then
            MsgBox "Please enter a valid Annual Fee amount.", vbExclamation
            txtAnnualFee.SetFocus
            ValidateForm = False
            Exit Function
        End If
        
        ' Check billing type selection
        If Not (optMilestone.Value Or optInstallments.Value) Then
            MsgBox "Please select a billing option.", vbExclamation
            ValidateForm = False
            Exit Function
        End If
    End If
    
    ' Check policy list if applicable
    If (optB.Value Or optC.Value Or optD.Value) And lstPolicies.ListCount = 0 Then
        MsgBox "Please add at least one policy with commission rate.", vbExclamation
        txtPolicyName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' All validations passed
    ValidateForm = True
End Function

' Cancel button handler
Private Sub btnCancel_Click()
    Unload Me
End Sub 
