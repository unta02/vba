' UtilityFunctions.bas - Utility Functions Module
Attribute VB_Name = "UtilityFunctions"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#Else
    Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#End If

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

' Function to get UTC time
Public Function GetUTCTime() As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim ret As Long
    
    ' Get the timezone information
    ret = GetTimeZoneInformation(tzi)
    
    ' Convert current time to UTC
    ' Bias is in minutes, positive means behind UTC, negative means ahead
    GetUTCTime = DateAdd("n", tzi.Bias, Now())
End Function

' Function to clean up LM number by removing extra spaces and special characters
Public Function CleanLMNumber(lmNum As String) As String
    Dim cleanedLM As String
    Dim i As Long
    Dim char As String
    
    cleanedLM = ""
    
    ' Remove any extra spaces or special characters
    For i = 1 To Len(lmNum)
        char = mid(lmNum, i, 1)
        If char <> " " And char <> "," And char <> ";" And char <> ":" Then
            cleanedLM = cleanedLM & char
        End If
    Next i
    
    CleanLMNumber = cleanedLM
End Function

' Function to check if a file is a document type (PDF or Word)
Public Function IsDocumentFile(fileName As String) As Boolean
    Dim fileExt As String
    
    ' Get the file extension
    If InStr(fileName, ".") > 0 Then
        fileExt = LCase(mid(fileName, InStrRev(fileName, ".")))
    Else
        fileExt = ""
    End If
    
    ' Check if it's a document extension we want
    Select Case fileExt
        Case ".pdf"  ' PDF files
            IsDocumentFile = True
        Case ".doc", ".docx", ".docm"  ' Word documents
            IsDocumentFile = True
        Case ".rtf"  ' Rich Text Format
            IsDocumentFile = True
        Case ".txt"  ' Text files
            IsDocumentFile = True
        Case ".xls", ".xlsx", ".xlsm"  ' Excel files
            IsDocumentFile = True
        Case ".ppt", ".pptx", ".pptm"  ' PowerPoint files
            IsDocumentFile = True
        Case Else
            IsDocumentFile = False
    End Select
End Function

' Function to check if an email address has a valid format
Public Function IsValidEmail(email As String) As Boolean
    Dim regex As Object
    Dim pattern As String
    
    ' Create regex object
    On Error Resume Next
    Set regex = CreateObject("VBScript.RegExp")
    
    If regex Is Nothing Then
        ' If regex isn't available, use a simpler validation
        IsValidEmail = InStr(email, "@") > 0 And InStr(email, ".") > InStr(email, "@")
        Exit Function
    End If
    
    ' Set up the regex pattern for basic email validation
    pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    
    With regex
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    ' Test if the email matches the pattern
    IsValidEmail = regex.Test(email)
    
    Set regex = Nothing
End Function

Public Function DecodeHtml(htmlText As String) As String
    Dim decodedText As String
    decodedText = htmlText
    
    ' Replace common HTML entities
    decodedText = Replace(decodedText, "&amp;", "&")
    decodedText = Replace(decodedText, "&lt;", "<")
    decodedText = Replace(decodedText, "&gt;", ">")
    decodedText = Replace(decodedText, "&quot;", """")
    decodedText = Replace(decodedText, "&#39;", "'")
    decodedText = Replace(decodedText, "&nbsp;", " ")
    
    DecodeHtml = decodedText
End Function

' Function to display a message in the form message label
Public Function DisplayMessage(msgText As String, Optional isVisible As Boolean = True, Optional durationSeconds As Integer = 0) As Boolean
    ' Access the UserForm object - will need to be called from the form
    ' This function should be called from the UserForm as: DisplayMessage "Your message"
    On Error Resume Next
    
    If Err.Number = 0 Then
        ' Set message text and visibility
        frmInput.lblMsg.Text = msgText
        frmInput.lblMsg.Visible = isVisible
        
        ' If duration is specified, set up timer to hide message
        If durationSeconds > 0 Then
            ' Need to use Application.OnTime to schedule hiding the message
            ' This would require additional code in the workbook module
        End If
        
        DisplayMessage = True
    Else
        DisplayMessage = False
    End If
    
    On Error GoTo 0
End Function 