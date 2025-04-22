' DataProcessing.bas - Data Processing Module

Option Explicit

' Function to extract LM number from HTML content
Public Function ExtractLMNumber(htmlContent As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim lmNumber As String
    
    startPos = InStr(htmlContent, "Received - Legal Matter")
    If startPos > 0 Then
        ' Move to the end of "Received - Legal Matter"
        startPos = startPos + Len("Received - Legal Matter")
        
        ' Look for the end of the LM number (usually before a </strong> tag)
        endPos = InStr(startPos, htmlContent, "</strong>")
        If endPos > 0 Then
            lmNumber = Trim(mid(htmlContent, startPos, endPos - startPos))
        End If
    End If
    
    ExtractLMNumber = lmNumber
End Function

' Function to extract Client Name from HTML content
Public Function ExtractClientName(htmlContent As String) As String
    Dim startPattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim clientName As String
    
    startPattern = "<td><b>Client or Supplier Name (full legal entity name if known)</b></td>" & vbCrLf & "<td>"
    startPos = InStr(htmlContent, startPattern)
    
    If startPos > 0 Then
        ' Move to start of client name
        startPos = startPos + Len(startPattern)
        
        ' Find end of client name (end of TD tag)
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            clientName = mid(htmlContent, startPos, endPos - startPos)
        End If
    End If
    
    ExtractClientName = clientName
End Function

' Function to extract Request Type from HTML content
Public Function ExtractRequestType(htmlContent As String) As String
    Dim startPattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim requestType As String
    
    startPattern = "<td><b>Request Type</b></td>" & vbCrLf & "<td>"
    startPos = InStr(htmlContent, startPattern)
    
    If startPos > 0 Then
        ' Move to start of request type
        startPos = startPos + Len(startPattern)
        
        ' Find end of request type (end of TD tag)
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            requestType = Trim(mid(htmlContent, startPos, endPos - startPos))
            
            ' Recode the request type based on rules
            Select Case UCase(requestType)
                Case "CLIENT"
                    requestType = "Contract Review"
                Case "CONTRACT UPLOAD"
                    requestType = "Contract Upload"
                Case Else
                    requestType = "" ' Leave blank if not matching our valid values
            End Select
        End If
    End If
    
    ExtractRequestType = requestType
End Function

' Function to extract Contract Type from HTML content
Public Function ExtractContractType(htmlContent As String) As String
    Dim startPattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim contractType As String
    
    startPattern = "<td><b>Document Type being requested</b></td>" & vbCrLf & "<td>"
    startPos = InStr(htmlContent, startPattern)
    
    If startPos > 0 Then
        ' Move to start of contract type
        startPos = startPos + Len(startPattern)
        
        ' Find end of contract type (end of TD tag)
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            contractType = mid(htmlContent, startPos, endPos - startPos)
        End If
    End If
    
    ExtractContractType = contractType
End Function

' Function to extract LOB from HTML content
Public Function ExtractLOB(htmlContent As String) As String
    Dim startPattern As String, additionalPattern As String
    Dim startPos As Long, endPos As Long
    Dim lob As String, additionalLob As String
    Dim allLobs As String
    
    ' Extract primary Line of Business
    startPattern = "<td><b>Line of Business</b></td>" & vbCrLf & "<td>"
    startPos = InStr(htmlContent, startPattern)
    
    If startPos > 0 Then
        ' Move to start of LOB
        startPos = startPos + Len(startPattern)
        
        ' Find end of LOB (end of TD tag)
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            lob = mid(htmlContent, startPos, endPos - startPos)
            
            ' Decode HTML entities before processing
            lob = DecodeHtml(lob)
            
            ' Extract only the first part before "("
            Dim parenPos As Long
            parenPos = InStr(lob, "(")
            If parenPos > 0 Then
                lob = Left(Trim(lob), 3)
            End If
        End If
    End If
    
    ' Extract Additional Line of Business if it exists
    additionalPattern = "<td><b>Additional Line Of Business</b></td>" & vbCrLf & "<td>"
    startPos = InStr(htmlContent, additionalPattern)
    
    If startPos > 0 Then
        ' Move to start of additional LOB
        startPos = startPos + Len(additionalPattern)
        
        ' Find end of additional LOB (end of TD tag)
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            additionalLob = mid(htmlContent, startPos, endPos - startPos)
            
            ' Decode HTML entities
            additionalLob = DecodeHtml(additionalLob)
            
            ' Process multiple additional LOBs if they exist
            If InStr(additionalLob, ",") > 0 Then
                Dim lobArray() As String
                Dim i As Long
                Dim tempLob As String
                
                ' Split by comma
                lobArray = Split(additionalLob, ",")
                
                ' Process each additional LOB
                For i = 0 To UBound(lobArray)
                    tempLob = Trim(lobArray(i))
                    ' Extract first 3 letters before "("
                    parenPos = InStr(tempLob, "(")
                    If parenPos > 0 Then
                        tempLob = Left(Trim(tempLob), 3)
                    End If
                    
                    ' Add to allLobs if not empty
                    If tempLob <> "" Then
                        If allLobs = "" Then
                            allLobs = tempLob
                        Else
                            allLobs = allLobs & ", " & tempLob
                        End If
                    End If
                Next i
            Else
                ' Single additional LOB
                parenPos = InStr(additionalLob, "(")
                If parenPos > 0 Then
                    additionalLob = Left(Trim(additionalLob), 3)
                End If
                
                If additionalLob <> "" Then
                    allLobs = additionalLob
                End If
            End If
        End If
    End If
    
    ' Combine primary and additional LOBs
    If lob <> "" Then
        If allLobs = "" Then
            allLobs = lob
        Else
            allLobs = lob & ", " & allLobs
        End If
    End If
    
    ' If no LOBs found, return default value
    If Trim(allLobs) = "" Then
        allLobs = "N/A"
    End If
    
    ExtractLOB = allLobs
End Function

' Function to lookup region information from Admin Sheet
Public Function LookupRegion(regionValue As String) As String
    ' Function to lookup region information from Admin Sheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Admin")
    If ws Is Nothing Then
        LookupRegion = regionValue ' Return original value if sheet not found
        Exit Function
    End If
    
    ' Find the last row in column AB (Region lookup column)
    lastRow = ws.Cells(ws.Rows.Count, "AB").End(xlUp).Row
    
    ' Loop through regions to find a match
    For i = 4 To lastRow ' Start from row 4 as per specification
        If Trim(ws.Cells(i, "AB").Value) = Trim(regionValue) Then
            ' Return the recoded region from column AC
            LookupRegion = ws.Cells(i, "AC").Value
            Exit Function
        End If
    Next i
    
    ' If no match found, return "Not Applicable"
    LookupRegion = "Not Applicable"
End Function

' Function to extract Region from HTML content
Public Function ExtractRegion(htmlContent As String) As String
    Dim startPattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim region As String
    Dim requestType As String
    
    ' First get the request type to determine which field to use
    requestType = ExtractRequestType(htmlContent)
    
    ' Choose the appropriate field based on request type
    If requestType = "Contract Review" Then
        startPattern = "<td><b>Region where services are provided</b></td>" & vbCrLf & "<td>"
    Else
        startPattern = "<td><b>Client / Counterparty Location</b></td>" & vbCrLf & "<td>"
    End If
    
    startPos = InStr(htmlContent, startPattern)
    
    If startPos > 0 Then
        ' Move to start of region
        startPos = startPos + Len(startPattern)
        
        ' Find end of region (end of TD tag)
        endPos = InStr(startPos, htmlContent, "</td>")
        If endPos > 0 Then
            region = Trim(mid(htmlContent, startPos, endPos - startPos))
            
            ' Recode the region using the lookup table
            region = LookupRegion(region)
        End If
    End If
    
    ' If no region found or empty, return "Not Applicable"
    If region = "" Then
        region = "Not Applicable"
    End If
    
    ExtractRegion = region
End Function

' Function to extract Requested For name from HTML content
Public Function ExtractRequestedForName(htmlContent As String) As String
    Dim startPattern As String
    Dim startPos As Long, endPos As Long, nameStart As Long
    Dim fullName As String, firstName As String
    
    startPattern = "Requested For"
    startPos = InStr(htmlContent, startPattern)
    
    If startPos > 0 Then
        ' Find the name after "Requested For"
        nameStart = InStr(startPos, htmlContent, ":")
        If nameStart = 0 Then ' Try with different pattern
            nameStart = InStr(startPos, htmlContent, "</b></td>" & vbCrLf & "<td>")
            If nameStart > 0 Then
                nameStart = nameStart + Len("</b></td>" & vbCrLf & "<td>")
                endPos = InStr(nameStart, htmlContent, "</td>")
                If endPos > 0 Then
                    fullName = Trim(mid(htmlContent, nameStart, endPos - nameStart))
                End If
            End If
        Else
            nameStart = nameStart + 1
            endPos = InStr(nameStart, htmlContent, vbCrLf)
            If endPos > 0 Then
                fullName = Trim(mid(htmlContent, nameStart, endPos - nameStart))
            End If
        End If
        
        ' Extract first name only if full name contains a space
        If InStr(fullName, " ") > 0 Then
            firstName = Left(fullName, InStr(fullName, " ") - 1)
            ExtractRequestedForName = firstName
        Else
            ExtractRequestedForName = fullName
        End If
    Else
        ExtractRequestedForName = ""
    End If
End Function

' Function to extract Contact's email from HTML content
Public Function ExtractContactEmail(htmlContent As String) As String
    Dim startPattern As String
    Dim startPos As Long, endPos As Long, emailStart As Long
    
    startPattern = "Contact's email"
    startPos = InStr(htmlContent, startPattern)
    
    If startPos > 0 Then
        ' Find the email after "Contact's email"
        emailStart = InStr(startPos, htmlContent, ":")
        If emailStart = 0 Then ' Try with different pattern
            emailStart = InStr(startPos, htmlContent, "</b></td>" & vbCrLf & "<td>")
            If emailStart > 0 Then
                emailStart = emailStart + Len("</b></td>" & vbCrLf & "<td>")
                endPos = InStr(emailStart, htmlContent, "</td>")
                If endPos > 0 Then
                    ExtractContactEmail = Trim(mid(htmlContent, emailStart, endPos - emailStart))
                End If
            End If
        Else
            emailStart = emailStart + 1
            endPos = InStr(emailStart, htmlContent, vbCrLf)
            If endPos > 0 Then
                ExtractContactEmail = Trim(mid(htmlContent, emailStart, endPos - emailStart))
            End If
        End If
    Else
        ExtractContactEmail = ""
    End If
End Function

' Function to lookup coordinator information from Admin Sheet
Public Function LookupCoordinator(userID As String) As String
    ' Function to lookup coordinator information from Admin Sheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Admin")
    If ws Is Nothing Then
        LookupCoordinator = userID ' Return original userID if sheet not found
        Exit Function
    End If
    
    ' Find the last row in column W (UserID column)
    lastRow = ws.Cells(ws.Rows.Count, "W").End(xlUp).Row
    
    ' Loop through UserIDs to find a match
    For i = 4 To lastRow ' Start from row 4 as per specification
        If UCase(Trim(ws.Cells(i, "W").Value)) = UCase(Trim(userID)) Then
            ' Return the Name from column X
            LookupCoordinator = ws.Cells(i, "Y").Value
            Exit Function
        End If
    Next i
    
    ' If no match found, return original userID
    LookupCoordinator = userID
End Function

' Function to format plain content for display
Public Function FormatPlainContent(plainContent As String) As String
    Dim formattedContent As String
    Dim lines() As String
    Dim i As Long
    Dim inSummarySection As Boolean
    Dim line As String
    Dim parts() As String
    Dim label As String
    Dim Value As String
    
    ' Split the content into lines
    lines = Split(plainContent, vbCrLf)
    formattedContent = ""
    inSummarySection = False
    
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Skip empty lines and unwanted content
        If line <> "" Then
            ' Skip if line contains "View legal matter"
            If InStr(1, line, "View legal matter", vbTextCompare) > 0 Then
                Exit For
            End If
            
            ' Check for the start of the summary section
            If InStr(1, line, "Summary of the Matter:") > 0 Then
                inSummarySection = True
                formattedContent = formattedContent & vbCrLf & "SUMMARY OF THE MATTER:" & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
            ' Check for the end of relevant content
            ElseIf InStr(1, line, "ENTERPRISE TECHNOLOGY SERVICES") > 0 Then
                Exit For
            ' Process lines within the summary section
            ElseIf inSummarySection Then
                ' Clean up the line and format with colon
                If Left(line, 1) <> "_" And line <> "Summary of the Matter:" Then
                    ' Split the line by tab or multiple spaces
                    parts = Split(line, vbTab)
                    If UBound(parts) < 1 Then
                        ' Try splitting by multiple spaces if tab didn't work
                        parts = Split(line, "  ")
                    End If
                    
                    If UBound(parts) >= 0 Then
                        label = Trim(parts(0))
                        If UBound(parts) >= 1 Then
                            Value = Trim(parts(1))
                        Else
                            Value = ""
                        End If
                        
                        ' Format with consistent spacing and colon
                        formattedContent = formattedContent & label & ": " & Value & vbCrLf & vbCrLf
                    Else
                        formattedContent = formattedContent & line & vbCrLf & vbCrLf
                    End If
                End If
            ' Capture the header information
            ElseIf InStr(1, line, "Received - Legal Matter") > 0 Then
                formattedContent = line & vbCrLf & String(50, "-") & vbCrLf & vbCrLf
            End If
        End If
    Next i
    
    ' Remove any trailing blank lines
    Do While Right(formattedContent, 4) = vbCrLf & vbCrLf
        formattedContent = Left(formattedContent, Len(formattedContent) - 2)
    Loop
    
    FormatPlainContent = formattedContent
End Function

' Function to get Contract Manager Full Name from Admin sheet
Public Function GetContractManagerFullName(cmShortName As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Admin")
    If ws Is Nothing Then
        GetContractManagerFullName = cmShortName
        Exit Function
    End If
    
    ' Find the last row in column W (UserID column)
    lastRow = ws.Cells(ws.Rows.Count, "W").End(xlUp).Row
    
    ' Loop through to find the Contract Manager by Short Name
    For i = 4 To lastRow ' Start from row 4 as specified
        If UCase(Trim(ws.Cells(i, "Y").Value)) = UCase(Trim(cmShortName)) Then
            ' Return the Full Name from column X
            GetContractManagerFullName = ws.Cells(i, "X").Value
            Exit Function
        End If
    Next i
    
    ' If not found, return the original name
    GetContractManagerFullName = cmShortName
End Function

' Function to get Contract Manager Email from Admin sheet
Public Function GetContractManagerEmail(cmShortName As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Admin")
    If ws Is Nothing Then
        GetContractManagerEmail = ""
        Exit Function
    End If
    
    ' Find the last row in column W (UserID column)
    lastRow = ws.Cells(ws.Rows.Count, "W").End(xlUp).Row
    
    ' Loop through to find the Contract Manager by Short Name
    For i = 4 To lastRow ' Start from row 4 as specified
        If UCase(Trim(ws.Cells(i, "Y").Value)) = UCase(Trim(cmShortName)) Then
            ' Return the Email from column Z
            GetContractManagerEmail = ws.Cells(i, "Z").Value
            Exit Function
        End If
    Next i
    
    ' If not found, return empty string
    GetContractManagerEmail = ""
End Function
