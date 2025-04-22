Option Explicit

Public Function GetLiabilityType(ByVal cellValue As String) As String
    ' Convert to lowercase for easier comparison
    Dim value As String
    value = LCase(cellValue)
    
    ' Check for uncapped liability indicators
    If InStr(1, value, "uncapped") > 0 Or _
       InStr(1, value, "unlimited") > 0 Or _
       InStr(1, value, "uncap") > 0 Or _
       InStr(1, value, "no cap") > 0 Or _
       InStr(1, value, "no limitation") > 0 Then
        GetLiabilityType = "Uncapped"
        Exit Function
    End If
    
    ' Split multiple values (separated by ||)
    Dim values() As String
    values = Split(value, "||")
    
    Dim v As Variant
    Dim maxAmount As Double
    maxAmount = 0
    
    For Each v In values
        v = Trim(v)
        
        ' Extract numeric value and convert to USD
        Dim amount As Double
        amount = ExtractAmount(CStr(v))
        
        ' Track the maximum amount found
        If amount > maxAmount Then
            maxAmount = amount
        End If
    Next v
    
    ' Categorize based on amount
    If maxAmount >= 30000000 Then
        GetLiabilityType = "Over30M"
    ElseIf maxAmount >= 20000000 Then
        GetLiabilityType = "20M"
    ElseIf maxAmount >= 15000000 Then
        GetLiabilityType = "15M"
    ElseIf maxAmount >= 10000000 Then
        GetLiabilityType = "10M"
    Else
        GetLiabilityType = "Below10M"
    End If
End Function

Private Function ExtractAmount(ByVal value As String) As Double
    Dim numStr As String
    Dim i As Long
    Dim char As String
    Dim multiplier As Double
    
    ' Default multiplier (for USD)
    multiplier = 1
    
    ' Set currency multipliers (approximate exchange rates)
    If InStr(1, value, "£") > 0 Then
        multiplier = 1.27  ' GBP to USD
    ElseIf InStr(1, value, "€") > 0 Then
        multiplier = 1.09  ' EUR to USD
    End If
    
    ' Extract numeric value
    numStr = ""
    For i = 1 To Len(value)
        char = Mid(value, i, 1)
        If IsNumeric(char) Or char = "." Then
            numStr = numStr & char
        End If
    Next i
    
    ' Convert to number if possible
    If IsNumeric(numStr) Then
        ExtractAmount = CDbl(numStr) * multiplier
    Else
        ExtractAmount = 0
    End If
End Function

Public Sub FilterLiabilityCaps()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim liabilityType As String
    Dim uncapText As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    ' Enable status bar and show initial message
    Application.StatusBar = "Analyzing liability caps..."
    
    ' Clear existing values in columns O through T
    ws.Range("O5:T" & lastRow).ClearContents
    
    ' Set headers
    ws.Cells(4, "O").value = "$10M Cap"
    ws.Cells(4, "P").value = "$15M Cap"
    ws.Cells(4, "Q").value = "$20M Cap"
    ws.Cells(4, "R").value = "Over $30M"
    ws.Cells(4, "S").value = "Uncapped"
    ws.Cells(4, "T").value = "Uncapped Text"
    
    ' Loop through each row
    For i = 5 To lastRow
        ' Check column K for financial caps
        liabilityType = GetLiabilityType(ws.Cells(i, "K").value)
        
        ' Mark the appropriate column
        Select Case liabilityType
            Case "10M"
                ws.Cells(i, "O").value = "Yes"
            Case "15M"
                ws.Cells(i, "P").value = "Yes"
            Case "20M"
                ws.Cells(i, "Q").value = "Yes"
            Case "Over30M"
                ws.Cells(i, "R").value = "Yes"
        End Select
        
        ' Check column J for uncapped liabilities
        uncapText = ExtractUncappedText(ws.Cells(i, "J").value)
        If uncapText <> "" Then
            ws.Cells(i, "S").value = "Yes"
            ws.Cells(i, "T").value = uncapText
        End If
        
        ' Update status bar every 100 rows
        If i Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & i & " of " & lastRow & " (" & Format((i / lastRow) * 100, "0.0") & "%)"
            DoEvents
        End If
    Next i
    
    ' Format headers
    With ws.Range("O4:T4")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Format columns
    Dim formatRange As Range
    For i = 15 To 19 ' Columns O to S (15 to 19)
        Set formatRange = ws.Range(ws.Cells(5, i), ws.Cells(lastRow, i))
        With formatRange
            .Font.Bold = True
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                
                ' Different color for each column
                Select Case i
                    Case 15 ' Column O - $10M
                        .Color = RGB(255, 255, 153) ' Light yellow
                    Case 16 ' Column P - $15M
                        .Color = RGB(255, 204, 153) ' Light orange
                    Case 17 ' Column Q - $20M
                        .Color = RGB(255, 153, 153) ' Light red
                    Case 18 ' Column R - Over $30M
                        .Color = RGB(204, 153, 255) ' Light purple
                    Case 19 ' Column S - Uncapped
                        .Color = RGB(255, 102, 102) ' Darker red
                End Select
            End With
        End With
    Next i
    
    ' Format column T
    With ws.Range(ws.Cells(5, 20), ws.Cells(lastRow, 20))
        .Font.Bold = True
        .Font.Italic = True
    End With
    
    ' Reset status bar
    Application.StatusBar = False
    
    MsgBox "Analysis complete. Liability caps have been categorized in columns O through T.", vbInformation
End Sub

' Function to extract the uncapped text from a cell
Private Function ExtractUncappedText(ByVal cellValue As String) As String
    Dim value As String
    Dim result As String
    Dim i As Long
    Dim startPos As Long
    Dim uncapTerms As Variant
    
    If Trim(cellValue) = "" Then
        ExtractUncappedText = ""
        Exit Function
    End If
    
    ' Define uncapped liability terms
    uncapTerms = Array("uncapped", "unlimited", "uncap", "no cap", "no limitation")
    
    value = LCase(cellValue)
    result = ""
    
    ' Check for each term
    For i = LBound(uncapTerms) To UBound(uncapTerms)
        startPos = InStr(1, value, uncapTerms(i))
        If startPos > 0 Then
            ' Extract a reasonable context (20 characters before and after)
            Dim beforeContext As Long
            Dim afterContext As Long
            Dim termLength As Long
            
            termLength = Len(uncapTerms(i))
            beforeContext = Application.Max(1, startPos - 20)
            afterContext = Application.Min(Len(value), startPos + termLength + 20)
            
            ' Get the context
            result = Mid(cellValue, beforeContext, afterContext - beforeContext + 1)
            
            ' Clean up - Trim to sentence boundaries if possible
            ' Find the start of the phrase
            Dim periodBefore As Long
            periodBefore = InStrRev(Left(result, startPos - beforeContext), ".")
            If periodBefore > 0 Then
                result = Mid(result, periodBefore + 1)
            End If
            
            ' Find the end of the phrase
            Dim periodAfter As Long
            periodAfter = InStr(startPos - beforeContext + termLength, result, ".")
            If periodAfter > 0 Then
                result = Left(result, periodAfter)
            End If
            
            ' Trim spaces
            result = Application.Trim(result)
            Exit For
        End If
    Next i
    
    ExtractUncappedText = result
End Function

