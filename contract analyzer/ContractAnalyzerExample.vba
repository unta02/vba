' Example implementation for Contract Clause Analyzer
' This code demonstrates how to use the analyzer in a real Excel workbook

Sub SetupContractAnalyzerWorksheet()
    ' Create a worksheet for contract analysis if it doesn't exist
    Dim ws As Worksheet
    Dim wsExists As Boolean
    
    ' Check if the worksheet exists
    wsExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Contract Analysis" Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    ' Create the worksheet if it doesn't exist
    If Not wsExists Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Contract Analysis"
    Else
        Set ws = ThisWorkbook.Worksheets("Contract Analysis")
    End If
    
    ' Clear the worksheet
    ws.Cells.Clear
    
    ' Set up headers
    ws.Cells(1, 1).Value = "Contract Clause"
    ws.Cells(1, 2).Value = "Analysis Result"
    
    ' Format headers
    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Adjust column widths
    ws.Columns("A").ColumnWidth = 60
    ws.Columns("B").ColumnWidth = 40
    
    ' Add example data
    ws.Cells(2, 1).Value = "In no event shall either party be liable for any indirect, incidental, special, punitive, or consequential damages. Direct damages shall be limited to the total amount paid under this agreement."
    ws.Cells(3, 1).Value = "Each party shall be liable for all damages arising from their breach of this agreement, including but not limited to direct, indirect, consequential, and incidental damages."
    ws.Cells(4, 1).Value = "Liability for direct damages shall be capped at the amount paid in the last 12 months. Neither party shall be liable for any indirect, special, or consequential damages."
    
    ' Add a button to run the analysis
    Dim btn As Button
    On Error Resume Next
    ActiveSheet.Buttons.Delete
    Set btn = ActiveSheet.Buttons.Add(400, 5, 100, 25)
    
    With btn
        .Caption = "Analyze Clauses"
        .Name = "btnAnalyze"
        .OnAction = "AnalyzeContractClauses"
    End With
    
    ' Show a message
    MsgBox "Contract Analysis worksheet has been set up." & vbCrLf & _
           "1. Add your contract clauses in Column A" & vbCrLf & _
           "2. Click the 'Analyze Clauses' button to start the analysis", vbInformation
End Sub

Sub CreateContractAnalyzerRibbon()
    ' This is a placeholder for creating a custom ribbon
    ' Actual implementation requires XML customization
    MsgBox "To add this tool to the Excel ribbon, you would need to:" & vbCrLf & _
           "1. Create a Custom UI Editor for Office" & vbCrLf & _
           "2. Add XML code for the ribbon customization" & vbCrLf & _
           "3. Assign callbacks to the ribbon buttons", vbInformation
End Sub

Sub TestGeminiApiConnection()
    ' Test the Gemini API connection
    Dim result As String
    Dim testText As String
    
    ' Simple test text
    testText = "The vendor shall be liable for all damages arising from breach of contract."
    
    ' Call the API function from the main module
    result = AnalyzeWithGemini(testText)
    
    ' Display the result
    MsgBox "API Test Result:" & vbCrLf & vbCrLf & result, vbInformation
End Sub

Sub ExportResultsToReport()
    ' Example of exporting the analysis results to a report
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim foundCount As Long
    
    ' Get the worksheet
    Set ws = ThisWorkbook.Worksheets("Contract Analysis")
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range for the report
    Set rng = ws.Range("A1:B" & lastRow)
    
    ' Count uncapped liability instances
    foundCount = 0
    Dim i As Long
    For i = 2 To lastRow
        If InStr(1, ws.Cells(i, 2).Value, "UNCAPPED LIABILITY FOUND", vbTextCompare) > 0 Then
            foundCount = foundCount + 1
        End If
    Next i
    
    ' Create a simple report
    Dim reportWs As Worksheet
    
    ' Check if report worksheet exists and create it if needed
    On Error Resume Next
    Set reportWs = ThisWorkbook.Worksheets("Analysis Report")
    On Error GoTo 0
    
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Worksheets.Add(After:=ws)
        reportWs.Name = "Analysis Report"
    End If
    
    ' Clear the report worksheet
    reportWs.Cells.Clear
    
    ' Add report title
    reportWs.Cells(1, 1).Value = "Contract Liability Analysis Report"
    reportWs.Cells(1, 1).Font.Bold = True
    reportWs.Cells(1, 1).Font.Size = 14
    
    ' Add summary information
    reportWs.Cells(3, 1).Value = "Total clauses analyzed:"
    reportWs.Cells(3, 2).Value = lastRow - 1
    
    reportWs.Cells(4, 1).Value = "Clauses with uncapped liability:"
    reportWs.Cells(4, 2).Value = foundCount
    
    reportWs.Cells(5, 1).Value = "Percentage with uncapped liability:"
    If lastRow > 1 Then
        reportWs.Cells(5, 2).Value = Format(foundCount / (lastRow - 1), "0.0%")
    Else
        reportWs.Cells(5, 2).Value = "N/A"
    End If
    
    ' Format the summary area
    reportWs.Range("A3:A5").Font.Bold = True
    
    ' Add detailed findings
    reportWs.Cells(7, 1).Value = "Detailed Findings:"
    reportWs.Cells(7, 1).Font.Bold = True
    
    reportWs.Cells(8, 1).Value = "Clause"
    reportWs.Cells(8, 2).Value = "Analysis Result"
    
    ' Format headers
    With reportWs.Range("A8:B8")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Copy analysis results to the report
    Dim rowOffset As Long
    rowOffset = 9
    Dim uncappedCount As Long
    uncappedCount = 0
    
    For i = 2 To lastRow
        If InStr(1, ws.Cells(i, 2).Value, "UNCAPPED LIABILITY FOUND", vbTextCompare) > 0 Then
            reportWs.Cells(rowOffset + uncappedCount, 1).Value = ws.Cells(i, 1).Value
            reportWs.Cells(rowOffset + uncappedCount, 2).Value = ws.Cells(i, 2).Value
            
            ' Highlight uncapped liability findings
            reportWs.Cells(rowOffset + uncappedCount, 2).Interior.Color = RGB(255, 200, 200)
            
            uncappedCount = uncappedCount + 1
        End If
    Next i
    
    ' Adjust column widths
    reportWs.Columns("A").ColumnWidth = 60
    reportWs.Columns("B").ColumnWidth = 40
    
    ' Show the report
    reportWs.Activate
    MsgBox "Report generated successfully!", vbInformation
End Sub 