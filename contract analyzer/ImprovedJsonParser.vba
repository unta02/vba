' Improved JSON parser for Gemini API responses
' This can replace the simplified parser in the main module

Private Function ImprovedParseJson(jsonString As String) As Object
    ' Create a Dictionary to hold the parsed JSON
    Dim jsonObject As Object
    Set jsonObject = CreateObject("Scripting.Dictionary")
    
    ' Skip if empty
    If Trim(jsonString) = "" Then
        Set ImprovedParseJson = jsonObject
        Exit Function
    End If
    
    ' For Gemini API specifically, we're looking for the text field in the response
    
    ' First, let's clean up the string for easier parsing
    Dim cleanJson As String
    cleanJson = jsonString
    
    ' Extract the text content using regex-like string operations
    Dim candidates As Object
    Set candidates = CreateObject("Scripting.Dictionary")
    
    ' Look for the text content in the response
    Dim textContent As String
    textContent = ExtractTextFromGeminiResponse(jsonString)
    
    ' Create the structure similar to what we'd expect from a JSON parser
    If textContent <> "" Then
        ' Create nested structure
        Dim content As Object
        Set content = CreateObject("Scripting.Dictionary")
        
        Dim parts As Object
        Set parts = CreateObject("Scripting.Dictionary")
        
        ' Build the structure we expect
        Dim textPart As Object
        Set textPart = CreateObject("Scripting.Dictionary")
        textPart.Add "text", textContent
        
        parts.Add 1, textPart
        content.Add "parts", parts
        
        Dim candidate As Object
        Set candidate = CreateObject("Scripting.Dictionary")
        candidate.Add "content", content
        
        candidates.Add 1, candidate
        jsonObject.Add "candidates", candidates
    End If
    
    Set ImprovedParseJson = jsonObject
End Function

Private Function ExtractTextFromGeminiResponse(jsonString As String) As String
    Dim result As String
    result = ""
    
    ' Patterns for Gemini API response (these will need to be adjusted based on actual responses)
    Dim patterns(3) As String
    patterns(0) = """text"":""" ' Basic pattern
    patterns(1) = """text"": """ ' With space
    patterns(2) = "{""text"":""" ' In object
    patterns(3) = """parts"":[{""text"":""" ' In parts array
    
    Dim i As Integer
    Dim startPos As Long
    Dim endPos As Long
    
    ' Try each pattern
    For i = 0 To UBound(patterns)
        startPos = InStr(1, jsonString, patterns(i))
        If startPos > 0 Then
            ' Found a match
            startPos = startPos + Len(patterns(i))
            
            ' Look for end of text (accounting for escaped quotes)
            endPos = startPos
            Dim escapeCount As Integer
            
            Do
                endPos = InStr(endPos + 1, jsonString, """")
                If endPos = 0 Then Exit Do
                
                ' Count backslashes before quote
                escapeCount = 0
                Dim j As Long
                j = endPos - 1
                Do While j > 0 And Mid(jsonString, j, 1) = "\"
                    escapeCount = escapeCount + 1
                    j = j - 1
                Loop
                
                ' If even number of backslashes, it's not escaped
                If escapeCount Mod 2 = 0 Then Exit Do
            Loop
            
            If endPos > startPos Then
                result = Mid(jsonString, startPos, endPos - startPos)
                Exit For
            End If
        End If
    Next i
    
    ' Unescape the JSON string
    If result <> "" Then
        result = UnescapeJsonString(result)
    End If
    
    ExtractTextFromGeminiResponse = result
End Function

Private Function UnescapeJsonString(jsonStr As String) As String
    Dim result As String
    result = jsonStr
    
    ' Handle common JSON escape sequences
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    result = Replace(result, "\/", "/")
    result = Replace(result, "\b", Chr(8))
    result = Replace(result, "\f", Chr(12))
    result = Replace(result, "\n", vbLf)
    result = Replace(result, "\r", vbCr)
    result = Replace(result, "\t", vbTab)
    
    ' Handle Unicode escapes \uXXXX
    Dim reUnicode As Object
    Dim matches As Object
    
    ' This is a simple version - for full Unicode support
    ' you would need more comprehensive handling
    Dim i As Long
    Dim pos As Long
    Dim hex As String
    
    pos = InStr(1, result, "\u")
    Do While pos > 0
        If pos + 5 <= Len(result) Then
            hex = Mid(result, pos + 2, 4)
            If IsHex(hex) Then
                ' Convert hex to character
                Dim charCode As Long
                charCode = CLng("&H" & hex)
                
                ' Replace the escape sequence with the character
                result = Left(result, pos - 1) & ChrW(charCode) & Mid(result, pos + 6)
            End If
        End If
        
        ' Find next occurrence
        pos = InStr(pos + 1, result, "\u")
    Loop
    
    UnescapeJsonString = result
End Function

Private Function IsHex(text As String) As Boolean
    ' Check if a string is a valid hexadecimal number
    Dim i As Long
    Dim c As String
    
    If Len(text) = 0 Then
        IsHex = False
        Exit Function
    End If
    
    For i = 1 To Len(text)
        c = UCase(Mid(text, i, 1))
        If InStr("0123456789ABCDEF", c) = 0 Then
            IsHex = False
            Exit Function
        End If
    Next i
    
    IsHex = True
End Function

' Example usage:
' 1. Replace ParseJson with ImprovedParseJson in your AnalyzeWithGemini function
' 2. For more robust handling, update the response extraction as follows:

Private Function ExtractGeminiResult(jsonResponse As Object) As String
    Dim result As String
    result = ""
    
    On Error Resume Next
    
    ' Try multiple possible paths in the response structure
    If Not jsonResponse Is Nothing Then
        ' Path 1: Standard response format
        If Not jsonResponse("candidates") Is Nothing Then
            If Not jsonResponse("candidates")(1) Is Nothing Then
                If Not jsonResponse("candidates")(1)("content") Is Nothing Then
                    If Not jsonResponse("candidates")(1)("content")("parts") Is Nothing Then
                        If Not jsonResponse("candidates")(1)("content")("parts")(1) Is Nothing Then
                            result = jsonResponse("candidates")(1)("content")("parts")(1)("text")
                        End If
                    End If
                End If
            End If
        End If
        
        ' If standard path didn't work, try alternatives
        If result = "" Then
            ' Path 2: Simplified format
            If Not jsonResponse("candidates") Is Nothing Then
                If Not jsonResponse("candidates")(1) Is Nothing Then
                    If Not jsonResponse("candidates")(1)("content") Is Nothing Then
                        result = jsonResponse("candidates")(1)("content")("text")
                    End If
                End If
            End If
        End If
        
        ' Path 3: Direct text field
        If result = "" Then
            If Not jsonResponse("text") Is Nothing Then
                result = jsonResponse("text")
            End If
        End If
    End If
    
    On Error GoTo 0
    
    ExtractGeminiResult = result
End Function 