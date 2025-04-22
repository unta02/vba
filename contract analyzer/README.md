# Contract Clause Analyzer

This VBA tool analyzes contract clauses to identify uncapped liability in limitation of liability sections using Google's Gemini Flash 2.0 API.

## Setup Instructions

### Prerequisites
- Microsoft Excel
- References required in VBA:
  - Microsoft Scripting Runtime
  - Microsoft XML, v6.0 (or available version)

### Installation Steps

1. Open your Excel workbook
2. Press Alt + F11 to open the VBA editor
3. In the VBA Project Explorer, right-click on "VBAProject"
4. Select Insert → Module
5. Copy the entire code from the `ContractClauseAnalyzer.bas` file and paste it into the new module
6. Save the workbook as a macro-enabled workbook (.xlsm)

### Setting up VBA References

1. In the VBA editor, go to Tools → References
2. Check the boxes for:
   - Microsoft Scripting Runtime
   - Microsoft XML, v6.0 (or your available version)
3. Click OK

## Usage

### Preparing Your Data
1. Place your contract clauses in Column A of your worksheet
2. (Optional) Add a header in cell A1 (e.g., "Contract Clause")
3. (Optional) Add a header in cell B1 (e.g., "Analysis Result")

### Testing the API Connection
Before analyzing a full dataset, it's recommended to test the API connection:

1. In the VBA editor, run the `TestApiConnection` procedure
2. A form will appear with a sample contract clause
3. Click "Test API" to verify that the connection works
4. The API response will be displayed in the results box

### Running the Tool

#### Method 1: Using the UI
1. In the VBA editor, run the `ShowContractAnalyzerUI` procedure
2. A simple form will appear
3. Click "Start Analysis" to begin processing

#### Method 2: Direct Execution
1. In the VBA editor, run the `AnalyzeContractClauses` procedure
2. Alternatively, you can create a button on your worksheet and assign the `AnalyzeContractClauses` macro to it

### Results

The tool will:
1. Process each clause in column A
2. Call the Gemini API to analyze the text
3. Output the results in column B
4. Display a completion message when finished

## API Information

This tool uses the Google Gemini Flash 2.0 API. The API key is embedded in the code but can be changed if needed.

### Understanding API Responses

The API will respond with one of three possible formats:
1. `UNCAPPED LIABILITY FOUND: [brief explanation]` - When uncapped liability is detected
2. `No uncapped liability found: [brief explanation]` - When all liability is properly capped
3. `UNCERTAIN: [explanation]` - When the model cannot determine with confidence

## Troubleshooting

If you encounter errors:

### API Connection Issues
1. Use the `TestApiConnection` procedure to verify if the API is working
2. Check the API URL in the code (currently set to "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent")
3. Verify that the API key is correct and has proper permissions
4. Check your internet connection

### JSON Parsing Issues
If you see "Error parsing JSON response" messages:
1. Run the `TestJsonExtraction` procedure to verify the JSON parser is working
2. Check the sample JSON in the code against the actual response format
3. The code has been updated to handle the standard Gemini API response format

### Other Issues
1. Verify that the required references are set up correctly
2. For large contract texts, consider splitting them into smaller sections
3. Try analyzing one clause at a time using the TestApiConnection form

## Customization

You can modify the JSON request in the `AnalyzeWithGemini` function to customize the prompt sent to the Gemini API based on your specific needs.

### Key Parameters to Modify:
- The prompt text that instructs the model
- The response format instructions
- The safety settings

## Updates

### Latest Updates (Version 1.1)
- Fixed JSON parsing to properly handle Gemini API responses
- Added TestApiConnection functionality for easier troubleshooting
- Improved prompt engineering for more consistent responses
- Added TestJsonExtraction feature to verify JSON parsing 