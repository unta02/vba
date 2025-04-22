# Word Contract Analyzer

This VBA tool analyzes Microsoft Word documents to identify and highlight key contract clauses using Google's Gemini AI model. 

## Features

The analyzer automatically identifies, highlights, and adds comments to three critical types of contract clauses:

1. **Payment Terms** (highlighted in **blue**)
   - Standard payment periods (e.g., 30, 60, 90 days)
   - Invoice requirements
   - Late payment penalties

2. **Limitation of Liability** (highlighted in **red**)
   - Liability caps
   - Excluded damages
   - Uncapped liability scenarios

3. **Termination Clauses** (highlighted in **green**)
   - Notice periods
   - Termination for convenience provisions
   - Termination fees and calculations

## Installation

### Prerequisites
- Microsoft Word (Office 365, 2019, 2016)
- Internet connection for API access
- Google Gemini API key (provided in the code)

### Step 1: Import VBA Modules
1. Open Microsoft Word
2. Press `Alt + F11` to open the VBA Editor
3. Right-click on "Project" in the Project Explorer and select "Import File"
4. Import both modules:
   - `WordContractAnalyzer.bas`
   - `WordContractAnalyzerSetup.bas`

### Step 2: Add Required References
1. In the VBA Editor, go to `Tools > References`
2. Check the following references:
   - Microsoft Scripting Runtime
   - Microsoft XML, v6.0 (or your available version) 
   - Microsoft Word Object Library

### Step 3: Set Up the Tool
1. In the VBA Editor, double-click on `WordContractAnalyzerSetup`
2. Run the `ShowSetupInstructions` function
3. Follow the on-screen instructions to add the analyzer to Word's menu

## Usage

1. Open a contract document in Word
2. Click the "Contract Analyzer" menu item
3. Wait for the analysis to complete (a progress bar will be displayed)
4. Review the highlighted clauses and their associated comments:
   - **Blue highlights**: Payment terms
   - **Red highlights**: Limitation of liability clauses
   - **Green highlights**: Termination clauses

## Testing and Examples

The setup module includes helpful functions to test and demonstrate the tool:

- `TestApiConnection`: Verifies that the Gemini API connection is working properly
- `CreateSampleDocument`: Creates a sample contract document with example clauses
- `SampleUsage`: Displays usage information and offers to analyze a document

## How It Works

1. The tool extracts the full text of the active Word document
2. It sends the text to Google's Gemini AI model with a specialized prompt
3. Gemini analyzes the text to identify the three types of contract clauses
4. The tool processes Gemini's response and highlights the identified clauses in the document
5. Comments are added with analysis of each clause

## Customization

You can modify the analysis criteria by editing the `GetAnalysisPrompt` function in the `WordContractAnalyzer.bas` file.

## Troubleshooting

- **Reference Errors**: Make sure you've added all the required references
- **API Errors**: Run the `TestApiConnection` function to verify connectivity
- **Menu Not Appearing**: Run the `AddMenuItemToWord` function again
- **Analysis Failures**: Check your internet connection and try again

## Limitations

- Very long documents may encounter limitations due to API constraints
- The accuracy of clause identification depends on the Gemini AI model
- Some highly specialized legal language may not be correctly identified

## Security Note

This tool sends contract text to Google's Gemini API for analysis. While the API connection is secure, please ensure you have appropriate authorization to share contract data externally.

## License

This code is provided for educational and demonstration purposes.

## Acknowledgments

- Uses Google's Gemini AI for natural language processing
- Based on VBA integration examples for Microsoft Office 